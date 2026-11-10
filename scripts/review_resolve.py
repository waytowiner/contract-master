"""复核结果回写数据库，形成 verified 闭环。

用法：
  python review_resolve.py --db <contracts.db> --queue <04_review_queue.xlsx>

规则：
- 优先按「复核项ID」精确关闭一行；无 ID 的旧队列仅在同字段恰好一条 pending 时兼容
- 「人工修正值」非空 -> 用修正值更新合同字段并标记 resolved；金额必须是有限数值
- 「人工修正值」为空但「状态」为已确认（自动去首尾空格） -> 保留原值，仅标记 resolved
- 无法识别的状态即使有修正值也保持 pending；其余未确认行不影响数据库
- 更新 title/party_a/party_b 后同步 contract_fts；原始 source_value 与物理页码保留
- 复核、开放问题和字段待复核标记全部清零后，合同才置为 verified
"""
import argparse
import math
import sqlite3
from pathlib import Path

from openpyxl import load_workbook

FIELD_COLUMNS = {
    "contract_no", "title", "party_a", "party_b", "amount",
    "amount_in_words", "sign_place", "bank_name", "bank_account",
}
FTS_COLUMNS = {"title", "party_a", "party_b"}
KNOWN_STATUS = {"待确认", "已确认"}


def parse_amount(value):
    """把人工输入转为有限数值；无法可靠转换时返回 None，绝不静默确认。"""
    try:
        number = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", required=True)
    ap.add_argument("--queue", required=True)
    ap.add_argument("--reviewer", default="人工复核")
    args = ap.parse_args()

    conn = sqlite3.connect(args.db)
    cur = conn.cursor()
    wb = load_workbook(args.queue)
    ws = wb.active

    header = [c.value for c in ws[1]]
    idx = {name: header.index(name) for name in
           ["合同组", "字段/问题", "当前识别值", "状态", "人工修正值"]}
    root_idx = header.index("来源根目录") if "来源根目录" in header else None
    review_key_idx = header.index("复核项ID") if "复核项ID" in header else None

    resolved = pending = skipped = ambiguous = invalid_amount = 0
    bad_status = {}  # 状态原值 -> 出现次数（无法识别的值，绝不静默吞掉）
    touched_contracts = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        group, field, current, status, fixed = (row[idx["合同组"]], row[idx["字段/问题"]],
                                                row[idx["当前识别值"]], row[idx["状态"]],
                                                row[idx["人工修正值"]])
        if group is None:
            continue
        status = str(status).strip() if status is not None else ""
        root_value = row[root_idx] if root_idx is not None else None
        review_key = str(row[review_key_idx]).strip() if review_key_idx is not None and row[review_key_idx] else ""
        if root_value:
            hits = cur.execute("SELECT id FROM contracts WHERE group_name=? AND source_root=?",
                               (str(group), str(root_value))).fetchall()
        else:
            hits = cur.execute("SELECT id FROM contracts WHERE group_name=?", (str(group),)).fetchall()
        if not hits:
            skipped += 1
            continue
        if len(hits) > 1:
            ambiguous += 1
            continue
        cid = hits[0][0]

        if status and status not in KNOWN_STATUS:
            bad_status[status] = bad_status.get(status, 0) + 1
            pending += 1
            continue

        has_fix = fixed is not None and str(fixed).strip() != ""
        confirmed = (status == "已确认") or has_fix
        if not confirmed:
            pending += 1
            continue

        # 新版队列按 review_key 精确关闭一行。旧队列仅在同字段恰好一条 pending 时兼容，
        # 多条同字段绝不能按组+字段批量确认。
        if review_key:
            pending_rows = cur.execute(
                "SELECT id FROM review_items WHERE review_key=? AND contract_id=? AND status='pending'",
                (review_key, cid)).fetchall()
        else:
            pending_rows = cur.execute(
                "SELECT id FROM review_items WHERE contract_id=? AND field_name=? AND status='pending'",
                (cid, field)).fetchall()
        if not pending_rows:
            skipped += 1
            continue
        if len(pending_rows) != 1:
            ambiguous += 1
            continue
        pending_id = pending_rows[0][0]

        final_value = str(fixed).strip() if has_fix else ("" if current is None else str(current))
        if field in FIELD_COLUMNS:
            # 队列的“当前识别值”是导出时的快照。同字段另一条复核项可能已修正
            # 数据库中的值；仅确认本项时必须保留最新值，不能用旧快照还原修正。
            if not has_fix:
                stored = cur.execute(f"SELECT {field} FROM contracts WHERE id=?", (cid,)).fetchone()[0]
                final_value = "" if stored is None else str(stored)
            col_value = final_value
            if field == "amount":
                if final_value == "" and not has_fix:
                    col_value = None
                else:
                    col_value = parse_amount(final_value)
                    if col_value is None:
                        invalid_amount += 1
                        pending += 1
                        continue
            # field 已被 FIELD_COLUMNS 白名单约束，不构成 SQL 注入面
            cur.execute(f"UPDATE contracts SET {field}=? WHERE id=?", (col_value, cid))
            if field in FTS_COLUMNS:
                cur.execute(f"UPDATE contract_fts SET {field}=? WHERE rowid=?", (col_value, cid))
            cur.execute(
                """INSERT INTO field_sources(contract_id, field_name, value, source_value,
                   file_id, page_no, source_page_no, snippet, confidence, method,
                   needs_review, verified_by, verified_at)
                   VALUES(?,?,?,NULL,NULL,1,NULL,'人工复核回写',1.0,'manual_review',0,?,datetime('now','localtime'))
                   ON CONFLICT(contract_id, field_name) DO UPDATE SET
                     value=excluded.value, confidence=1.0, method='manual_review', needs_review=0,
                     snippet=CASE WHEN field_sources.snippet IS NULL OR field_sources.snippet=''
                                  THEN excluded.snippet ELSE field_sources.snippet END,
                     verified_by=excluded.verified_by, verified_at=excluded.verified_at""",
                (cid, field, final_value, args.reviewer))
        cur.execute(
            """UPDATE review_items SET status='resolved', resolved_value=?, resolved_by=?,
               resolved_at=datetime('now','localtime')
               WHERE id=? AND status='pending'""",
            (final_value, args.reviewer, pending_id))
        resolved += cur.rowcount
        if field not in FIELD_COLUMNS:
            cur.execute("UPDATE issues SET status='resolved' WHERE contract_id=? AND issue_type=? AND status='open'",
                        (cid, field))
        touched_contracts.add(cid)

    for cid in touched_contracts:
        cur.execute("UPDATE contracts SET status='unverified' WHERE id=?", (cid,))
        cur.execute(
            """UPDATE contracts SET status='verified'
               WHERE id=?
                 AND NOT EXISTS (SELECT 1 FROM review_items WHERE contract_id=? AND status='pending')
                 AND NOT EXISTS (SELECT 1 FROM issues WHERE contract_id=? AND status='open')
                 AND NOT EXISTS (SELECT 1 FROM field_sources WHERE contract_id=? AND needs_review=1)""",
            (cid, cid, cid, cid))
    conn.commit()

    if bad_status:
        total_bad = sum(bad_status.values())
        print(f"警告：{total_bad} 行的状态值无法识别（只认“待确认”/“已确认”），已按待确认处理，未回写：")
        for val, n in sorted(bad_status.items(), key=lambda kv: -kv[1]):
            print(f"  - {val!r} x{n}")

    if invalid_amount:
        print(f"警告：{invalid_amount} 行金额不是有限数值，已保持待复核，未回写")
    if ambiguous:
        print(f"警告：{ambiguous} 行复核项定位歧义（跨根目录重名或无 ID 的同字段多项），未回写")

    total = cur.execute("SELECT COUNT(*) FROM contracts").fetchone()[0]
    verified = cur.execute("SELECT COUNT(*) FROM contracts WHERE status='verified'").fetchone()[0]
    remain = cur.execute("SELECT COUNT(*) FROM review_items WHERE status='pending'").fetchone()[0]
    print(f"回写完成：resolved {resolved} | 仍待复核 {pending} | 未匹配组 {skipped}")
    print(f"合同 {total} 份，其中 verified {verified} 份，剩余待复核项 {remain}")
    conn.close()


if __name__ == "__main__":
    main()
