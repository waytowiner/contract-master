"""技能自检：对照 SKILL.md 验收标准逐项实测，输出通过/不通过与证据。

用法：python verify_skill.py --out <输出目录> [--install <安装副本目录>]

统计口径：PASS/FAIL 计入总数；SKIP（如未提供 --install）不计入通过数。
"""
import argparse
import hashlib
import json
import os
import random
import re
import sqlite3
import subprocess
import sys
import tempfile
from pathlib import Path

REQUIRED = [
    "01_inventory.xlsx", "ocr", "04_review_queue.xlsx",
    "05_contract_ledger.xlsx", "05_contract_ledger.csv",
    "06_contracts.db", "07_report.md", "run_log.json",
]
def sha(p: Path):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def check(name, ok, detail="", skipped=False):
    return {"check": name, "pass": bool(ok), "detail": detail, "skipped": skipped}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--install", default=None,
                    help="安装副本目录；不提供则跳过源码与安装副本一致性检查")
    ap.add_argument("--sample", type=int, default=10)
    args = ap.parse_args()

    out = Path(args.out).expanduser().resolve()
    src_root = Path(__file__).resolve().parents[1]
    if out == src_root or src_root in out.parents:
        raise SystemExit("错误：--out 不能指向技能目录或其子目录")
    inv_path = out / "inventory.json"
    if inv_path.exists():
        inv = json.loads(inv_path.read_text(encoding="utf-8"))
        if inv.get("root"):
            source_root = Path(inv["root"]).expanduser().resolve()
            if out == source_root or source_root in out.parents:
                raise SystemExit("错误：--out 不能位于合同输入目录内")
    results = []

    missing = [f for f in REQUIRED if not (out / f).exists()]
    results.append(check("交付物齐全（8 项）", not missing, f"缺失：{missing}" if missing else "全部存在"))

    if not args.install:
        results.append(check("安装副本一致性", False, "未提供 --install，跳过（不计入通过数）", skipped=True))
    else:
        inst = Path(args.install)
        drift = []
        sys.path.insert(0, str(src_root / "scripts"))
        from package_dist import iter_sources, should_sync
        pairs = [(p, inst / p.relative_to(src_root)) for p, _ in iter_sources()
                 if should_sync(p.relative_to(src_root))]
        for a, b in pairs:
            if not b.exists():
                drift.append(f"{b.relative_to(inst).as_posix()} 未安装")
            elif sha(a) != sha(b):
                drift.append(f"{b.relative_to(inst).as_posix()} 不一致")
        results.append(check("安装副本与源码一致（含 tests）", not drift, "；".join(drift) or "全部一致"))

    # 金额解析回归：直接调用被测函数断言（防角分/零占位等历史 bug 复发）
    try:
        sys.path.insert(0, str(src_root / "scripts"))
        from extract_fields import cn_amount
        regression = [
            ("壹佰元伍角", 100.5), ("壹佰元贰角叁分", 100.23), ("玖拾玖元零角玖分", 99.09),
            ("壹佰零伍元整", 105), ("贰拾捌万玖仟陆佰元整", 289600),
            ("贰拾捌万玖任陆佰元整", 289600), ("壹佰元叁X整", None),
        ]
        bad = [f"{t}->{cn_amount(t)}(期望{w})" for t, w in regression if cn_amount(t) != w]
        results.append(check("金额解析回归（角分/零/容错/残缺）", not bad,
                             "全部正确" if not bad else "；".join(bad)))
    except Exception as exc:
        results.append(check("金额解析回归（角分/零/容错/残缺）", False, f"无法导入：{exc}"))

    try:
        from extract_fields import sanitize_party
        party_cases = [("（盖章）", ""), ("公司", ""),
                       ("某科技有限公司（以下简称乙方)", "某科技有限公司"),
                       ("6、本合同传真签署与正本具有同等法律效力。", "")]
        bad_party = [f"{text!r}->{sanitize_party(text)[0]!r}" for text, want in party_cases
                     if sanitize_party(text)[0] != want]
        results.append(check("甲乙方值卫生回归", not bad_party,
                             "全部正确" if not bad_party else "；".join(bad_party)))
    except Exception as exc:
        results.append(check("甲乙方值卫生回归", False, f"无法导入：{exc}"))

    db = out / "06_contracts.db"
    conn = sqlite3.connect(db)
    cur = conn.cursor()
    records = json.loads((out / "extracted.json").read_text(encoding="utf-8"))
    bad_checks = []
    for r in records:
        passed = r.get("amount_check") == "大小写金额一致"
        candidates = [c.get("value") for c in r.get("amount_candidates", [])]
        if passed and (r.get("cn_amount") is None or
                       not any(v is not None and abs(v - r["cn_amount"]) < 0.01 for v in candidates)):
            bad_checks.append(f"{r.get('group')}：无双候选证据却标一致")
        if passed and r.get("amount_check_detail"):
            bad_checks.append(f"{r.get('group')}：一致与异常状态并存")
        candidates_with_priority = [c for c in r.get("amount_candidates", []) if c.get("priority") is not None]
        if candidates_with_priority:
            top_priority = min(c["priority"] for c in candidates_with_priority)
            top_values = {c.get("value") for c in candidates_with_priority if c["priority"] == top_priority}
        else:
            top_values = {c.get("value") for c in r.get("amount_candidates", [])
                          if c.get("method") == r.get("numeric_method")}
        if passed and len(top_values) > 1:
            bad_checks.append(f"{r.get('group')}：同优先级小写候选冲突却标一致")
    results.append(check("金额勾稽状态可信", not bad_checks,
                         "全部一致状态均有仲裁前双候选证据" if not bad_checks else "；".join(bad_checks[:5])))
    multi_page_groups = sum(1 for r in records if len(r.get("pages", [])) > 1)
    n_contracts = cur.execute("SELECT COUNT(*) FROM contracts").fetchone()[0]
    n_sources = cur.execute("SELECT COUNT(*) FROM field_sources").fetchone()[0]
    n_pending = cur.execute("SELECT COUNT(*) FROM review_items WHERE status='pending'").fetchone()[0]
    n_snippet = cur.execute(
        "SELECT COUNT(*) FROM field_sources WHERE snippet IS NOT NULL AND snippet!=''").fetchone()[0]
    n_page1 = cur.execute("SELECT COUNT(*) FROM field_sources WHERE page_no=1").fetchone()[0]
    n_page_valid = cur.execute(
        "SELECT COUNT(*) FROM field_sources WHERE page_no IS NOT NULL AND page_no>0").fetchone()[0]
    source_errors = []
    try:
        source_rows = cur.execute(
            """SELECT fs.id, fs.method, fs.file_id, fs.source_page_no, fs.snippet,
                      fs.verified_by, f.path, op.text
               FROM field_sources fs
               LEFT JOIN files f ON f.id=fs.file_id
               LEFT JOIN ocr_pages op ON op.file_id=fs.file_id AND op.page_no=fs.source_page_no"""
        ).fetchall()
        for sid, method, file_id, physical_page, snippet, reviewer, path, page_text in source_rows:
            if method == "manual_review":
                if not reviewer:
                    source_errors.append(f"{sid}:人工溯源缺复核人")
                if file_id is not None and page_text is None:
                    source_errors.append(f"{sid}:人工复核关联的原始页不存在")
            elif not file_id or not path or not physical_page or page_text is None:
                source_errors.append(f"{sid}:规则溯源缺真实文件或物理页")
            if page_text and snippet and snippet != "人工复核回写":
                if re.sub(r"\s", "", snippet) not in re.sub(r"\s", "", page_text):
                    source_errors.append(f"{sid}:片段不在所关联物理页")
    except sqlite3.OperationalError as exc:
        source_errors.append(f"schema 缺 source_page_no 等必需列：{exc}")
    snippet_ok = n_snippet == n_sources and n_sources > 0 and not any("片段" in e for e in source_errors)
    results.append(check("字段溯源含原文片段", snippet_ok,
                         f"{n_snippet}/{n_sources} 条有片段；溯源错误 {source_errors[:5]}"))
    page_ok = n_page_valid == n_sources and n_sources > 0 and not source_errors
    results.append(check("字段溯源页码有效", page_ok,
                         f"有效展示页码 {n_page_valid}/{n_sources}；page_no=1 为 {n_page1} 条；"
                         f"多页组 {multi_page_groups} 个；物理页错误 {source_errors[:5]}"))

    # files 表质量：全量、真实 ext、sha256、绝对路径（可一键定位原件）
    n_files = cur.execute("SELECT COUNT(*) FROM files").fetchone()[0]
    n_sha = cur.execute("SELECT COUNT(*) FROM files WHERE sha256 IS NOT NULL AND sha256!=''").fetchone()[0]
    n_abs = cur.execute("SELECT COUNT(*) FROM files WHERE path LIKE '_:%' OR path LIKE '/%'").fetchone()[0]
    n_ext = cur.execute("SELECT COUNT(DISTINCT ext) FROM files").fetchone()[0]
    n_failed_files = cur.execute("SELECT COUNT(*) FROM files WHERE status='ocr_failed'").fetchone()[0]
    n_pending_files = cur.execute("SELECT COUNT(*) FROM files WHERE status='pending'").fetchone()[0]
    n_unknown_files = cur.execute(
        "SELECT COUNT(*) FROM files WHERE status NOT IN ('ocr_done','duplicate','ocr_failed','pending')").fetchone()[0]
    files_ok = (n_files >= n_contracts and n_contracts > 0 and n_sha == n_files and
                n_abs == n_files and n_pending_files == 0 and n_unknown_files == 0)
    results.append(check("files 表全量且真实（含 sha256/绝对路径）", files_ok,
                         f"files {n_files} 条（contracts {n_contracts}），sha256 {n_sha}，"
                         f"绝对路径 {n_abs}，ext 种类 {n_ext}，pending {n_pending_files}，未知状态 {n_unknown_files}"))

    n_ocr_pages = cur.execute("SELECT COUNT(*) FROM ocr_pages").fetchone()[0]
    bad_file_pages = cur.execute(
        """SELECT f.path, f.page_count, COUNT(op.id) AS actual
           FROM files f LEFT JOIN ocr_pages op ON op.file_id=f.id
           WHERE f.status IN ('ocr_done','duplicate')
           GROUP BY f.id
           HAVING f.page_count IS NULL OR f.page_count<=0 OR f.page_count!=actual"""
    ).fetchall()
    results.append(check("ocr_pages 已填充（库自包含全文）",
                         not bad_file_pages and n_pending_files == 0 and n_ocr_pages > 0,
                         f"ocr_pages {n_ocr_pages} 条 / files {n_files} 条（失败 {n_failed_files}）；"
                         f"逐文件页数异常 {bad_file_pages[:5]}"))

    fts_n = cur.execute("SELECT COUNT(*) FROM contract_fts").fetchone()[0]
    results.append(check("全文检索已填充", fts_n >= n_contracts and n_contracts > 0,
                         f"contract_fts 记录 {fts_n} 条，合同 {n_contracts} 份"))

    q1 = cur.execute("SELECT COUNT(*) FROM contracts WHERE amount>100000").fetchone()[0]
    q2 = cur.execute("SELECT COUNT(*) FROM contracts WHERE party_b LIKE '%公司%'").fetchone()[0]
    q3 = cur.execute("SELECT COUNT(*) FROM contracts WHERE contract_no IS NOT NULL AND contract_no!=''").fetchone()[0]
    results.append(check("常规查询可用", (q1 + q2 + q3) > 0,
                         f"金额>10万 {q1} 份；乙方含'公司' {q2} 份；有编号 {q3} 份"))

    ocr_dir = out / "ocr"
    texts = {}
    for p in ocr_dir.glob("*.json"):
        d = json.loads(p.read_text(encoding="utf-8"))
        texts[d["file"]] = d["text"]

    random.seed(42)
    sample = random.sample(records, min(args.sample, len(records)))
    total_fields = hit_fields = 0
    misses = []
    for r in sample:
        body = "\n".join(texts.get(f, "") for f in r["pages"])
        for key in ["contract_no", "party_a", "party_b", "amount_in_words"]:
            v = r["fields"].get(key, {}).get("value")
            if not v:
                continue
            total_fields += 1
            probe = re.sub(r"\s", "", str(v))[:14]
            if probe and probe in re.sub(r"\s", "", body):
                hit_fields += 1
            else:
                misses.append(f"{r['group']}.{key}={str(v)[:18]}")
    ratio = hit_fields / total_fields if total_fields else 0
    results.append(check(f"抽样 {len(sample)} 组字段原文命中率", ratio >= 0.9,
                         f"{hit_fields}/{total_fields} = {ratio:.0%}（验证抽取与 OCR 文本一致，"
                         f"不等于图片级人工核对）；未命中：{misses[:6]}"))

    # 增量行为实测：空目录清理失效 OCR；已入库组变更必须拒绝 append。
    incremental_ok = False
    incremental_detail = ""
    try:
        from build_output import prepare_append
        with tempfile.TemporaryDirectory(prefix="contract-verify-") as td:
            base = Path(td)
            empty = base / "empty"
            work = base / "work"
            ocr = work / "ocr"
            empty.mkdir()
            ocr.mkdir(parents=True)
            old_json = ocr / "old.json"
            old_json.write_text(json.dumps({"file": "old.jpg", "source_rel": "old.jpg", "text": "old"}),
                                encoding="utf-8")
            proc = subprocess.run([sys.executable, str(src_root / "scripts" / "ocr_extract.py"),
                                   "--input", str(empty), "--out", str(work), "--incremental"],
                                  capture_output=True, text=True, encoding="utf-8",
                                  env={**os.environ, "PYTHONIOENCODING": "utf-8"}, timeout=20)
            ocr_ok = proc.returncode == 0 and not old_json.exists() and \
                json.loads((work / "ocr_failures.json").read_text(encoding="utf-8"))["failed_count"] == 0

            root = base / "source"
            root.mkdir()
            db_tmp = base / "append.db"
            c = sqlite3.connect(db_tmp)
            c.executescript((src_root / "assets" / "schema.sql").read_text(encoding="utf-8"))
            c.execute("INSERT INTO files(path,file_name,sha256,status) VALUES(?,?,?,'ocr_done')",
                      (str(root / "a.jpg"), "a.jpg", "old"))
            c.execute("INSERT INTO contracts(group_name,source_root,status) VALUES('A',?,'verified')",
                      (str(root),))
            c.commit()
            c.close()
            inv = {"root": str(root), "files": [{"rel_path": "a.jpg", "group": "A", "sha256": "new"}]}
            (base / "inventory.json").write_text(json.dumps(inv), encoding="utf-8")
            append_rejected = False
            try:
                prepare_append(db_tmp, str(root), [{"group": "A", "pages": ["a.jpg"]}],
                               inv["files"], base, True)
            except SystemExit:
                append_rejected = True
            incremental_ok = ocr_ok and append_rejected
            incremental_detail = f"空目录清理={ocr_ok}，变更组拒绝追加={append_rejected}"
    except Exception as exc:
        incremental_detail = f"增量行为测试异常：{exc}"
    results.append(check("增量行为安全（清理 / 拒绝变更追加）", incremental_ok, incremental_detail))

    # 重复副本信息落台账（真实核对：extracted 与台账列一致）
    dup_expected = sum(len(r.get("duplicate_groups", [])) for r in records)
    dup_in_ledger = 0
    try:
        from openpyxl import load_workbook
        wb = load_workbook(out / "05_contract_ledger.xlsx", read_only=True)
        ws = wb.active
        header = [c.value for c in next(ws.iter_rows(max_row=1))]
        col = header.index("重复副本组") + 1
        for row in ws.iter_rows(min_row=2):
            if row[col - 1].value:
                dup_in_ledger += len(str(row[col - 1].value).split("、"))
        wb.close()
    except Exception:
        dup_in_ledger = -1
    results.append(check("重复副本已落台账", dup_in_ledger == dup_expected,
                         f"台账 {dup_in_ledger} / extracted {dup_expected} 个副本组"))

    # 失败文件单列（红线 3）：若有失败清单，报告必须含失败节
    fail_path = out / "ocr_failures.json"
    if fail_path.exists():
        n_fail = len(json.loads(fail_path.read_text(encoding="utf-8")).get("files", []))
        report = (out / "07_report.md").read_text(encoding="utf-8")
        results.append(check("识别失败文件在报告中单列", n_fail == 0 or "识别失败文件" in report,
                             f"失败 {n_fail} 个，报告含失败节：{'识别失败文件' in report}"))
    else:
        results.append(check("识别失败文件在报告中单列", False,
                             "无 ocr_failures.json（未运行新版 ocr_extract 或全量成功）", skipped=True))

    unverified = cur.execute("SELECT COUNT(*) FROM contracts WHERE status='unverified'").fetchone()[0]
    unknown_contract_status = cur.execute(
        "SELECT COUNT(*) FROM contracts WHERE status NOT IN ('verified','unverified')").fetchone()[0]
    bad_verified = cur.execute(
        """SELECT c.id, c.group_name FROM contracts c
           WHERE c.status='verified' AND (
             EXISTS(SELECT 1 FROM review_items r WHERE r.contract_id=c.id AND r.status='pending')
             OR EXISTS(SELECT 1 FROM issues i WHERE i.contract_id=c.id AND i.status='open')
             OR EXISTS(SELECT 1 FROM field_sources s WHERE s.contract_id=c.id AND s.needs_review=1))"""
    ).fetchall()
    pending_without_key = cur.execute(
        "SELECT COUNT(*) FROM review_items WHERE status='pending' AND (review_key IS NULL OR review_key='')").fetchone()[0]
    db_pending_keys = {row[0] for row in cur.execute(
        "SELECT review_key FROM review_items WHERE status='pending' AND review_key IS NOT NULL")}
    queue_pending_keys = set()
    try:
        from openpyxl import load_workbook
        wb = load_workbook(out / "04_review_queue.xlsx", read_only=True)
        ws = wb.active
        header = [c.value for c in next(ws.iter_rows(max_row=1))]
        key_col = header.index("复核项ID")
        status_col = header.index("状态")
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[status_col] or "").strip() == "待确认":
                queue_pending_keys.add(str(row[key_col] or ""))
        wb.close()
    except Exception:
        queue_pending_keys = {"<queue-read-failed>"}
    status_ok = (not bad_verified and unknown_contract_status == 0 and pending_without_key == 0 and
                 db_pending_keys == queue_pending_keys and (n_pending == 0 or unverified > 0))
    results.append(check("未确认数据如实标注", status_ok,
                         f"unverified {unverified} 份，待复核 {n_pending} 项；"
                         f"verified 状态冲突 {bad_verified[:5]}，未知状态 {unknown_contract_status}；"
                         f"无 ID pending {pending_without_key}，队列/库 ID 一致 {db_pending_keys == queue_pending_keys}"))

    executed = [r for r in results if not r["skipped"]]
    passed = sum(1 for r in executed if r["pass"])
    skipped = [r for r in results if r["skipped"]]
    print("=" * 62)
    print(f"技能自检：{passed}/{len(executed)} 项通过"
          + (f"（另 {len(skipped)} 项跳过，不计入）" if skipped else ""))
    print("=" * 62)
    for r in results:
        tag = "SKIP" if r["skipped"] else ("PASS" if r["pass"] else "FAIL")
        print(f"[{tag}] {r['check']} — {r['detail']}")
    report_data = {"passed": passed, "total": len(executed),
                   "skipped": len(skipped), "results": results}
    (out / "acceptance.json").write_text(json.dumps(report_data, ensure_ascii=False, indent=2),
                                         encoding="utf-8")
    print(f"\n明细已写入 {out / 'acceptance.json'}")
    conn.close()
    sys.exit(0 if passed == len(executed) else 1)


if __name__ == "__main__":
    main()
