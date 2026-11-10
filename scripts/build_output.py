"""把抽取结果落地为台账、复核队列、数据库、CSV 与交付报告。"""
import csv
import datetime
import hashlib
import json
import re
import sqlite3
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[1]
COLUMNS = [
    ("group", "合同组"), ("title", "合同名称"), ("contract_no", "合同编号"),
    ("party_a", "甲方"), ("party_b", "乙方"), ("amount", "金额(元)"),
    ("amount_in_words", "大写金额"), ("sign_place", "签订地点"),
    ("dates_found", "识别到日期"), ("bank_name", "开户行"), ("bank_account", "账号"),
]


def value(rec, key):
    v = rec["fields"].get(key, {}).get("value", "")
    return "、".join(v) if isinstance(v, list) else v


def save_sheet(path, headers, rows, widths=None):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    ws.freeze_panes = "A2"
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w
    wb.save(path)


def load_failures(out: Path):
    """读取 ocr_extract 落盘的失败清单；不存在视为无失败。"""
    p = out / "ocr_failures.json"
    if not p.exists():
        return []
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
        return d.get("files", [])
    except Exception as exc:
        raise ValueError(f"OCR 失败清单无法解析，不能按零失败继续：{exc}") from exc


def load_file_list(out: Path):
    """全量文件清单：优先 inventory.json（真实 ext/sha/大小/根目录）；
    缺失时从 OCR 结果反推（哈希可用，大小缺失）。"""
    inv_path = out / "inventory.json"
    failures = load_failures(out)
    if inv_path.exists():
        inv = json.loads(inv_path.read_text(encoding="utf-8"))
        files = [{
            "rel_path": f["rel_path"],
            "sha256": f.get("sha256"),
            "ext": f.get("ext"),
            "size_bytes": int(round((f.get("size_kb") or 0) * 1024)) or None,
        } for f in inv.get("files", [])]
        # 失败清单里可能有不属于本次盘点的文件（盘点后被移动/改名），兜底补入
        known = {f["rel_path"] for f in files}
        for fl in failures:
            if fl["file"] not in known:
                files.append({
                    "rel_path": fl["file"],
                    "sha256": fl.get("sha256"),
                    "ext": Path(fl["file"]).suffix.lower() or None,
                    "size_bytes": None,
                })
        return files, inv.get("root")
    print("警告：未找到 inventory.json，文件清单从 OCR 结果反推（大小与根目录不可知，路径为相对路径）")
    files, seen = [], set()
    for p in (out / "ocr").glob("*.json"):
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        rel = d.get("source_rel") or re.sub(r"#p\d+$", "", d["file"])
        if rel in seen:
            continue
        seen.add(rel)
        files.append({
            "rel_path": rel,
            "sha256": d.get("source_sha256"),
            "ext": Path(rel).suffix.lower() or None,
            "size_bytes": None,
        })
    for fl in failures:
        if fl["file"] not in seen:
            files.append({"rel_path": fl["file"], "sha256": fl.get("sha256"),
                          "ext": Path(fl["file"]).suffix.lower() or None, "size_bytes": None})
    return files, None


def page_source(value):
    return re.sub(r"#p\d+$", "", value)


def prepare_append(db_path, root_str, records, file_list, out, append):
    """追加前对账：仅允许全新合同组；变更/删除已入库组时明确拒绝，保护已确认数据。"""
    root_key = str(Path(root_str).resolve()) if root_str else ""
    if not append or not db_path.exists():
        return records, set(), root_key
    if not root_key or not (out / "inventory.json").exists():
        raise SystemExit("错误：--append 必须基于完整 inventory.json 与绝对来源根目录")
    uri = db_path.resolve().as_uri() + "?mode=ro"
    conn = sqlite3.connect(uri, uri=True)
    try:
        columns = {row[1] for row in conn.execute("PRAGMA table_info(contracts)")}
        if "source_root" not in columns:
            raise SystemExit("错误：旧数据库缺少 source_root，不能安全追加；请备份后全量重建")
        existing = {row[0] for row in conn.execute(
            "SELECT group_name FROM contracts WHERE source_root=?", (root_key,))}
        pending_existing = {row[0] for row in conn.execute(
            """SELECT DISTINCT c.group_name FROM contracts c
               JOIN review_items r ON r.contract_id=c.id AND r.status='pending'""")}
        if pending_existing:
            raise SystemExit("错误：既有合同仍有待复核项，--append 不会用新增队列覆盖旧队列；"
                             f"请先完成复核。组：{sorted(pending_existing)[:10]}")
        previous = {}
        for path, digest in conn.execute("SELECT path, sha256 FROM files"):
            try:
                rel = Path(path).resolve().relative_to(Path(root_key)).as_posix()
            except (ValueError, OSError):
                continue
            previous[rel] = digest
    finally:
        conn.close()

    inv = json.loads((out / "inventory.json").read_text(encoding="utf-8"))
    current = {f["rel_path"]: f.get("sha256") for f in inv.get("files", [])}
    deleted = sorted(set(previous) - set(current))
    if deleted:
        raise SystemExit("错误：检测到已入库来源文件被删除，--append 不会保留悬空旧记录；"
                         f"请备份后全量重建。示例：{deleted[:5]}")
    changed = sorted(rel for rel, digest in current.items()
                     if rel in previous and previous[rel] != digest)
    added = sorted(set(current) - set(previous))
    rel2group = {f["rel_path"]: f["group"] for f in inv.get("files", [])}
    touched = set(changed + added)
    affected = {rel2group.get(rel, "") for rel in touched}
    # 变更可能发生在已归并副本中，必须同时检查归并后的 records.pages，而非只看原始组号。
    for record in records:
        if any(page_source(pg) in touched for pg in record.get("pages", [])):
            affected.add(record["group"])
    conflicts = set(existing & affected)
    for record in records:
        related = {record["group"], *record.get("duplicate_groups", [])}
        record_touched = bool(related & affected) or any(
            page_source(pg) in touched for pg in record.get("pages", []))
        if record_touched and related & existing:
            conflicts.update(related & existing)
    if conflicts:
        raise SystemExit("错误：已入库合同组或其归并副本发生增页、内容变更或关系变化，"
                         "--append 已停止，不会跳过变更或覆盖已确认队列。请备份后全量重建。"
                         f"冲突组：{sorted(conflicts)[:10]}")
    output = [r for r in records if r["group"] not in existing]
    alias_conflicts = sorted({g for r in output for g in r.get("duplicate_groups", []) if g in existing})
    if alias_conflicts:
        raise SystemExit("错误：新增组试图吸收已入库组为副本，--append 已停止："
                         f"{alias_conflicts[:10]}")
    return output, existing, root_key


def ensure_review_keys(records, root_key):
    """为每一条复核/问题行生成稳定唯一 ID，回写按 ID 精确关闭，不按字段批量关闭。"""
    for record in records:
        for kind, items in (("field", record.get("review_items", [])),
                            ("issue", record.get("issues", []))):
            for index, item in enumerate(items):
                field = item.get("field") or item.get("type") or ""
                reason = item.get("reason") or item.get("detail") or ""
                raw = json.dumps([root_key, record["group"], kind, index, field, reason],
                                 ensure_ascii=False, separators=(",", ":"))
                item["_review_key"] = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def build_report(records, out, stats):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    cls = Counter(r.get("class", "partial") for r in records)
    pending = sum(len(r["review_items"]) for r in records)
    reasons = Counter(i["reason"].split("，")[0][:24] for r in records for i in r["review_items"])
    lines = [
        "# 合同整理交付报告", "",
        f"- 生成时间：{now}",
        f"- 合同组：{len(records)}（complete {cls.get('complete',0)} / partial {cls.get('partial',0)} / suspect {cls.get('suspect',0)}）",
        f"- 识别页数：{stats['pages']}，失败文件 {stats['failed']}（失败文件单列，未静默丢弃）",
        f"- 抽到金额：{sum(1 for r in records if r['fields'].get('amount',{}).get('value') is not None and r['fields'].get('amount',{}).get('value') != '')}/{len(records)}",
        f"- 金额大小写勾稽通过：{sum(1 for r in records if r.get('amount_check') == '大小写金额一致')}",
        f"- 含明细表：{sum(1 for r in records if r.get('line_items'))}",
        f"- 归并副本：组 {stats['dup_groups']} 个、涉及副本页 {stats['dup_pages']} 张（全部页面保留）",
        f"- 运行模式：{stats['mode']}；跳过已存在组 {stats['skipped_existing']} 个",
        f"- 字段溯源记录：{stats['sources']} 条",
        "",
    ]
    if stats["failed"]:
        lines += [
            "## 识别失败文件（单列，未静默丢弃）", "",
        ]
        lines += [f"- `{f['file']}`：{f.get('error','')[:100]}" for f in stats["failures"]]
        lines.append("")
    lines += [
        "## 复核队列（共 %d 项）" % pending, "",
    ]
    if pending:
        lines.append("> **注意：本批结果含未确认数据，不得表述为“已完成建档”。**")
        lines.append("> 逐项确认请编辑 `04_review_queue.xlsx`，确认后运行 `review_resolve.py` 回写数据库。")
        lines += ["", "| 复核原因 | 数量 |", "| --- | --- |"]
        lines += [f"| {k} | {v} |" for k, v in reasons.most_common()]
    elif records:
        lines.append("本次未产生规则复核项，但 OCR 与抽取结果仍需人工确认，不得直接视为已建档。")
    else:
        lines.append("本次没有新增合同组；既有已确认数据未被重写。")
    lines += [
        "",
        "## 使用说明", "",
        "- 台账：`05_contract_ledger.xlsx` / `05_contract_ledger.csv`",
        "- 数据库：`06_contracts.db`（SQLite，支持按甲乙方、金额、编号查询）",
        "- 复核闭环：改 `04_review_queue.xlsx` → `python scripts/review_resolve.py --db 06_contracts.db --queue 04_review_queue.xlsx`",
        "",
        "## 已知限制", "",
        "OCR 与规则抽取均可能出错，字段值以溯源列与复核确认为准；",
        "已知失败模式与边界见技能目录 `references/test-report-2026-11-09.md`。",
    ]
    (out / "07_report.md").write_text("\n".join(lines), encoding="utf-8")


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True,
                    help="输出目录（必填，运行数据绝不写入技能安装目录）")
    ap.add_argument("--append", action="store_true",
                    help="增量入库：不删库，跳过已存在的合同组，保留已 verified 记录")
    args = ap.parse_args()
    out = Path(args.out)
    out = out.expanduser().resolve()
    if out == ROOT.resolve() or ROOT.resolve() in out.parents:
        raise SystemExit("错误：--out 不能指向技能目录或其子目录，请指定独立的工作输出目录")
    if not (out / "inventory.json").exists():
        raise SystemExit("错误：缺少 inventory.json，请先运行 scan_files.py；不能按零缺失文件继续")
    records = json.loads((out / "extracted.json").read_text(encoding="utf-8"))
    failures = load_failures(out)
    failure_map = {f["file"]: f.get("error", "") for f in failures}

    # OCR 逐页记录（供 ocr_pages 落库与页数统计）
    ocr_records = []
    for p in (out / "ocr").glob("*.json"):
        try:
            ocr_records.append(json.loads(p.read_text(encoding="utf-8")))
        except Exception as exc:
            raise SystemExit(f"错误：OCR 记录 {p.name} 无法解析，不能静默丢页：{exc}")
    # 全量文件清单与合同主文件集合；追加前先对账，避免重写已确认队列。
    file_list, root_str = load_file_list(out)
    if root_str:
        source_root = Path(root_str).expanduser().resolve()
        if out == source_root or source_root in out.parents:
            raise SystemExit("错误：--out 不能位于合同输入目录内")
    current_sha = {f["rel_path"]: f.get("sha256") for f in file_list}
    valid_ocr = []
    for d in ocr_records:
        rel = d.get("source_rel") or page_source(d["file"])
        if rel in current_sha and (not current_sha[rel] or d.get("source_sha256") == current_sha[rel]):
            valid_ocr.append(d)
    if len(valid_ocr) != len(ocr_records):
        print(f"警告：落库时排除 {len(ocr_records) - len(valid_ocr)} 条已删除或已变更源文件的旧 OCR 记录")
    ocr_records = valid_ocr
    page_source_map = {d["file"]: d.get("source_rel") or page_source(d["file"])
                       for d in ocr_records}
    page_counts = Counter(d.get("source_rel") or page_source(d["file"]) for d in ocr_records)
    stale_pages = [(r["group"], pg) for r in records for pg in r.get("pages", [])
                   if pg not in page_source_map]
    if stale_pages:
        raise SystemExit("错误：extracted.json 含已删除、已变更或缺 OCR 的页面，"
                         "请先重跑 extract_fields.py，不能沿用旧抽取结果。"
                         f"示例：{stale_pages[:5]}")
    db_path = out / "06_contracts.db"
    output_records, existing, root_key = prepare_append(
        db_path, root_str, records, file_list, out, args.append)
    ensure_review_keys(output_records, root_key)
    main_files = {page_source_map.get(pg, page_source(pg)) for r in records for pg in r["pages"]}
    failed_files = set(failure_map)

    ledger_rows = []
    for r in output_records:
        row = [value(r, k) for k, _ in COLUMNS]
        row += [
            r.get("amount_check") or r.get("amount_check_detail", ""), r["ocr_avg_confidence"],
            len(r["review_items"]), r.get("class", ""), r["status"],
            r.get("page_composition", ""),
            "；".join(r["pages"]),
            "、".join(r.get("duplicate_groups", [])),
            (f"{r['suggested_host']['group']}({r['suggested_host']['score']})，仅建议人工核对"
             if r.get("suggested_host") else ""),
        ]
        ledger_rows.append(row)
    headers = [label for _, label in COLUMNS] + [
        "金额勾稽", "OCR置信度", "待复核数", "完整度", "状态", "页面构成", "来源图片",
        "重复副本组", "孤立页归并建议"]
    save_sheet(out / "05_contract_ledger.xlsx", headers, ledger_rows,
               widths=[8, 26, 24, 22, 26, 14, 26, 18, 22, 24, 20, 16, 12, 10, 12, 12, 12, 18, 26, 20, 18])
    with open(out / "05_contract_ledger.csv", "w", newline="", encoding="utf-8-sig") as fh:
        cw = csv.writer(fh)
        cw.writerow(headers)
        cw.writerows(ledger_rows)

    review_rows = []
    for r in output_records:
        for item in r["review_items"]:
            field_name = item["field"]
            fld = r["fields"].get(field_name, {})
            review_rows.append([
                item["_review_key"], r["group"], root_key, field_name, value(r, field_name), item["reason"],
                fld.get("source_file") or r["pages"][0], fld.get("page_no") or "",
                fld.get("snippet", ""), "待确认", "",
            ])
        for issue in r["issues"]:
            review_rows.append([issue["_review_key"], r["group"], root_key, issue["type"], "", issue["detail"],
                                r["pages"][0], "", "", "待确认", ""])
    save_sheet(out / "04_review_queue.xlsx",
               ["复核项ID", "合同组", "来源根目录", "字段/问题", "当前识别值", "复核原因", "来源图片", "来源页码",
                "原文片段", "状态", "人工修正值"],
               review_rows, widths=[26, 8, 30, 18, 28, 46, 20, 10, 46, 10, 20])

    if db_path.exists() and not args.append:
        db_path.unlink()
    conn = sqlite3.connect(db_path)
    conn.executescript((ROOT / "assets" / "schema.sql").read_text(encoding="utf-8"))
    cur = conn.cursor()
    # existing/source_root 已在写入任何交付物前由 prepare_append 核对。

    n_done = n_dup = n_fail = n_pending = 0
    rel2fileid = {}
    cur.execute("INSERT INTO runs(mode, config_json, total_files, succeeded, failed) VALUES(?,?,?,?,?)",
                ("append" if args.append else "full",
                 json.dumps({"source_root": root_key, "append": args.append}, ensure_ascii=False),
                 len(file_list), 0, len(failures)))
    run_id = cur.lastrowid

    # files 表：全量对账。未处理是 pending，不能冒充 duplicate；失败状态优先于旧 OCR。
    ocr_source_files = {d.get("source_rel") or page_source(d["file"]) for d in ocr_records}
    for f in file_list:
        rel = f["rel_path"]
        if rel in failed_files:
            status, n_fail = "ocr_failed", n_fail + 1
        elif rel in main_files:
            status, n_done = "ocr_done", n_done + 1
        elif rel in ocr_source_files:
            status, n_dup = "duplicate", n_dup + 1
        else:
            status, n_pending = "pending", n_pending + 1
        abs_path = str((Path(root_key) / rel).resolve()) if root_key else rel
        size = f["size_bytes"]
        if root_key:
            try:
                size = Path(abs_path).stat().st_size
            except OSError:
                pass
        cur.execute(
            """INSERT INTO files(path, file_name, ext, size_bytes, sha256,
               page_count, status, error_message, first_seen_run_id)
               VALUES(?,?,?,?,?,?,?,?,?)
               ON CONFLICT(path) DO UPDATE SET
                 file_name=excluded.file_name, ext=excluded.ext, size_bytes=excluded.size_bytes,
                 sha256=excluded.sha256, page_count=COALESCE(excluded.page_count, files.page_count),
                 status=excluded.status, error_message=excluded.error_message,
                 updated_at=datetime('now','localtime')""",
            (abs_path, Path(rel).name, f["ext"], size, f["sha256"],
             page_counts.get(rel), status, failure_map.get(rel), run_id))
        row = cur.execute("SELECT id FROM files WHERE path=?", (abs_path,)).fetchone()
        rel2fileid[rel] = row[0] if row else None
    cur.execute("UPDATE runs SET succeeded=? WHERE id=?",
                (n_done + n_dup, run_id))

    inserted = 0
    text_by_file = {d["file"]: d["text"] for d in ocr_records}
    page_physical_no = {d["file"]: d["page_no"] for d in ocr_records}
    for r in output_records:
        inserted += 1
        first_page = r["pages"][0]
        first_rel = page_source_map.get(first_page, page_source(first_page))
        file_id = rel2fileid.get(first_rel)
        amt = value(r, "amount")
        amt = amt if amt not in ("", None) else None  # 金额 0 是有效值，不得 or None 吞掉
        cur.execute(
            """INSERT INTO contracts(group_name, source_root, file_id, run_id,
               contract_no, title, party_a, party_b, amount, amount_in_words, sign_place, status)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (r["group"], root_key, file_id, run_id, value(r, "contract_no"), value(r, "title"),
             value(r, "party_a"), value(r, "party_b"), amt, value(r, "amount_in_words"),
             value(r, "sign_place"), r["status"]))
        cid = cur.lastrowid
        for key, fld in r["fields"].items():
            if not fld.get("value") or isinstance(fld["value"], list):
                continue
            source_page = fld.get("source_file")
            src_rel = page_source_map.get(source_page, page_source(source_page)) if source_page else None
            cur.execute(
                """INSERT INTO field_sources(contract_id, field_name, value, source_value,
                   file_id, page_no, source_page_no, snippet, confidence, method, needs_review, bbox_json)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                (cid, key, str(fld["value"]), str(fld["value"]),
                 rel2fileid.get(src_rel) if src_rel else None, fld.get("page_no"),
                 page_physical_no.get(source_page) if source_page else None,
                 fld.get("snippet", ""), fld["confidence"], fld.get("method") or "rule",
                 int(fld["needs_review"]),
                 json.dumps(fld["bbox"], ensure_ascii=False) if fld.get("bbox") else None))
        full_text = "\n".join(text_by_file.get(f, "") for f in r["pages"])
        cur.execute(
            "INSERT INTO contract_fts(rowid, title, party_a, party_b, full_text) VALUES(?,?,?,?,?)",
            (cid, value(r, "title"), value(r, "party_a"), value(r, "party_b"), full_text))
        for item in r["review_items"]:
            cur.execute("""INSERT INTO review_items(review_key, contract_id, field_name, current_value, reason)
                           VALUES(?,?,?,?,?)""",
                        (item["_review_key"], cid, item["field"], value(r, item["field"]), item["reason"]))
        for issue in r["issues"]:
            cur.execute("""INSERT INTO review_items(review_key, contract_id, field_name, current_value, reason)
                           VALUES(?,?,?,?,?)""",
                        (issue["_review_key"], cid, issue["type"], "", issue["detail"]))
            cur.execute("INSERT INTO issues(contract_id, issue_type, detail) VALUES(?,?,?)",
                        (cid, issue["type"], issue["detail"]))

    # ocr_pages：逐页全文落库（数据库自包含，脱离 ocr/ 目录仍可查全文）
    for d in ocr_records:
        rel = d.get("source_rel") or re.sub(r"#p\d+$", "", d["file"])
        fid = rel2fileid.get(rel)
        if fid is None:
            continue
        cur.execute(
            """INSERT INTO ocr_pages(file_id, page_no, text, avg_confidence, blocks_json)
               VALUES(?,?,?,?,?)
               ON CONFLICT(file_id, page_no) DO UPDATE SET
                 text=excluded.text, avg_confidence=excluded.avg_confidence,
                 blocks_json=excluded.blocks_json, created_at=datetime('now','localtime')""",
            (fid, d["page_no"], d["text"], d.get("avg_confidence"),
             json.dumps(d.get("blocks", []), ensure_ascii=False)))
    cur.execute("UPDATE runs SET finished_at=datetime('now','localtime') WHERE id=?", (run_id,))
    conn.commit()

    stats = {
        "pages": len(ocr_records),
        "failed": len(failures),
        "failures": failures,
        "dup_groups": sum(len(r.get("duplicate_groups", [])) for r in output_records),
        "dup_pages": sum(r.get("duplicate_pages", 0) for r in output_records),
        "sources": cur.execute("SELECT COUNT(*) FROM field_sources fs JOIN contracts c ON c.id=fs.contract_id WHERE c.run_id=?", (run_id,)).fetchone()[0],
        "mode": "append" if args.append else "full",
        "skipped_existing": len(records) - len(output_records),
    }
    build_report(output_records, out, stats)
    log = {
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "records": len(output_records),
        "class_breakdown": dict(Counter(r.get("class", "partial") for r in output_records)),
        "review_items": sum(len(r["review_items"]) for r in output_records),
        **{k: v for k, v in stats.items() if k != "failures"},
    }
    (out / "run_log.json").write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")

    print("已生成：")
    for name in ["05_contract_ledger.xlsx", "05_contract_ledger.csv", "04_review_queue.xlsx",
                 "06_contracts.db", "07_report.md", "run_log.json"]:
        print(" -", out / name)
    print("\n数据库：files", cur.execute("SELECT COUNT(*) FROM files").fetchone()[0],
          "| contracts", cur.execute("SELECT COUNT(*) FROM contracts").fetchone()[0],
          "| ocr_pages", cur.execute("SELECT COUNT(*) FROM ocr_pages").fetchone()[0],
          "| field_sources", cur.execute("SELECT COUNT(*) FROM field_sources").fetchone()[0],
          "| 待复核", cur.execute("SELECT COUNT(*) FROM review_items WHERE status='pending'").fetchone()[0])
    print("文件状态分布：", dict(cur.execute("SELECT status, COUNT(*) FROM files GROUP BY status").fetchall()))
    print("\n台账预览：")
    for r in output_records:
        print(f"  组{r['group']} | {value(r,'party_a')} <-> {value(r,'party_b')} | "
              f"{value(r,'amount')} 元 | {r.get('amount_check') or r.get('amount_check_detail','-')} | "
              f"待复核 {len(r['review_items'])} 项")
    conn.close()


if __name__ == "__main__":
    main()
