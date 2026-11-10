"""contract-master 单元测试：金额解析、明细行、页面类型、行重建、置信度复核、FTS 转义。

运行：
  python -m unittest discover tests -v        # 技能根目录下
  python tests/test_extract_fields.py         # 直接运行
"""
import hashlib
import json
import os
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from extract_fields import (  # noqa: E402
    LOW_GROUP_CONF, LOW_PAGE_CONF, attach_orphans, cn_amount, extract_group,
    lcs_len, merge_duplicates, page_type, parse_line_items, rebuild_rows, sanitize_party,
)
from query_db import fts_phrase  # noqa: E402
from ocr_extract import emit, safe_name, source_rel  # noqa: E402
from review_resolve import parse_amount  # noqa: E402
from package_dist import should_sync  # noqa: E402
import package_dist as package_dist_module  # noqa: E402
from build_output import prepare_append  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]


class TestCnAmount(unittest.TestCase):
    """大写金额解析：含角分、零占位、容错字、残缺与未知字符（红线 7）。"""

    CASES = [
        # 常规整数
        ("伍佰元整", 500),
        ("拾万元整", 100000),
        ("壹佰万元整", 1000000),
        ("贰拾捌万玖仟陆佰元整", 289600),
        ("叁佰捌拾叁万陆仟元整", 3836000),
        ("壹佰贰拾叁万肆仟伍佰陆拾柒", 1234567),   # 合成大位数用例
        # 角分（第五轮修复：曾把“角”当元尾数算成 105/503）
        ("壹佰元伍角", 100.5),
        ("壹佰元伍角整", 100.5),
        ("壹佰元贰角叁分", 100.23),
        ("壹佰元零贰角叁分", 100.23),
        ("伍佰元零叁角", 500.3),
        ("玖拾玖元零角玖分", 99.09),
        ("叁拾万零贰佰肆拾元零陆角捌分", 300240.68),
        ("壹佰元叁分", 100.03),
        # 零占位
        ("壹佰零伍元整", 105),
        ("零壹佰元整", 100),
        # OCR 错字容错（任→仟、参→叁）
        ("贰拾捌万玖任陆佰元整", 289600),
        ("叁佰元整", 300),
        # 圆字异体
        ("壹佰圆整", 100),
        # 未知字符/残缺 -> None（禁止截断硬算）
        ("壹佰元叁X整", None),
        ("贰拾捌万玖X陆佰元整", None),
        ("壹佰元伍", None),          # 残缺角分
        ("壹佰元角叁分", None),      # 角缺数字
        ("伍角", None),              # 无元界裸角分
        ("壹佰贰佰元整", None),      # 小单位重复
        ("壹万贰万元整", None),      # 大单位重复
        ("壹万亿元整", None),        # 不支持的叠乘，交人工
        ("壹贰元整", None),          # 连续数字不是规范大写金额
        ("壹零贰元整", None),        # 缺单位且含零，不得硬算
        ("", None),
    ]

    def test_all_cases(self):
        for text, want in self.CASES:
            with self.subTest(text=text):
                self.assertEqual(cn_amount(text), want)


class TestParseLineItems(unittest.TestCase):
    def test_unit_price_times_qty(self):
        rows = [["笔记本计算机", "2,800.00", "137", "383,600.00"]]
        items = parse_line_items(rows)
        self.assertEqual(len(items), 1)
        self.assertEqual((items[0]["unit"], items[0]["qty"], items[0]["line"]),
                         (2800.0, 137, 383600.0))

    def test_qty_and_unit_both_integer_prefers_smaller_as_qty(self):
        rows = [["设备", "137", "2800", "383600"]]
        items = parse_line_items(rows)
        self.assertEqual((items[0]["unit"], items[0]["qty"]), (2800, 137))

    def test_insufficient_numbers_skipped(self):
        self.assertEqual(parse_line_items([["仅两个数", "100", "200"]]), [])


class TestPageType(unittest.TestCase):
    def test_award_notice(self):
        self.assertEqual(page_type("成交通知书\n编号：X", 10), "成交通知书")

    def test_seal_page(self):
        self.assertEqual(page_type("乙方（盖章）\n法人代表\n经办人", 20), "签章页")

    def test_body(self):
        self.assertEqual(page_type("采购合同\n甲方：某某", 30), "正文")

    def test_unknown(self):
        self.assertEqual(page_type("随便一段文字", 30), "未知")


class TestRebuildRows(unittest.TestCase):
    def test_same_y_grouped_and_x_sorted(self):
        blocks = [
            {"box": [[100, 50], [150, 50], [150, 70], [100, 70]], "text": "右"},
            {"box": [[10, 52], [60, 52], [60, 68], [10, 68]], "text": "左"},
            {"box": [[10, 200], [60, 200], [60, 220], [10, 220]], "text": "第二行"},
        ]
        rows = rebuild_rows(blocks)
        self.assertEqual(rows, [["左", "右"], ["第二行"]])


class TestConfidenceReview(unittest.TestCase):
    """阶段 4 规则：低置信页/组必须进复核队列。"""

    @staticmethod
    def _page(file, conf, text="采购合同"):
        return {"file": file, "page_no": 1, "avg_confidence": conf,
                "block_count": 5, "text": text, "blocks": []}

    def test_low_page_and_group_trigger_review(self):
        rec = extract_group("t", [self._page("a.jpg", 0.42), self._page("b.jpg", 0.61)])
        conf_items = [i for i in rec["review_items"] if i["field"] == "ocr_quality"]
        self.assertEqual(len(conf_items), 2)  # 单页 <0.55 + 组均 <0.70

    def test_normal_confidence_no_extra_review(self):
        rec = extract_group("t", [self._page("a.jpg", 0.82)])
        conf_items = [i for i in rec["review_items"] if i["field"] == "ocr_quality"]
        self.assertEqual(conf_items, [])
        self.assertGreaterEqual(rec["ocr_avg_confidence"], LOW_GROUP_CONF)

    def test_thresholds_documented(self):
        self.assertEqual((LOW_PAGE_CONF, LOW_GROUP_CONF), (0.55, 0.70))


class TestSanitizeParty(unittest.TestCase):
    """甲乙方值卫生规则：可疑形态清空复核，宁缺勿错。"""

    def test_seal_placeholder_cleared(self):
        v, why = sanitize_party("（盖章）")
        self.assertEqual((v, why != ""), ("", True))

    def test_clause_text_cleared(self):
        v, why = sanitize_party("6、本合同传真签署与正本具有同等法律效力。")
        self.assertEqual(v, "")
        self.assertTrue(why)

    def test_short_fragment_cleared(self):
        v, why = sanitize_party("公司")
        self.assertEqual(v, "")
        self.assertTrue(why)

    def test_aka_tail_truncated(self):
        v, why = sanitize_party("示例电子科技有限公司（以下简称乙方)")
        self.assertEqual(v, "示例电子科技有限公司")
        self.assertEqual(why, "")

    def test_normal_value_kept(self):
        v, why = sanitize_party("某职业学院")
        self.assertEqual((v, why), ("某职业学院", ""))


class TestLcsLen(unittest.TestCase):
    def test_error_twin_detected(self):
        # 合成案例：错字版公司名混入另一方栏
        self.assertGreaterEqual(
            lcs_len("错字示例信息技术有陵", "某省示例信息技术有限公司"), 5)

    def test_normal_parties_below_threshold(self):
        self.assertLess(lcs_len("某职业学院", "另一科技有限公司"), 5)

    def test_empty(self):
        self.assertEqual(lcs_len("", "abc"), 0)


class TestFtsPhrase(unittest.TestCase):
    def test_plain_input(self):
        self.assertEqual(fts_phrase("示例设备"), '"示例设备"')

    def test_double_quote_escaped(self):
        # FTS5 phrase 内的双引号加倍；单引号是普通字符
        self.assertEqual(fts_phrase('示"例'), '"示""例"')

    def test_single_quote_kept(self):
        self.assertEqual(fts_phrase("示例's"), '"示例\'s"')


def make_page(text, file="a.jpg", conf=0.9, page_no=1, digest=None, multipage=False):
    return {"file": file, "source_rel": file.split("#p")[0], "page_no": page_no,
            "avg_confidence": conf, "block_count": 3, "text": text, "blocks": [],
            "source_sha256": digest, "is_multipage": multipage}


def make_record(group, pages, no="", amount=None, party_b="某供应商", party_a="某采购单位"):
    return {"group": group, "pages": list(pages), "ocr_avg_confidence": 0.9,
            "cn_amount": None, "review_items": [], "issues": [],
            "fields": {"contract_no": {"value": no}, "amount": {"value": amount},
                       "party_a": {"value": party_a}, "party_b": {"value": party_b},
                       "title": {"value": "采购合同"}}}


class TestAmountAuditTrail(unittest.TestCase):
    def test_consistent_uses_original_two_candidates(self):
        rec = extract_group("ok", [make_page("采购合同\n合同金额：￥1000\n人民币：壹仟元整")])
        self.assertEqual(rec.get("amount_check"), "大小写金额一致")
        self.assertEqual(rec["fields"]["amount"]["value"], 1000)

    def test_conflict_never_marked_consistent(self):
        rec = extract_group("conflict", [make_page("采购合同\n合同金额：￥1000\n人民币：贰仟元整")])
        self.assertNotIn("amount_check", rec)
        self.assertEqual(rec["fields"]["amount"]["value"], 2000)
        self.assertIn("不一致", rec["amount_check_detail"])
        self.assertTrue(rec["fields"]["amount"]["needs_review"])

    def test_same_priority_numeric_conflict_not_counted_as_cross_check(self):
        rec = extract_group("multi", [make_page(
            "采购合同\n合同金额：￥1000\n合同金额：￥2000\n人民币：壹仟元整")])
        self.assertNotIn("amount_check", rec)
        self.assertEqual(rec["numeric_candidate_conflict"], [1000.0, 2000.0])
        self.assertIn("同优先级小写候选冲突", rec["amount_check_detail"])

    def test_uppercase_only_not_counted_as_cross_check(self):
        rec = extract_group("upper", [make_page("采购合同\n人民币：壹仟元整")])
        self.assertNotIn("amount_check", rec)
        self.assertIn("仅大写", rec["amount_check_detail"])

    def test_numeric_only_not_counted_as_cross_check(self):
        rec = extract_group("numeric", [make_page("采购合同\n合同金额：￥1000")])
        self.assertNotIn("amount_check", rec)
        self.assertIn("仅小写", rec["amount_check_detail"])

    def test_contract_no_and_bank_account_boundaries(self):
        rec = extract_group("ids", [make_page(
            "采购合同\n合同编号：HT-2026（01）\n合同编号：12345678\n账号：622200001111")])
        self.assertIn("HT-2026", rec["fields"]["contract_no"]["value"])
        self.assertEqual(rec["fields"]["bank_account"]["value"], "622200001111")


class TestPageSafety(unittest.TestCase):
    def test_zero_confidence_page_triggers_even_with_high_group_average(self):
        pages = [make_page(f"采购合同 第{i}页", f"{i}.jpg", c)
                 for i, c in enumerate([0, 1, 1, 1], start=1)]
        rec = extract_group("zero", pages)
        items = [i for i in rec["review_items"] if i["field"] == "ocr_quality"]
        self.assertEqual(len(items), 1)
        self.assertIn("最低 0.00", items[0]["reason"])

    def test_empty_pages_not_deduplicated_and_are_reviewed(self):
        rec = extract_group("empty", [make_page("", "a.jpg"), make_page("", "b.jpg")])
        self.assertEqual(len(rec["pages"]), 2)
        self.assertEqual(len([i for i in rec["review_items"] if "未识别到文本" in i["reason"]]), 2)

    def test_only_exact_source_page_duplicates_are_removed(self):
        pages = [make_page("相同模板", "a.jpg", digest="same"),
                 make_page("相同模板", "a-copy.jpg", digest="same"),
                 make_page("相同模板", "b.jpg", digest="other")]
        rec = extract_group("dup", pages)
        self.assertEqual(len(rec["pages"]), 2)
        self.assertIn("b.jpg", rec["pages"])
        self.assertEqual(len({"a.jpg", "a-copy.jpg"} & set(rec["pages"])), 1)

    def test_pdf_physical_page_number_is_preserved(self):
        rec = extract_group("pdf", [
            make_page("采购合同", "x.pdf#p1", page_no=1, multipage=True),
            make_page("甲方：某采购单位", "x.pdf#p2", page_no=2, multipage=True),
        ])
        self.assertEqual(rec["fields"]["party_a"]["page_no"], 2)

    def test_natural_order_for_parenthesized_image_names(self):
        rec = extract_group("natural", [make_page("第十页", "1(10).jpg"),
                                      make_page("第二页", "1(2).jpg")])
        self.assertEqual(rec["pages"], ["1(2).jpg", "1(10).jpg"])


class TestConservativeMerge(unittest.TestCase):
    def test_same_amount_different_party_not_merged(self):
        records = [make_record("A", ["a1", "a2", "a3"], amount=100000, party_b="甲公司"),
                   make_record("B", ["b1"], amount=100000, party_b="乙公司")]
        self.assertEqual(len(merge_duplicates(records)), 2)

    def test_same_amount_missing_party_not_merged(self):
        records = [make_record("A", ["a1", "a2", "a3"], amount=100000),
                   make_record("B", ["b1"], amount=100000, party_b="")]
        self.assertEqual(len(merge_duplicates(records)), 2)

    def test_fractional_amounts_not_truncated_for_merge(self):
        records = [make_record("A", ["a1", "a2", "a3"], amount=10000.10),
                   make_record("B", ["b1"], amount=10000.90)]
        self.assertEqual(len(merge_duplicates(records)), 2)

    def test_short_number_not_strong_merge_evidence(self):
        records = [make_record("A", ["a1"], no="1", party_b="甲公司"),
                   make_record("B", ["b1"], no="1", party_b="乙公司")]
        self.assertEqual(len(merge_duplicates(records)), 2)

    def test_same_number_with_conflicting_amount_not_merged(self):
        records = [make_record("A", ["a1"], no="HT-2026-01", amount=10000),
                   make_record("B", ["b1"], no="HT-2026-01", amount=20000)]
        self.assertEqual(len(merge_duplicates(records)), 2)

    def test_valid_merge_preserves_all_pages_and_review(self):
        records = [make_record("A", ["a1", "a2", "a3"], amount=100000),
                   make_record("B", ["b1"], amount=100000)]
        merged = merge_duplicates(records)
        self.assertEqual(len(merged), 1)
        self.assertEqual(set(merged[0]["pages"]), {"a1", "a2", "a3", "b1"})
        self.assertTrue(any(i["field"] == "duplicate_group" for i in merged[0]["review_items"]))

    def test_orphan_gets_suggestion_without_auto_merge(self):
        host = make_record("host", ["h1", "h2"])
        host["class"] = "complete"
        orphan = make_record("orphan", ["o1"])
        orphan["class"] = "suspect"
        records = [host, orphan]
        text_map = {"h1": "采购合同 项目名称 相同的模板正文", "h2": "其他",
                    "o1": "采购合同 项目名称 相同的模板正文"}
        attach_orphans(records, text_map)
        self.assertEqual(len(records), 2)
        self.assertEqual(host["pages"], ["h1", "h2"])
        self.assertEqual(orphan["suggested_host"]["group"], "host")


class TestOcrOutputIdentity(unittest.TestCase):
    def test_safe_name_collision_regression(self):
        self.assertNotEqual(safe_name(Path("a.jpg")), safe_name(Path("a.png")))
        self.assertNotEqual(safe_name(Path("a/b.jpg")), safe_name(Path("a__b.jpg")))

    def test_hash_in_filename_does_not_break_structured_source(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / "root"
            out = Path(td) / "out"
            root.mkdir()
            out.mkdir()
            src = root / "a#p1.jpg"
            src.write_bytes(b"fake-image")
            result = [([[0, 0], [1, 0], [1, 1], [0, 1]], "文本", 0.9)]
            written = emit(out, root, src, [(result, 1)])
            record = json.loads((out / written[0]).read_text(encoding="utf-8"))
            self.assertEqual(source_rel(record), "a#p1.jpg")
            self.assertEqual(record["source_rel"], "a#p1.jpg")


    def test_empty_full_incremental_reconciles_deleted_sources(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            source = base / "empty"
            out = base / "out"
            ocr = out / "ocr"
            source.mkdir(); ocr.mkdir(parents=True)
            old = make_page("旧文本", "old.jpg", digest="old")
            (ocr / "old.json").write_text(json.dumps(old), encoding="utf-8")
            result = subprocess.run([sys.executable, str(ROOT / "scripts" / "ocr_extract.py"),
                                     "--input", str(source), "--out", str(out), "--incremental"],
                                    capture_output=True, text=True, timeout=20)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertFalse((ocr / "old.json").exists())
            failures = json.loads((out / "ocr_failures.json").read_text(encoding="utf-8"))
            self.assertEqual(failures["failed_count"], 0)


class TestReviewWritebackSafety(unittest.TestCase):
    def _fixture(self, td):
        db = Path(td) / "test.db"
        conn = sqlite3.connect(db)
        conn.executescript((ROOT / "assets" / "schema.sql").read_text(encoding="utf-8"))
        conn.execute("INSERT INTO runs(mode) VALUES('full')")
        conn.execute("INSERT INTO contracts(group_name, source_root, amount, status) VALUES('G','/root',100,'unverified')")
        cid = conn.execute("SELECT id FROM contracts").fetchone()[0]
        conn.execute("INSERT INTO contract_fts(rowid,title,party_a,party_b,full_text) VALUES(?,'','','','')", (cid,))
        conn.execute("INSERT INTO review_items(contract_id,field_name,current_value,reason) VALUES(?,'amount','100','test')", (cid,))
        conn.commit()
        conn.close()
        queue = Path(td) / "queue.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["合同组", "来源根目录", "字段/问题", "当前识别值", "状态", "人工修正值"])
        ws.append(["G", "/root", "amount", "100", "OK", "200"])
        wb.save(queue)
        return db, queue

    def _run(self, db, queue):
        return subprocess.run([sys.executable, str(ROOT / "scripts" / "review_resolve.py"),
                               "--db", str(db), "--queue", str(queue)],
                              capture_output=True, text=True, timeout=20)

    def test_invalid_status_never_write_even_with_fix(self):
        with tempfile.TemporaryDirectory() as td:
            db, queue = self._fixture(td)
            result = self._run(db, queue)
            self.assertEqual(result.returncode, 0, result.stderr)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT amount FROM contracts").fetchone()[0], 100)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM review_items WHERE status='pending'").fetchone()[0], 1)
            conn.close()

    def test_invalid_amount_stays_pending_then_valid_fix_verifies(self):
        with tempfile.TemporaryDirectory() as td:
            db, queue = self._fixture(td)
            wb = load_workbook(queue)
            ws = wb.active
            ws.cell(2, 5).value = "待确认"
            ws.cell(2, 6).value = "not-a-number"
            wb.save(queue)
            self._run(db, queue)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT amount FROM contracts").fetchone()[0], 100)
            conn.close()
            ws.cell(2, 6).value = "1,234.56"
            wb.save(queue)
            result = self._run(db, queue)
            self.assertEqual(result.returncode, 0, result.stderr)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT amount FROM contracts").fetchone()[0], 1234.56)
            self.assertEqual(conn.execute("SELECT status FROM contracts").fetchone()[0], "verified")
            self.assertEqual(conn.execute("SELECT method FROM field_sources WHERE field_name='amount'").fetchone()[0], "manual_review")
            conn.close()

    def test_same_field_items_require_independent_confirmation(self):
        with tempfile.TemporaryDirectory() as td:
            db, _ = self._fixture(td)
            conn = sqlite3.connect(db)
            cid = conn.execute("SELECT id FROM contracts").fetchone()[0]
            conn.execute("UPDATE review_items SET review_key='k1' WHERE contract_id=?", (cid,))
            conn.execute("""INSERT INTO review_items(review_key,contract_id,field_name,current_value,reason)
                            VALUES('k2',?,'amount','100','second')""", (cid,))
            conn.commit(); conn.close()
            queue = Path(td) / "two-items.xlsx"
            wb = Workbook(); ws = wb.active
            ws.append(["复核项ID", "合同组", "来源根目录", "字段/问题", "当前识别值", "状态", "人工修正值"])
            ws.append(["k1", "G", "/root", "amount", "100", "已确认", "200"])
            ws.append(["k2", "G", "/root", "amount", "100", "待确认", ""])
            wb.save(queue)
            result = self._run(db, queue)
            self.assertEqual(result.returncode, 0, result.stderr)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM review_items WHERE status='pending'").fetchone()[0], 1)
            self.assertEqual(conn.execute("SELECT status FROM contracts").fetchone()[0], "unverified")
            conn.close()
            ws.cell(3, 6).value = "已确认"
            wb.save(queue)
            self._run(db, queue)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM review_items WHERE status='pending'").fetchone()[0], 0)
            self.assertEqual(conn.execute("SELECT status FROM contracts").fetchone()[0], "verified")
            self.assertEqual(conn.execute("SELECT amount FROM contracts").fetchone()[0], 200)
            self.assertEqual(conn.execute("SELECT value FROM field_sources WHERE field_name='amount'").fetchone()[0], "200.0")
            conn.close()

    def test_legacy_same_field_multiple_rows_are_ambiguous(self):
        with tempfile.TemporaryDirectory() as td:
            db, queue = self._fixture(td)
            conn = sqlite3.connect(db)
            cid = conn.execute("SELECT id FROM contracts").fetchone()[0]
            conn.execute("INSERT INTO review_items(contract_id,field_name,current_value,reason) VALUES(?,'amount','100','second')", (cid,))
            conn.commit(); conn.close()
            result = self._run(db, queue)
            self.assertEqual(result.returncode, 0, result.stderr)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("SELECT amount FROM contracts").fetchone()[0], 100)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM review_items WHERE status='pending'").fetchone()[0], 2)
            conn.close()

    def test_amount_parser_rejects_non_finite(self):
        self.assertIsNone(parse_amount("nan"))
        self.assertIsNone(parse_amount("inf"))
        self.assertEqual(parse_amount("1,000.25"), 1000.25)


class TestAppendReconciliation(unittest.TestCase):
    def test_changed_existing_group_and_deleted_file_are_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            root = base / "root"
            out = base / "out"
            root.mkdir()
            out.mkdir()
            db = out / "06_contracts.db"
            conn = sqlite3.connect(db)
            conn.executescript((ROOT / "assets" / "schema.sql").read_text(encoding="utf-8"))
            conn.execute("INSERT INTO runs(mode) VALUES('full')")
            conn.execute("INSERT INTO files(path,file_name,sha256,status) VALUES(?,?,?,'ocr_done')",
                         (str(root / "a.jpg"), "a.jpg", "old"))
            conn.execute("INSERT INTO contracts(group_name,source_root) VALUES('1',?)", (str(root),))
            conn.commit()
            conn.close()
            inv = {"root": str(root), "files": [{"rel_path": "a.jpg", "group": "1", "sha256": "new"}]}
            (out / "inventory.json").write_text(json.dumps(inv), encoding="utf-8")
            records = [{"group": "1"}]
            with self.assertRaises(SystemExit):
                prepare_append(db, str(root), records, inv["files"], out, True)
            (out / "inventory.json").write_text(json.dumps({"root": str(root), "files": []}), encoding="utf-8")
            with self.assertRaises(SystemExit):
                prepare_append(db, str(root), records, [], out, True)

    def test_change_in_merged_duplicate_is_also_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            root, out = base / "root", base / "out"
            root.mkdir(); out.mkdir()
            db = out / "06_contracts.db"
            conn = sqlite3.connect(db)
            conn.executescript((ROOT / "assets" / "schema.sql").read_text(encoding="utf-8"))
            conn.execute("INSERT INTO files(path,file_name,sha256,status) VALUES(?,?,?,'ocr_done')",
                         (str(root / "copy.jpg"), "copy.jpg", "old"))
            conn.execute("INSERT INTO contracts(group_name,source_root) VALUES('MAIN',?)", (str(root),))
            conn.commit(); conn.close()
            inv = {"root": str(root), "files": [
                {"rel_path": "copy.jpg", "group": "COPY", "sha256": "new"}]}
            (out / "inventory.json").write_text(json.dumps(inv), encoding="utf-8")
            records = [{"group": "MAIN", "pages": ["copy.jpg"]}]
            with self.assertRaises(SystemExit):
                prepare_append(db, str(root), records, inv["files"], out, True)

    def test_new_duplicate_relationship_cannot_bypass_append_guard(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            root, out = base / "root", base / "out"
            root.mkdir(); out.mkdir()
            db = out / "06_contracts.db"
            conn = sqlite3.connect(db)
            conn.executescript((ROOT / "assets" / "schema.sql").read_text(encoding="utf-8"))
            conn.execute("INSERT INTO contracts(group_name,source_root,status) VALUES('A',?,'verified')", (str(root),))
            conn.commit(); conn.close()
            inv = {"root": str(root), "files": [{"rel_path": "b.jpg", "group": "B", "sha256": "new"}]}
            (out / "inventory.json").write_text(json.dumps(inv), encoding="utf-8")
            cases = [[{"group": "A", "pages": [], "duplicate_groups": ["B"]}],
                     [{"group": "B", "pages": [], "duplicate_groups": ["A"]}]]
            for records in cases:
                with self.subTest(records=records), self.assertRaises(SystemExit):
                    prepare_append(db, str(root), records, inv["files"], out, True)

    def test_existing_pending_queue_is_not_overwritten_by_append(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            root, out = base / "root", base / "out"
            root.mkdir(); out.mkdir()
            db = out / "06_contracts.db"
            conn = sqlite3.connect(db)
            conn.executescript((ROOT / "assets" / "schema.sql").read_text(encoding="utf-8"))
            conn.execute("INSERT INTO contracts(group_name,source_root) VALUES('A',?)", (str(root),))
            conn.execute("INSERT INTO review_items(contract_id,field_name,reason) VALUES(1,'title','pending')")
            conn.commit(); conn.close()
            inv = {"root": str(root), "files": [{"rel_path": "b.jpg", "group": "B", "sha256": "new"}]}
            (out / "inventory.json").write_text(json.dumps(inv), encoding="utf-8")
            with self.assertRaises(SystemExit):
                prepare_append(db, str(root), [{"group": "B", "pages": []}], inv["files"], out, True)

    def test_same_group_in_other_source_root_is_not_blocked(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            root1, root2, out = base / "r1", base / "r2", base / "out"
            root1.mkdir(); root2.mkdir(); out.mkdir()
            db = out / "06_contracts.db"
            conn = sqlite3.connect(db)
            conn.executescript((ROOT / "assets" / "schema.sql").read_text(encoding="utf-8"))
            conn.execute("INSERT INTO contracts(group_name,source_root) VALUES('1',?)", (str(root1),))
            conn.commit(); conn.close()
            inv = {"root": str(root2), "files": []}
            (out / "inventory.json").write_text(json.dumps(inv), encoding="utf-8")
            output, existing, _ = prepare_append(db, str(root2), [{"group": "1"}], [], out, True)
            self.assertEqual(len(output), 1)
            self.assertEqual(existing, set())


class TestReleaseHygiene(unittest.TestCase):
    def test_package_allowlist_rejects_runtime_data(self):
        self.assertTrue(should_sync(Path("LICENSE")))
        self.assertTrue(should_sync(Path("README.md")))
        self.assertTrue(should_sync(Path("docs/images/workflow.png")))
        self.assertFalse(should_sync(Path("docs/images/contract.png")))
        self.assertFalse(should_sync(Path("docs/images/workflow.jpg")))
        self.assertTrue(should_sync(Path("scripts/query_db.py")))
        self.assertTrue(should_sync(Path("templates/run-config.example.json")))
        self.assertFalse(should_sync(Path("scripts/output/ocr.json")))
        self.assertFalse(should_sync(Path("references/06_contracts.db")))
        self.assertFalse(should_sync(Path("templates/other.json")))

    def test_missing_required_scripts_abort_packaging(self):
        with tempfile.TemporaryDirectory() as td:
            original = package_dist_module.SRC
            try:
                package_dist_module.SRC = Path(td)
                with self.assertRaises(SystemExit):
                    package_dist_module.validate_sources()
            finally:
                package_dist_module.SRC = original

    def test_all_out_guards_reject_skill_subdirectory(self):
        out = ROOT / "release-guard-should-not-exist"
        self.assertFalse(out.exists())
        with tempfile.TemporaryDirectory() as td:
            commands = []
            for name in ("scan_files.py", "ocr_extract.py"):
                commands.append([sys.executable, str(ROOT / "scripts" / name),
                                 "--input", td, "--out", str(out)])
            for name in ("extract_fields.py", "build_output.py", "verify_skill.py"):
                commands.append([sys.executable, str(ROOT / "scripts" / name), "--out", str(out)])
            for command in commands:
                with self.subTest(command=command[1]):
                    result = subprocess.run(command, capture_output=True, text=True, timeout=20)
                    self.assertNotEqual(result.returncode, 0)
            self.assertFalse(out.exists())
            for name in ("scan_files.py", "ocr_extract.py"):
                result = subprocess.run([sys.executable, str(ROOT / "scripts" / name),
                                         "--input", td, "--out", str(Path(td) / "generated")],
                                        capture_output=True, text=True, timeout=20)
                self.assertNotEqual(result.returncode, 0)


class TestBuildOutputIntegration(unittest.TestCase):
    def test_unprocessed_is_pending_and_unlocated_source_is_null(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            root = base / "root"
            out = base / "out"
            ocr = out / "ocr"
            root.mkdir(); ocr.mkdir(parents=True)
            (root / "a.jpg").write_bytes(b"a")
            (root / "b.jpg").write_bytes(b"b")
            sha_a = hashlib.sha256(b"a").hexdigest()
            sha_b = hashlib.sha256(b"b").hexdigest()
            inv = {"root": str(root), "files": [
                {"rel_path": "a.jpg", "group": "A", "sha256": sha_a, "ext": ".jpg", "size_kb": 0.001},
                {"rel_path": "b.jpg", "group": "B", "sha256": sha_b, "ext": ".jpg", "size_kb": 0.001}],
                "contract_groups": []}
            (out / "inventory.json").write_text(json.dumps(inv), encoding="utf-8")
            page = make_page("采购合同", "a.jpg", digest=sha_a)
            (ocr / "a.json").write_text(json.dumps(page), encoding="utf-8")
            record = make_record("A", ["a.jpg"])
            record["class"] = "suspect"
            record["status"] = "unverified"
            record["page_composition"] = "正文"
            record["fields"] = {"title": {"value": "未在原文出现的标题", "source_file": "",
                                           "confidence": 0.85, "needs_review": True}}
            record["review_items"] = [{"field": "title", "reason": "测试无来源一"},
                                      {"field": "title", "reason": "测试无来源二"}]
            (out / "extracted.json").write_text(json.dumps([record]), encoding="utf-8")
            result = subprocess.run([sys.executable, str(ROOT / "scripts" / "build_output.py"),
                                     "--out", str(out)], capture_output=True, text=True,
                                    encoding="gbk", env={**os.environ, "PYTHONIOENCODING": "gbk"}, timeout=20)
            self.assertEqual(result.returncode, 0, result.stderr)
            conn = sqlite3.connect(out / "06_contracts.db")
            statuses = dict(conn.execute("SELECT file_name,status FROM files").fetchall())
            self.assertEqual(statuses, {"a.jpg": "ocr_done", "b.jpg": "pending"})
            self.assertIsNone(conn.execute("SELECT file_id FROM field_sources WHERE field_name='title'").fetchone()[0])
            db_review_keys = {row[0] for row in conn.execute("SELECT review_key FROM review_items")}
            conn.close()
            wb = load_workbook(out / "04_review_queue.xlsx", read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            headers = list(rows[0])
            self.assertIn("来源页码", headers)
            self.assertIn("原文片段", headers)
            queue_review_keys = {row[0] for row in rows[1:]}
            self.assertEqual(len(queue_review_keys), 2)
            self.assertEqual(queue_review_keys, db_review_keys)
            wb.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
