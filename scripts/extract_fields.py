"""按合同分组聚合 OCR 结果，规则抽取要素并做勾稽校验。

输入：output/ocr/*.json
输出：output/extracted.json（每合同一条，字段带溯源与置信度）
"""
import json
import re
from collections import Counter
from copy import deepcopy
from pathlib import Path

CN_DIGITS = {
    "零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9,
    "壹": 1, "贰": 2, "叁": 3, "参": 3, "肆": 4, "伍": 5,
    "陆": 6, "柒": 7, "捌": 8, "玖": 9,
}
CN_UNITS = {"十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000,
            "任": 1000}
BIG_UNITS = {"万": 10000, "亿": 100000000}

# ===== 经验阈值与规则参数（来源：111 文件五轮实测；换批次先小样本校准，见 references/rules-of-thumb.md）=====
LOW_PAGE_CONF = 0.55        # 单页平均置信度低于此值 -> 该页建议人工核对
LOW_GROUP_CONF = 0.70       # 组平均置信度低于此值 -> 整组字段建议人工核对
CLAIM_MIN_AMOUNT = 10000    # 金额认领键下限：低于此值不作归并指纹，防不同合同小额巧合撞车
ORPHAN_SIM_THRESHOLD = 0.55 # 孤立单页人工归并建议阈值（只提示，不自动改组）
SEAL_PAGE_MAX_BLOCKS = 45   # 签章页判定：页面文本块数上限
LINE_ITEM_TOL = 0.011       # 明细行 单价×数量=金额 的容差（元）
TITLE_MAX_LEN = 30          # 首页标题行长度上限（超长多为正文段落）
PARTY_LCS_REVIEW = 5        # 甲乙方最长公共子串达到此长度 -> 疑为同一公司错字混入，强制复核


def cn_amount(text: str):
    """中文大写金额转数值；无法解析返回 None。

    支持角分：壹佰元伍角 -> 100.5，壹佰元贰角叁分 -> 100.23。
    以“元/圆”为界拆分整数与角分；整数中“零”仅占位不清数（壹佰零伍 -> 105）。
    残缺或未知字符一律返回 None，禁止截断后硬算（红线）。
    """
    text = re.sub(r"[人民币￥¥\s（）()大写：:，,]", "", text).replace("圆", "元")
    if not text or not any(c in CN_DIGITS or c in CN_UNITS or c in BIG_UNITS
                           or c in "元角分" for c in text):
        return None
    jiao = fen = 0
    if "元" in text:
        int_part, _, rest = text.partition("元")
        rest = rest.replace("整", "").replace("正", "")
        m = re.fullmatch(r"(?:零?([壹贰叁肆伍陆柒捌玖])角|零角)?(?:零?([壹贰叁肆伍陆柒捌玖])分)?零?",
                         rest)
        if not m:
            return None
        if m.group(1):
            jiao = CN_DIGITS[m.group(1)]
        if m.group(2):
            fen = CN_DIGITS[m.group(2)]
        text = int_part
    else:
        text = text.replace("整", "").replace("正", "")
        if "角" in text or "分" in text:
            return None  # 无“元”界的裸角分不可信

    total, section, number = 0, 0, 0
    last_small_unit = None
    seen_big_units = set()
    last_was_digit = False
    for ch in text:
        if ch == "零":
            if last_was_digit:
                return None
            continue
        if ch in CN_DIGITS:
            if last_was_digit:
                return None
            number = CN_DIGITS[ch]
            last_was_digit = True
        elif ch in CN_UNITS:
            unit = CN_UNITS[ch]
            # 同一“万/亿”段内，小单位必须按千→百→十递减，拒绝壹佰贰佰等残缺形态。
            if last_small_unit is not None and unit >= last_small_unit:
                return None
            section += (number if number else 1) * unit
            number = 0
            last_small_unit = unit
            last_was_digit = False
        elif ch in BIG_UNITS:
            # 当前算法支持常规的“亿…万…”结构，不支持“万亿”叠乘，遇到即交人工而非硬算。
            if ch in seen_big_units or (ch == "亿" and "万" in seen_big_units):
                return None
            seen_big_units.add(ch)
            section = (section + number) * BIG_UNITS[ch]
            total += section
            section, number = 0, 0
            last_small_unit = None
            last_was_digit = False
        else:
            return None
    value = total + section + number
    if jiao or fen:
        value = round(value + jiao * 0.1 + fen * 0.01, 2)
    return value


def rebuild_rows(blocks):
    """按 y 坐标把 OCR 文本块聚类成行，行内按 x 排序，恢复表格布局。"""
    items = []
    for b in blocks:
        pts = b["box"]
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        items.append({"x": sum(xs) / len(xs), "y": sum(ys) / len(ys),
                      "h": (max(ys) - min(ys)) or 10, "text": b["text"]})
    if not items:
        return []
    items.sort(key=lambda i: i["y"])
    rows, cur, cur_y, cur_h = [], [items[0]], items[0]["y"], items[0]["h"]
    for it in items[1:]:
        if abs(it["y"] - cur_y) <= max(cur_h, it["h"]) * 0.6:
            cur.append(it)
            cur_y = sum(c["y"] for c in cur) / len(cur)
            cur_h = max(cur_h, it["h"])
        else:
            rows.append(cur)
            cur, cur_y, cur_h = [it], it["y"], it["h"]
    rows.append(cur)
    return [[c["text"] for c in sorted(r, key=lambda z: z["x"])] for r in rows]


NUM_RE = re.compile(r"\d[\d,]*(?:\.\d{1,2})?")
LINE_MIN_AMOUNT = 100       # 明细行金额下限：小于此值不视为明细（过滤编号、数量等小数字）


def parse_line_items(rows):
    """在重建后的表格行里找 单价 × 数量 = 金额 三元组。"""
    items = []
    for row in rows:
        tokens = NUM_RE.findall(" ".join(row))
        nums = [float(t.replace(",", "")) for t in tokens]
        if len(nums) < 3:
            continue
        hit = None
        for a_i in range(len(nums)):
            for q_i in range(len(nums)):
                for m_i in range(len(nums)):
                    if len({a_i, q_i, m_i}) < 3:
                        continue
                    u, q, amt = nums[a_i], nums[q_i], nums[m_i]
                    if (u > 0 and float(q).is_integer() and 1 <= q <= 99999
                            and amt >= LINE_MIN_AMOUNT and abs(u * q - amt) < LINE_ITEM_TOL):
                        if float(u).is_integer() and u < q:
                            hit = (q, int(u), amt)
                        else:
                            hit = (u, int(q), amt)
                        break
                if hit:
                    break
            if hit:
                break
        if hit:
            items.append({"unit": hit[0], "qty": hit[1], "line": hit[2], "ok": True,
                          "desc": (row[0][:30] if row else "")})
    return items



PARTY_TRAP_WORDS = ("法律效力", "本合同", "经双方", "协商一致")


def lcs_len(a: str, b: str) -> int:
    """最长公共子串长度（甲乙方同名检测用）。"""
    if not a or not b:
        return 0
    prev = [0] * (len(b) + 1)
    best = 0
    for i in range(1, len(a) + 1):
        cur = [0] * (len(b) + 1)
        for j in range(1, len(b) + 1):
            if a[i - 1] == b[j - 1]:
                cur[j] = prev[j - 1] + 1
                best = max(best, cur[j])
        prev = cur
    return best


def sanitize_party(value):
    """甲乙方值卫生检查：可疑形态清空并给复核理由，宁缺勿错。

    覆盖实测踩过的坑：（盖章）占位、条款文本误抓、过短碎片、
    “（以下简称乙方）”尾巴截断。值有但错比没有更危险。
    """
    if not value:
        return value, ""
    v = value.strip().rstrip("，,；;").strip()
    if re.fullmatch(r"[（(]?\s*盖章\s*[）)]?", v):
        return "", "甲/乙方栏仅识别到（盖章），名称缺失，需人工填写"
    if "以下简称" in v:
        v = re.split(r"[（(]\s*以下简称", v)[0].strip()
        return (v, "") if len(v) >= 3 else ("", f"甲/乙方值截断后过短（{v!r}），需人工确认")
    if len(v) < 3:
        return "", f"甲/乙方值过短（{v!r}），疑为 OCR 碎片，需人工确认"
    if "。" in v or any(w in v for w in PARTY_TRAP_WORDS):
        return "", f"甲/乙方值含条款文本（{v[:20]!r}），疑为误抓，需人工确认"
    return v, ""


def find(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None


def money_values(text):
    return [float(v.replace(",", "")) for v in re.findall(r"(\d[\d,]*\.\d{2})", text)]


def locate_value(value, pages):
    """把字段值定位回真实页：返回来源文件、页码、原文片段与文本块坐标。"""
    if value in (None, "", []):
        return {}
    if isinstance(value, float):
        probe = str(int(value))
    else:
        probe = re.sub(r"\s", "", str(value))[:16]
    if len(probe) < 2:
        return {}
    for p in pages:
        flat = re.sub(r"\s", "", p["text"])
        idx = flat.find(probe)
        if idx < 0:
            continue
        snippet = flat[max(0, idx - 25):idx + len(probe) + 25]
        bbox = None
        for b in p.get("blocks", []):
            if probe[:8] in re.sub(r"\s", "", b["text"]):
                bbox = b["box"]
                break
        # 多页 PDF 必须使用物理页号；独立图片组使用组内展示顺序。
        source_page = p["page_no"] if p.get("is_multipage") else p.get("seq", p["page_no"])
        return {"source_file": p["file"], "page_no": source_page,
                "snippet": snippet, "bbox": bbox}
    return {}


def extract_group(group, pages):
    def natural_page_key(page):
        # page_no 相同的多文件页（如 1(10).jpg、1(2).jpg）按自然顺序排列。
        parts = re.split(r"(\d+)", page["file"])
        return page.get("page_no", 1), [int(x) if x.isdigit() else x for x in parts]

    seen, uniq = set(), []
    for p in sorted(pages, key=natural_page_key):
        # 只合并“同一源文件哈希 + 同一物理页”的精确副本；空白页和模板页绝不按文本去重。
        digest = p.get("source_sha256")
        key = (digest, p.get("page_no", 1)) if digest else ("unique", p["file"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)
    # 无论是否发生去重，都使用排序后的页面列表；否则 seq 会受 JSON 文件名序影响。
    pages = uniq
    for i, p in enumerate(pages, start=1):
        p["seq"] = i
    full = "\n".join(p["text"] for p in pages)
    first = pages[0]
    issues, review = [], []
    for p in pages:
        if not p.get("text", "").strip():
            review.append({"field": "ocr_quality",
                           "reason": f"页面 {p['file']} 未识别到文本，需对照原图确认是否为空白页或 OCR 失败"})

    def field(name, value, note="", conf=0.9, needs_review=False, page=None):
        return {
            # 未定位到值时保持空来源；build_output 会落 NULL，绝不拿组首文件充数。
            "value": value, "source_file": page or "", "note": note,
            "confidence": conf, "needs_review": needs_review,
        }

    rec = {"group": group, "pages": [p["file"] for p in pages], "fields": {}}
    f = rec["fields"]

    f["contract_no"] = field(
        "contract_no", find(r"合同编号[：:]\s*([^\s，,；;]+)", full) or ""
    )
    if not f["contract_no"]["value"]:
        review.append({"field": "contract_no", "reason": "未识别到合同编号，需人工确认是否无编号"})
    lines = [l.strip() for l in first["text"].split("\n") if l.strip()]
    title = next((l for l in lines[:6] if "合同" in l and len(l) <= TITLE_MAX_LEN and "编号" not in l), "")
    f["title"] = field("title", title, conf=0.85 if title else 0.0, needs_review=not title)
    if not title:
        review.append({"field": "title", "reason": "首页未识别到合同标题，可能缺封面页或需人工填写"})
    # 甲乙方：括号注释形态随合同类型变化（采购单位/供货单位/委托方/受托方/出租方/承包方…），
    # 用通用形态一次覆盖；OCR 吞掉括号等边角情况会因抽取为空进复核队列，不静默。
    party_a = find(r"甲方(?:[（(][^：]{1,12}[）)])?[：:]\s*(.+)", full)
    party_b = find(r"乙方(?:[（(][^：]{1,12}[）)])?[：:]\s*(.+)", full)
    supplier = find(r"(?:供应商|乙方|受托方|承包方|服务方|出租方)开户名称[：:]\s*(.+)", full)

    if not party_a:
        m = re.search(r"项目名称[：:]\s*([一-龥]+?)(?:采购|租赁|服务|咨询|，|,)", full)
        party_a = m.group(1) if m else ""
        if party_a:
            review.append({"field": "party_a", "reason": "甲方取自项目名称而非甲方栏，需人工确认"})
    if not party_b and supplier:
        party_b = supplier
        review.append({"field": "party_b", "reason": "乙方取自供应商开户名称，需人工确认"})

    # 值卫生：可疑形态清空并复核；甲乙同名高度相似也要复核（实测错字版乙方混入甲方栏）
    party_a, why_a = sanitize_party(party_a or "")
    party_b, why_b = sanitize_party(party_b or "")
    if why_a:
        review.append({"field": "party_a", "reason": why_a})
    if why_b:
        review.append({"field": "party_b", "reason": why_b})
    if party_a and party_b:
        if lcs_len(party_a, party_b) >= PARTY_LCS_REVIEW:
            review.append({"field": "party_a",
                           "reason": f"甲方与乙方共享长名片段（{party_a[:16]} / {party_b[:16]}），"
                                     f"疑为同一公司错字混入，需对照原件确认"})

    f["party_a"] = field("party_a", party_a, conf=0.9 if party_a else 0.0,
                         needs_review=not party_a)
    f["party_b"] = field("party_b", party_b or "", conf=0.9 if party_b else 0.0,
                         needs_review=not party_b)
    if not party_a:
        review.append({"field": "party_a", "reason": "未识别到甲方名称，需人工填写"})
    if not party_b:
        review.append({"field": "party_b", "reason": "未识别到乙方名称，需人工填写"})

    sign_place = find(r"签订地点[：:]\s*(.+)", full)
    if sign_place:
        f["sign_place"] = field("sign_place", sign_place)

    dates = sorted(set(re.findall(r"\d{4}-\d{2}-\d{2}", full)))
    f["dates_found"] = field("dates_found", dates, conf=0.95)

    total, total_method = None, ""
    num = r"((?<![\d,])(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?(?![\d,]))"
    amount_candidates = []
    # 金额引导词按优先级排列：专用词在前；末位“总价类”为新形态兜底（服务/租赁等合同），
    # 排在既有词之后，保证不改变既有批次的仲裁顺序。
    for pri, (pattern, tag) in enumerate([
        (rf"商品合计[：:]\s*{num}", "商品合计"),
        (rf"合同(?:总)?金额[^\d]{{0,24}}{num}", "合同金额"),
        (rf"中标金额[^\d]{{0,24}}{num}", "中标金额"),
        (rf"(?:合同总价(?:款)?|总价款|租金总额|费用总额|服务费总额|成交金额|暂估价)[^\d]{{0,24}}{num}", "总价类"),
        (rf"人民币[^\d￥¥]{{0,12}}[￥¥]?\s*[:：]?\s*{num}", "人民币+小写"),
        (rf"[￥¥?？]\s*[:：]?\s*{num}", "货币符号"),
        (rf"(?:合计|总计)[^\d]{{0,12}}{num}", "合计"),
        (rf"{num}\s*元", "数字+元"),
    ]):
        for v in re.findall(pattern, full):
            val = float(v.replace(",", ""))
            if tag == "数字+元" and val < 1000:
                continue
            amount_candidates.append((val, tag, pri))
    numeric_priority_conflict = []
    if amount_candidates:
        by_pri = {}
        for v, tag, pri in amount_candidates:
            by_pri.setdefault(pri, (v, tag))
        selected_pri = min(by_pri)
        total, total_method = by_pri[selected_pri]
        numeric_priority_conflict = sorted({v for v, _, pri in amount_candidates
                                            if pri == selected_pri and abs(v - total) >= 0.01})
        if numeric_priority_conflict:
            rec["numeric_candidate_conflict"] = [total] + numeric_priority_conflict
            review.append({"field": "amount",
                           "reason": f"同优先级小写金额候选冲突：{[total] + numeric_priority_conflict}，"
                                     "需人工裁定，不得标为勾稽一致"})

    cn_text, cn_value = None, None
    m = re.search(r"人民币\s*[:：]?\s*([\u4e00-\u9fa5]{2,24})[元圆](?:整|正)", full) or \
        re.search(r"([\u4e00-\u9fa5]{2,24})[元圆](?:整|正)", full) or \
        re.search(r"人民币\s*[:：]?\s*"
                  r"([\u4e00-\u9fa5]{2,24}[元圆](?:零?[壹贰叁肆伍陆柒捌玖])?[角分]"
                  r"(?:[壹贰叁肆伍陆柒捌玖]分)?[整正]?)", full)
    if m:
        cn_text = m.group(1)
        cn_value = cn_amount(cn_text)
        if cn_value is None:
            review.append({"field": "amount_in_words",
                           "reason": "大写金额无法可靠解析（疑含 OCR 错字），需人工核对"})
    f["amount_in_words"] = field("amount_in_words", cn_text or "", conf=0.85)
    if cn_text and "参" in cn_text:
        review.append({"field": "amount_in_words",
                       "reason": "大写金额含“参”，可能是“叁”的 OCR 误识或原文异体字"})

    numeric_total = total  # 仲裁前小写值；勾稽必须比较原始双候选，不能比较覆盖后的同值。
    rec["numeric_method"] = total_method
    if cn_value is not None and (total is None or abs(total - cn_value) >= 0.01):
        if total is not None:
            rec["amount_check_detail"] = "大小写金额不一致（待复核）"
            review.append({"field": "amount",
                           "reason": f"小写候选 {total} 与大写 {cn_value} 不一致，已按大写取值，需核对"})
        else:
            rec["amount_check_detail"] = "仅大写金额，小写缺失（待复核）"
            review.append({"field": "amount", "reason": "仅识别到大写金额，小写缺失，需核对"})
        total, total_method = cn_value, "大写金额"
    f["amount"] = field("amount", total, conf=0.85 if total is not None else 0.0,
                        needs_review=total is None)
    f["amount"]["method"] = total_method
    rec["cn_amount"] = cn_value
    if amount_candidates:
        rec["amount_candidates"] = [{"value": v, "method": t, "priority": pri}
                                    for v, t, pri in amount_candidates]

    if numeric_total is not None and cn_value is not None:
        if abs(numeric_total - cn_value) < 0.01:
            rec["amount_check"] = "大小写金额一致"
    elif numeric_total is not None and not cn_text:
        rec["amount_check_detail"] = "仅小写金额，大写缺失（待复核）"
        review.append({"field": "amount_in_words", "reason": "仅识别到小写金额，大写缺失，需人工核对"})
    elif numeric_total is not None and cn_text and cn_value is None:
        rec["amount_check_detail"] = "大写金额无法解析（待复核）"
        review.append({"field": "amount_in_words", "reason": "大写金额无法解析，需人工核对"})
    if numeric_priority_conflict:
        rec.pop("amount_check", None)
        detail = f"同优先级小写候选冲突：{[numeric_total] + numeric_priority_conflict}（待复核）"
        if rec.get("amount_check_detail"):
            detail += "；" + rec["amount_check_detail"]
        rec["amount_check_detail"] = detail

    page_rows = []
    for p in pages:
        page_rows += rebuild_rows(p.get("blocks", []))
    calc = parse_line_items(page_rows)
    if calc:
        rec["line_items"] = calc
        bad = [c for c in calc if not c["ok"]]
        if bad:
            issues.append({"type": "line_item_mismatch", "detail": f"{len(bad)} 条明细单价×数量≠金额"})
        if total:
            s = round(sum(c["line"] for c in calc), 2)
            rec["sum_of_lines"] = s
            if abs(s - total) >= 0.01:
                review.append({"field": "amount",
                               "reason": f"明细合计 {s} 与总额 {total} 不一致"})

    f["bank_account"] = field(
        "bank_account", find(r"(?:银行账号|账户号码|账号|帐号)[：:]\s*(\d{8,})", full) or "")
    f["bank_name"] = field("bank_name",
                           find(r"(?:供应商|乙方|受托方|承包方)?开户行(?:名称)?[：:]\s*(.+)", full) or "")

    avg_conf = round(sum(p["avg_confidence"] for p in pages) / len(pages), 4)
    rec["ocr_avg_confidence"] = avg_conf
    low_pages = [p for p in pages
                 if p.get("avg_confidence") is not None and p["avg_confidence"] < LOW_PAGE_CONF]
    if low_pages:
        review.append({"field": "ocr_quality",
                       "reason": f"{len(low_pages)} 页 OCR 置信度低于 {LOW_PAGE_CONF}"
                                 f"（最低 {min(p['avg_confidence'] for p in low_pages):.2f}），建议对照原图核对该页字段"})
    if avg_conf < LOW_GROUP_CONF:
        review.append({"field": "ocr_quality",
                       "reason": f"整组 OCR 平均置信度 {avg_conf} 低于 {LOW_GROUP_CONF}，全部字段建议人工核对"})
    ptypes = {p["file"]: page_type(p["text"], p["block_count"]) for p in pages}
    rec["page_types"] = ptypes
    rec["page_composition"] = "、".join(f"{k}×{v}" for k, v in Counter(ptypes.values()).most_common())
    rec["issues"] = issues
    rec["review_items"] = review
    rec["status"] = "unverified" if review or issues else "unverified"
    for key, fld in f.items():
        if isinstance(fld.get("value"), (str, int, float)) and fld.get("value"):
            for k2, v2 in locate_value(fld["value"], pages).items():
                if v2:
                    fld[k2] = v2
    amount_fld = f.get("amount", {})
    if amount_fld.get("value") and not amount_fld.get("snippet"):
        aw = f.get("amount_in_words", {})
        if aw.get("page_no"):
            for k2 in ("page_no", "snippet", "source_file", "bbox"):
                if aw.get(k2):
                    amount_fld[k2] = aw[k2]
            amount_fld["note"] = "金额为大写解析值，溯源指向大写原文"
    # 复核队列与字段级 needs_review 必须一致，不能只在 Excel 里提示而数据库标记为无需复核。
    for item in review:
        if item["field"] in f:
            f[item["field"]]["needs_review"] = True
    return rec


def normalize_no(value):
    return re.sub(r"[^0-9A-Za-z]", "", str(value or "")).upper()


def reliable_no(record):
    """短编号或纯字母碎片不作为强归并证据，原始编号仍保留在台账中。"""
    no = normalize_no(record["fields"].get("contract_no", {}).get("value"))
    return no if len(no) >= 4 and any(c.isdigit() for c in no) else ""


def normalize_party(value):
    return re.sub(r"[\s（）()，,。；;]", "", str(value or ""))


def claim_keys(record):
    """金额认领键使用“分”为单位的精确值，禁止截断角分后误并不同合同。"""
    vals = set()
    amt = record["fields"].get("amount", {}).get("value")
    if amt is not None and amt != "":
        vals.add(int(round(float(amt) * 100)))
    cn = record.get("cn_amount")
    if cn is not None:
        vals.add(int(round(float(cn) * 100)))
    return {v for v in vals if v >= CLAIM_MIN_AMOUNT * 100}


def core_compatible(a, b):
    """两份记录的核心证据不冲突才允许归并；缺失不用于推翻已有值。"""
    no_a, no_b = reliable_no(a), reliable_no(b)
    if no_a and no_b and no_a != no_b:
        return False
    for key in ("party_a", "party_b"):
        av = normalize_party(a["fields"].get(key, {}).get("value"))
        bv = normalize_party(b["fields"].get(key, {}).get("value"))
        if av and bv and av != bv:
            return False
    av = a["fields"].get("amount", {}).get("value")
    bv = b["fields"].get("amount", {}).get("value")
    if av not in (None, "") and bv not in (None, ""):
        if abs(float(av) - float(bv)) >= 0.01:
            return False
    return True


def merge_duplicates(records):
    """保守归并：可靠编号，或“精确金额 + 相同乙方”且核心字段不冲突。
    所有页面和复核项都保留；互补空字段可取自副本并强制人工确认。"""
    clusters = []
    for r in records:
        no = reliable_no(r)
        target = None
        if no:
            for i, rs in enumerate(clusters):
                if reliable_no(rs[0]) == no and all(core_compatible(r, x) for x in rs):
                    target = i
                    break
        if target is None:
            clusters.append([r])
        else:
            clusters[target].append(r)

    # 少页组只能并入多页组，且金额与乙方必须同时匹配。
    amt_owner = {}
    for i, rs in enumerate(clusters):
        rs.sort(key=lambda x: (-len(x["pages"]), -x["ocr_avg_confidence"]))
        if sum(len(x["pages"]) for x in rs) >= 3:
            party = normalize_party(rs[0]["fields"].get("party_b", {}).get("value"))
            if party:
                for v in claim_keys(rs[0]):
                    amt_owner.setdefault((v, party), []).append(i)
    for i, rs in enumerate(clusters):
        if not rs or sum(len(x["pages"]) for x in rs) >= 3:
            continue
        party = normalize_party(rs[0]["fields"].get("party_b", {}).get("value"))
        if not party:
            continue
        target = None
        for v in claim_keys(rs[0]):
            for candidate in amt_owner.get((v, party), []):
                if candidate != i and all(core_compatible(rs[0], x) for x in clusters[candidate]):
                    target = candidate
                    break
            if target is not None:
                break
        if target is not None:
            clusters[target].extend(rs)
            rs[:] = []

    merged = []
    for rs in clusters:
        if not rs:
            continue
        rs.sort(key=lambda x: (-len(x["pages"]), -x["ocr_avg_confidence"]))
        main = rs[0]
        no = reliable_no(main)
        main["contract_key"] = f"NO:{no}" if no else f"GRP:{main['group']}"
        main["duplicate_groups"] = [x["group"] for x in rs[1:]]
        main["duplicate_pages"] = sum(len(x["pages"]) for x in rs[1:])
        # 保留归并组的全部唯一页面，避免副本携带的独有页面被静默丢弃。
        main["pages"] = list(dict.fromkeys(pg for x in rs for pg in x["pages"]))
        for other in rs[1:]:
            for key, fld in other["fields"].items():
                current = main["fields"].get(key, {}).get("value")
                if current in (None, "", []) and fld.get("value") not in (None, "", []):
                    main["fields"][key] = deepcopy(fld)
                    main["fields"][key]["needs_review"] = True
                    main["review_items"].append({
                        "field": key,
                        "reason": f"字段取自归并副本组 {other['group']}，需人工确认",
                    })
            for item in other.get("review_items", []):
                copied = deepcopy(item)
                copied["reason"] = f"归并副本组 {other['group']}：{copied['reason']}"
                main["review_items"].append(copied)
        if rs[1:]:
            main["review_items"].append({
                "field": "duplicate_group",
                "reason": "已与组 " + "、".join(main["duplicate_groups"]) +
                          " 按编号或金额+乙方证据归并，全部页面已保留，需人工确认副本关系",
            })
        merged.append(main)
    return merged


def classify(record):
    f = record["fields"]
    score = sum(1 for k in ("contract_no", "amount", "party_a", "party_b")
                if f.get(k, {}).get("value"))
    if score >= 3:
        return "complete"
    if score == 2:
        return "partial"
    return "suspect"


def page_type(text: str, block_count: int) -> str:
    """启发式页面类型：成交通知书 / 验收单 / 报价单 / 签章页 / 正文 / 未知。
    仅供参考，不作为排除依据。"""
    head = text[:300]
    if "成交通知书" in head or "中标通知书" in head:
        return "成交通知书"
    if ("验收单" in head or "验收报告" in head) and "验收" in head[:60]:
        return "验收单"
    if "报价单" in head or "报价表" in head:
        return "报价单"
    seal_pos = text.find("（盖章）")
    if seal_pos >= 0 and block_count <= SEAL_PAGE_MAX_BLOCKS and ("法人代表" in text or "经办人" in text):
        return "签章页"
    if "合同编号" in head or "采购合同" in head or "合同" in head[:60]:
        return "正文"
    if "合同" in text:
        return "正文"
    return "未知"


def attach_orphans(records, text_map):
    """孤立页只给归并建议，不自动改组；150 字模板相似度不能替代人工判断。"""
    from difflib import SequenceMatcher

    orphans = [r for r in records if len(r["pages"]) == 1 and r.get("class") == "suspect"]
    hosts = [r for r in records if len(r["pages"]) > 1]
    suggested = 0
    for o in orphans:
        o_head = text_map.get(o["pages"][0], "")[:150]
        if not o_head:
            continue
        best, best_score = None, ORPHAN_SIM_THRESHOLD
        for h in hosts:
            for p in h["pages"]:
                score = SequenceMatcher(None, o_head, text_map.get(p, "")[:150]).ratio()
                if score > best_score:
                    best, best_score = h, score
        if best is not None:
            o["suggested_host"] = {"group": best["group"], "score": round(best_score, 2)}
            o["review_items"].append({
                "field": "group",
                "reason": f"孤立页与组 {best['group']} 的文本相似度为 {best_score:.2f}，"
                          "仅建议人工核对，不自动归并",
            })
            suggested += 1
    if suggested:
        print(f"孤立单页：{suggested} 个组给出人工归并建议（未自动改组）")
    return records


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True,
                    help="输出目录（必填，运行数据绝不写入技能安装目录）")
    args = ap.parse_args()
    out = Path(args.out).expanduser().resolve()
    root = Path(__file__).resolve().parents[1]
    if out == root or root in out.parents:
        raise SystemExit("错误：--out 不能指向技能目录或其子目录，请指定独立的工作输出目录")
    ocr_dir = out / "ocr"
    inv_path = out / "inventory.json"
    rel2group, rel2sha, group_missing = {}, {}, {}
    inv = None
    if inv_path.exists():
        inv = json.loads(inv_path.read_text(encoding="utf-8"))
        if inv.get("root"):
            source_root = Path(inv["root"]).expanduser().resolve()
            if out == source_root or source_root in out.parents:
                raise SystemExit("错误：--out 不能位于合同输入目录内")
        rel2group = {f["rel_path"]: f["group"] for f in inv["files"]}
        rel2sha = {f["rel_path"]: f.get("sha256") for f in inv["files"]}
        group_missing = {g["group"]: g.get("missing_pages", [])
                         for g in inv.get("contract_groups", []) if g.get("missing_pages")}

    groups = {}
    stale = []
    for p in sorted(ocr_dir.glob("*.json")):
        data = json.loads(p.read_text(encoding="utf-8"))
        source_rel = data.get("source_rel") or re.sub(r"#p\d+$", "", data["file"])
        data["source_rel"] = source_rel
        if inv is not None:
            if source_rel not in rel2group:
                stale.append(f"{data['file']}（源文件已删除或不在盘点中）")
                continue
            current_sha = rel2sha.get(source_rel)
            if current_sha and data.get("source_sha256") != current_sha:
                stale.append(f"{data['file']}（源文件已变更，需重跑 OCR）")
                continue
        data["group"] = rel2group.get(source_rel, data["group"])
        groups.setdefault(data["group"], []).append(data)
    if stale:
        print(f"跳过 {len(stale)} 条过期 OCR 记录：")
        for message in stale[:20]:
            print("  -", message)
    if not groups:
        raise SystemExit("错误：没有可用且与当前盘点一致的 OCR 记录，请先运行 ocr_extract.py")
    raw = [extract_group(g, pages) for g, pages in sorted(groups.items())]
    for rec in raw:
        if group_missing.get(rec["group"]):
            rec["review_items"].append({
                "field": "page_missing",
                "reason": f"盘点检测到缺页：{group_missing[rec['group']]}，需人工核对原件",
            })
    records = merge_duplicates(raw)
    for r in records:
        r["class"] = classify(r)
    text_map = {d["file"]: d["text"] for plist in groups.values() for d in plist}
    records = attach_orphans(records, text_map)
    (out / "extracted.json").write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")

    dup_groups = sum(len(r["duplicate_groups"]) for r in records)
    dup_pages = sum(r["duplicate_pages"] for r in records)
    classes = {}
    for r in records:
        classes[r["class"]] = classes.get(r["class"], 0) + 1
    print(f"归并前 {len(raw)} 组 -> 归并后 {len(records)} 组 | "
          f"归并副本组 {dup_groups} 个、涉及副本页 {dup_pages} 张（全部页面保留）")
    print(f"分类 complete {classes.get('complete',0)} / partial {classes.get('partial',0)} / "
          f"suspect {classes.get('suspect',0)}")

    total = len(records)
    with_amount = [r for r in records if r["fields"].get("amount", {}).get("value")]
    with_no = [r for r in records if r["fields"].get("contract_no", {}).get("value")]
    both_party = [r for r in records
                  if r["fields"].get("party_a", {}).get("value") and r["fields"].get("party_b", {}).get("value")]
    checked = [r for r in records if r.get("amount_check")]
    line_ok = [r for r in records if r.get("line_items")]
    review_total = sum(len(r["review_items"]) for r in records)
    reasons = {}
    for r in records:
        for i in r["review_items"]:
            key = i["reason"].split("，")[0][:18]
            reasons[key] = reasons.get(key, 0) + 1

    print(f"合同组 {total} | 有金额 {len(with_amount)} | 有编号 {len(with_no)} | 双甲乙方 {len(both_party)}")
    print(f"金额勾稽通过 {len(checked)} | 含明细表 {len(line_ok)} | 复核项合计 {review_total}")
    print("\n复核原因分布：")
    for k, v in sorted(reasons.items(), key=lambda kv: -kv[1]):
        print(f"  {v:>3}  {k}")
    print("\n未抽到金额的组：")
    for r in records:
        if not r["fields"].get("amount", {}).get("value"):
            print(f"  - {r['group']}  页数{len(r['pages'])}  OCR {r['ocr_avg_confidence']}")
    print("\n各组概览：")
    for r in records:
        f = r["fields"]
        print(f"  {r['group'][:28]:<28} 金额 {str(f.get('amount',{}).get('value')):<12} "
              f"{'勾稽OK' if r.get('amount_check') else '—':<6} 复核{len(r['review_items'])}")


if __name__ == "__main__":
    main()
