"""合同数据库查询入口：常用筛选 + 中文全文检索。

用法：
  python query_db.py --db 06_contracts.db --party 某供应商
  python query_db.py --db 06_contracts.db --min-amount 100000 --status unverified
  python query_db.py --db 06_contracts.db --keyword 设备
  python query_db.py --db 06_contracts.db --no JXVC2019

--keyword 走 LIKE（中文任意长度可靠）；--fts 走 trigram 索引（需 ≥3 字，
输入自动按 FTS5 phrase 转义，含空格按相邻短语匹配）。
"""
import argparse
import sqlite3


def fts_phrase(user_input: str) -> str:
    """把用户输入转成安全的 FTS5 phrase：双引号包裹，内部引号加倍。"""
    return '"%s"' % user_input.replace('"', '""')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", required=True)
    ap.add_argument("--party", default=None, help="甲方或乙方包含")
    ap.add_argument("--no", default=None, help="合同编号包含")
    ap.add_argument("--min-amount", type=float, default=None)
    ap.add_argument("--max-amount", type=float, default=None)
    ap.add_argument("--status", default=None, help="verified / unverified")
    ap.add_argument("--keyword", default=None, help="全文 LIKE 检索（任意长度中文）")
    ap.add_argument("--fts", default=None, help="全文索引检索（需 3 字及以上）")
    ap.add_argument("--limit", type=int, default=50)
    args = ap.parse_args()

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row
    where, params = [], []
    if args.party:
        where.append("(party_a LIKE ? OR party_b LIKE ?)")
        params += [f"%{args.party}%", f"%{args.party}%"]
    if args.no:
        where.append("contract_no LIKE ?")
        params.append(f"%{args.no}%")
    if args.min_amount is not None:
        where.append("amount >= ?")
        params.append(args.min_amount)
    if args.max_amount is not None:
        where.append("amount <= ?")
        params.append(args.max_amount)
    if args.status:
        where.append("status = ?")
        params.append(args.status)

    if args.keyword:
        where.insert(0, "id IN (SELECT rowid FROM contract_fts WHERE full_text LIKE ?)")
        params.insert(0, f"%{args.keyword}%")
    elif args.fts:
        phrase = fts_phrase(args.fts)
        try:
            conn.execute("SELECT 1 FROM contract_fts WHERE full_text MATCH ? LIMIT 1",
                         (phrase,)).fetchone()
        except sqlite3.OperationalError as exc:
            print(f"FTS 查询出错（{exc}）。--fts 需 3 字及以上连续中文；"
                  f"含空格按相邻短语匹配，请去掉空格重试，或改用 --keyword。")
            raise SystemExit(1)
        where.insert(0, "id IN (SELECT rowid FROM contract_fts WHERE full_text MATCH ?)")
        params.insert(0, phrase)

    sql = "SELECT group_name, contract_no, title, party_a, party_b, amount, status FROM contracts"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY amount DESC LIMIT ?"
    params.append(args.limit)

    rows = conn.execute(sql, params).fetchall()
    print(f"命中 {len(rows)} 条：")
    for r in rows:
        amt = f"{r['amount']:,.0f}" if r["amount"] is not None else "—"
        print(f"  {r['group_name'][:20]:<20} {r['contract_no'] or '—':<22} "
              f"{(r['party_b'] or '—')[:20]:<20} {amt:>12} {r['status']}")
    conn.close()


if __name__ == "__main__":
    main()
