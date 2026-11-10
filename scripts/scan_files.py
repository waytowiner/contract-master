"""只读盘点：扫描目录、按合同分组、检测缺页与重复文件。

不修改任何原始文件。输出 output/inventory.json 与 01_inventory.xlsx。
"""
import argparse
import hashlib
import json
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[1]
EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp", ".pdf"}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def group_key(path: Path, root: Path) -> str:
    """优先用所在子目录名做组名，根目录下用文件名前缀。"""
    if path.parent != root:
        return f"{path.parent.relative_to(root).as_posix()}"
    stem = path.stem
    m = re.match(r"^(\d+)", stem)
    return m.group(1) if m else stem


def page_no(path: Path) -> int:
    stem = path.stem
    m = re.match(r"^\d+-(\d+)$", stem)
    return int(m.group(1)) + 1 if m else 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--recursive", action="store_true")
    args = ap.parse_args()

    root = Path(args.input).expanduser().resolve()
    out = Path(args.out).expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"错误：--input 不是有效目录：{root}")
    if out == ROOT or ROOT in out.parents:
        raise SystemExit("错误：--out 不能指向技能目录或其子目录，请指定独立的工作输出目录")
    if out == root or root in out.parents:
        raise SystemExit("错误：--out 不能位于合同输入目录内，避免把产物写入原件目录")
    out.mkdir(parents=True, exist_ok=True)

    it = root.rglob("*") if args.recursive else root.glob("*")
    files = sorted(p for p in it if p.is_file() and p.suffix.lower() in EXTS)

    rows, hashes = [], defaultdict(list)
    info = {}
    for p in files:
        digest = sha256(p)
        hashes[digest].append(p)
        info[p] = {
            "group": group_key(p, root),
            "page": page_no(p),
            "sha256": digest,
        }

    def strong_group(name: str) -> bool:
        return name.isdigit() or "/" in name

    for digest, paths in hashes.items():
        if len(paths) < 2:
            continue
        ranked = sorted(paths, key=lambda x: (not strong_group(info[x]["group"]),
                                              len(info[x]["group"]), x.name))
        target = info[ranked[0]]["group"]
        if not strong_group(target):
            continue
        for p in paths:
            info[p]["group"] = target
            info[p]["merged_from"] = "" if p == ranked[0] else target

    groups = defaultdict(list)
    for p in files:
        m = info[p]
        groups[m["group"]].append(p)
        rows.append({
            "group": m["group"],
            "file": p.name,
            "rel_path": p.relative_to(root).as_posix(),
            "page": m["page"],
            "ext": p.suffix.lower(),
            "size_kb": round(p.stat().st_size / 1024, 1),
            "sha256": m["sha256"],
            "duplicate_of": m.get("merged_from", ""),
        })

    duplicates = [
        {"sha256": d, "files": [x.relative_to(root).as_posix() for x in ps]}
        for d, ps in hashes.items() if len(ps) > 1
    ]

    group_info = []
    for g, ps in sorted(groups.items()):
        pages = sorted({info[p]["page"] for p in ps})
        missing = [n for n in range(1, max(pages) + 1) if n not in pages] if pages else []
        group_info.append({
            "group": g,
            "file_count": len(ps),
            "pages": pages,
            "missing_pages": missing,
            "duplicate_files": len(ps) - len({info[p]["sha256"] for p in ps}),
            "files": [p.name for p in sorted(ps, key=lambda x: info[x]["page"])],
        })

    report = {
        "root": str(root),
        "total_files": len(files),
        "groups": len(groups),
        "duplicate_groups": duplicates,
        "files": rows,
        "contract_groups": group_info,
    }
    (out / "inventory.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    wb = Workbook()
    ws = wb.active
    ws.title = "文件清单"
    ws.append(["合同组", "文件", "相对路径", "页码", "格式", "大小KB", "SHA256前12位", "重复归并"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append([r["group"], r["file"], r["rel_path"], r["page"], r["ext"],
                   r["size_kb"], r["sha256"][:12], r["duplicate_of"]])
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("分组与异常")
    ws2.append(["合同组", "文件数", "识别页码", "缺失页码", "重复文件数"])
    for c in ws2[1]:
        c.font = Font(bold=True)
    for g in group_info:
        ws2.append([g["group"], g["file_count"], "、".join(map(str, g["pages"])),
                    "、".join(map(str, g["missing_pages"])), g["duplicate_files"]])
    wb.save(out / "01_inventory.xlsx")

    print(f"文件 {len(files)} | 合同组 {len(groups)} | 重复文件组 {len(duplicates)}")
    no_page1 = [g["group"] for g in group_info if 1 not in g["pages"]]
    gaps = [g for g in group_info if g["missing_pages"]]
    print(f"缺首页的组 {len(no_page1)}: {no_page1[:10]}")
    print(f"页码不连续的组 {len(gaps)}: {[(g['group'], g['missing_pages']) for g in gaps][:10]}")
    if duplicates:
        print("重复文件示例：")
        for d in duplicates[:5]:
            print("  -", " == ".join(d["files"]))


if __name__ == "__main__":
    main()
