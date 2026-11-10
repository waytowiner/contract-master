"""打包同步：把源码同步到 dist/ 与（可选）安装副本。

用法：
  python scripts/package_dist.py                    # 只同步 dist/
  python scripts/package_dist.py --install <目录>    # 同时同步安装副本

规则：
- 必需文件缺失或源码树含数据库/表格/图片/PDF 等运行数据时停止。
- 只同步白名单代码与文档；排除 __pycache__、*.pyc、acceptance.json 与非模板 JSON。
- dist 与最终清单精确一致，历史多余文件会被移除；安装副本仅清理受管旧文件。
- 同步后逐文件 SHA256 核对，不一致即报错退出。
"""
import argparse
import hashlib
import shutil
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parents[1]
SYNC_DIRS = ["scripts", "assets", "references", "templates", "tests"]
SYNC_FILES = ["SKILL.md", "LICENSE", "README.md"]
SYNC_SUFFIXES = {"scripts": {".py"}, "assets": {".sql"}, "references": {".md"},
                 "templates": {".json"}, "tests": {".py"}}
REQUIRED_FILES = {"SKILL.md", "LICENSE", "README.md", "assets/schema.sql", "references/rules-of-thumb.md",
                  "templates/run-config.example.json", "tests/test_extract_fields.py"} | {
    f"scripts/{name}.py" for name in (
        "scan_files", "ocr_extract", "extract_fields", "build_output",
        "review_resolve", "query_db", "verify_skill", "package_dist",
    )
}
EXCLUDE_NAMES = {"__pycache__"}
EXCLUDE_SUFFIXES = {".pyc"}
EXCLUDE_FILES = {"acceptance.json"}
RUNTIME_SUFFIXES = {".db", ".sqlite", ".sqlite3", ".xlsx", ".xls", ".csv", ".jpg", ".jpeg",
                    ".png", ".bmp", ".tif", ".tiff", ".webp", ".pdf"}


def iter_sources():
    for name in SYNC_FILES:
        p = SRC / name
        if p.exists():
            yield p, p
    for d in SYNC_DIRS:
        base = SRC / d
        if not base.exists():
            continue
        for p in sorted(base.rglob("*")):
            if p.is_file():
                yield p, p


def should_sync(rel: Path) -> bool:
    if any(part in EXCLUDE_NAMES for part in rel.parts):
        return False
    if rel.suffix in EXCLUDE_SUFFIXES or rel.name in EXCLUDE_FILES:
        return False
    if rel.as_posix() in SYNC_FILES:
        return True
    if not rel.parts or rel.parts[0] not in SYNC_SUFFIXES:
        return False
    if rel.parts[0] == "templates" and rel.name != "run-config.example.json":
        return False
    return rel.suffix in SYNC_SUFFIXES[rel.parts[0]]


def validate_sources():
    missing = sorted(rel for rel in REQUIRED_FILES if not (SRC / rel).is_file())
    if missing:
        raise SystemExit(f"错误：发布必需文件缺失：{missing}")
    runtime = [p.relative_to(SRC).as_posix() for p in SRC.rglob("*")
               if p.is_file() and p.suffix.lower() in RUNTIME_SUFFIXES]
    if runtime:
        raise SystemExit(f"错误：源码树含合同/运行数据，已停止打包：{runtime[:20]}")


def sha(p: Path):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def sync_to(target: Path, clean=False):
    target = target.expanduser().resolve()
    if target == SRC or target in SRC.parents or SRC in target.parents:
        raise SystemExit(f"错误：分发目标不能是源码目录、其父目录或其子目录：{target}")
    target.mkdir(parents=True, exist_ok=True)
    expected = {p.relative_to(SRC): p for p, _ in iter_sources()
                if should_sync(p.relative_to(SRC))}

    # dist 是发布成品，必须与清单精确一致；安装副本只清理受管类型的旧文件。
    if clean:
        for p in sorted(target.rglob("*"), reverse=True):
            if p.is_file() and p.relative_to(target) not in expected:
                print(f"[非清单文件] {p.relative_to(target)}（从 dist 移除）")
                p.unlink()
    else:
        for d in SYNC_DIRS:
            base = target / d
            if not base.exists():
                continue
            for p in base.rglob("*"):
                if p.is_file():
                    rel = p.relative_to(target)
                    if should_sync(rel) and rel not in expected:
                        print(f"[多余] {rel}（源码已删除，移除）")
                        p.unlink()

    copied = mismatch = 0
    for rel, src_file in expected.items():
        dst = target / rel
        if dst.exists() and sha(dst) == sha(src_file):
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_file, dst)
        copied += 1
    for rel, src_file in expected.items():
        dst = target / rel
        if not dst.exists() or sha(dst) != sha(src_file):
            print(f"[不一致] {rel}")
            mismatch += 1
    if clean:
        actual = {p.relative_to(target) for p in target.rglob("*") if p.is_file()}
        for rel in sorted(actual - set(expected)):
            print(f"[目标多余] {rel}")
            mismatch += 1
    return copied, mismatch


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--install", default=None, help="同时同步的安装副本目录")
    args = ap.parse_args()

    validate_sources()
    copied, bad = sync_to(SRC.parent / "dist" / "contract-master", clean=True)
    print(f"dist -> {SRC.parent / 'dist' / 'contract-master'}：更新 {copied} 个文件，"
          f"{'全部一致' if bad == 0 else f'{bad} 个不一致!'}")
    if bad:
        sys.exit(1)
    if args.install:
        copied, bad = sync_to(Path(args.install), clean=False)
        print(f"安装副本 -> {args.install}：更新 {copied} 个文件，"
              f"{'全部一致' if bad == 0 else f'{bad} 个不一致!'}")
        if bad:
            sys.exit(1)


if __name__ == "__main__":
    main()
