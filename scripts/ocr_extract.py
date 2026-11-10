"""本地离线 OCR：合同图片/PDF -> 逐页文本、置信度与坐标。

用法：
  python ocr_extract.py --input <目录> --out <输出目录> [--recursive] [--workers 1]
  python ocr_extract.py --input <目录> --files 1.jpg,1-1.jpg --out <目录>
  python ocr_extract.py --input <目录> --group 4 --out <目录>

不联网、不修改原始文件。结果写入输出目录的 ocr/ 下。
--workers > 1 时用进程池并行识别，每个 worker 独立加载模型（约 2-3 秒）。
默认 1；实测进程并行可能更慢，只有在目标设备小样本实测有收益时才调大。
"""
import argparse
import hashlib
import json
import re
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
_ENGINE = None


def group_key(rel: Path) -> str:
    """子目录用目录名做组；根目录用文件名前缀。"""
    if rel.parent.as_posix() != ".":
        return rel.parent.as_posix()
    m = re.match(r"^(\d+)", rel.stem)
    return m.group(1) if m else rel.stem


def page_no(name: str) -> int:
    m = re.match(r"^\d+-(\d+)$", Path(name).stem)
    return int(m.group(1)) + 1 if m else 1


def safe_name(rel: Path) -> str:
    """生成稳定且不碰撞的输出名；同名不同扩展名、目录分隔符重名均不覆盖。"""
    rel_text = rel.as_posix()
    stem = rel.with_suffix("").as_posix().replace("/", "__")
    digest = hashlib.sha256(rel_text.encode("utf-8")).hexdigest()[:10]
    return f"{stem[:120]}-{digest}"


def collect_files(root: Path, files: str, group: str, recursive: bool):
    if files:
        return [root / n.strip() for n in files.split(",") if n.strip()]
    it = root.rglob("*") if recursive else root.glob("*")
    found = []
    for p in sorted(it):
        if not (p.is_file() and p.suffix.lower() in IMG_EXTS | {".pdf"}):
            continue
        rel = p.relative_to(root)
        if group is None or group_key(rel) == group:
            found.append(p)
    return found


def ocr_pdf(path: Path, engine, scale: float = 2.0):
    """逐页渲染 PDF 并识别，返回每页的 RapidOCR 结果列表。"""
    import numpy as np
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(path))
    results = []
    try:
        for i in range(len(pdf)):
            bitmap = pdf[i].render(scale=scale)
            pil = bitmap.to_pil().convert("RGB")
            img = np.asarray(pil)[:, :, ::-1].copy()
            result, _ = engine(img)
            results.append(result)
    finally:
        pdf.close()
    return results


def run_one(path: Path, engine):
    """识别单个文件，返回 (page_results, error)。"""
    try:
        if path.suffix.lower() == ".pdf":
            return [(r, i + 1) for i, r in enumerate(ocr_pdf(path, engine))], None
        result, _ = engine(str(path))
        return [(result, 1)], None
    except Exception as exc:
        return [], str(exc)


def _worker_init():
    global _ENGINE
    from rapidocr_onnxruntime import RapidOCR
    _ENGINE = RapidOCR()


def _worker_file(path_str):
    return path_str, *run_one(Path(path_str), _ENGINE)


def file_sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def source_rel(record):
    """优先读结构化来源路径；#pN 仅为旧版多页 PDF 记录的兼容兜底。"""
    return record.get("source_rel") or re.sub(r"#p\d+$", "", record["file"])


def load_ocr_index(out_dir: Path):
    index = []
    for p in out_dir.glob("*.json"):
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            index.append((p, data, source_rel(data)))
        except Exception:
            continue
    return index


def emit(out_dir: Path, root: Path, path: Path, page_results):
    """原子写入识别结果，返回本次生成的 JSON 文件名列表。"""
    rel = path.relative_to(root)
    digest = file_sha(path)
    mtime = path.stat().st_mtime
    written = []
    for result, page_i in page_results:
        blocks = [{"box": box, "text": text, "score": round(float(score), 4)}
                  for box, text, score in (result or [])]
        scores = [b["score"] for b in blocks]
        page_rel = rel.as_posix() if page_i == 1 and len(page_results) == 1 else \
            f"{rel.as_posix()}#p{page_i}"
        safe = safe_name(rel) if len(page_results) == 1 else \
            f"{safe_name(rel)}__p{page_i}"
        record = {
            "file": page_rel,
            "source_rel": rel.as_posix(),
            "source_page_count": len(page_results),
            "is_multipage": len(page_results) > 1,
            "name": f"{path.stem}#p{page_i}" if len(page_results) > 1 else path.name,
            "group": group_key(rel),
            "page_no": page_i if len(page_results) > 1 else page_no(path.name),
            "block_count": len(blocks),
            "avg_confidence": round(sum(scores) / len(scores), 4) if scores else 0,
            "text": "\n".join(b["text"] for b in blocks),
            "blocks": blocks,
            "source_sha256": digest,
            "source_mtime": mtime,
        }
        target = out_dir / f"{safe}.json"
        temp = target.with_suffix(".json.tmp")
        temp.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
        temp.replace(target)
        written.append(target.name)
    return written


def cleanup_stale(index, out_dir: Path, rel: str, keep=()):
    """只清理本函数已确认属于该源文件的旧 OCR JSON，绝不触碰原始合同。"""
    removed = 0
    for p, _, source in index:
        if source == rel and p.name not in keep and p.exists():
            p.unlink()
            removed += 1
    return removed


def cleanup_deleted(index, current_rels):
    removed = 0
    for p, _, source in index:
        if source not in current_rels and p.exists():
            p.unlink()
            removed += 1
    return removed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--files", default="")
    ap.add_argument("--group", default=None)
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--workers", type=int, default=1, help="并行进程数，默认 1；需按设备实测后调整")
    ap.add_argument("--incremental", action="store_true",
                    help="增量模式：跳过已识别且源文件哈希未变的文件")
    args = ap.parse_args()

    root = Path(args.input).expanduser().resolve()
    out = Path(args.out).expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"错误：--input 不是有效目录：{root}")
    if out == ROOT or ROOT in out.parents:
        raise SystemExit("错误：--out 不能指向技能目录或其子目录，请指定独立的工作输出目录")
    if out == root or root in out.parents:
        raise SystemExit("错误：--out 不能位于合同输入目录内，避免把产物写入原件目录")
    if args.workers < 1:
        raise SystemExit("错误：--workers 必须大于等于 1")
    out_dir = out / "ocr"
    out_dir.mkdir(parents=True, exist_ok=True)

    all_targets = collect_files(root, args.files, args.group, args.recursive)
    is_full_run = not args.files and args.group is None
    current_rels = {p.relative_to(root).as_posix() for p in all_targets}
    index = load_ocr_index(out_dir)

    previous_failures = []
    fail_path = out / "ocr_failures.json"
    if fail_path.exists():
        try:
            previous_failures = json.loads(fail_path.read_text(encoding="utf-8")).get("files", [])
        except Exception as exc:
            raise SystemExit(f"错误：既有失败清单无法解析，请先修复或备份后移除：{exc}")

    def write_failures(current):
        selected = current_rels
        merged = current if is_full_run else \
            [f for f in previous_failures if f.get("file") not in selected] + current
        fail_path.write_text(json.dumps({
            "failed_count": len(merged),
            "recorded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "files": merged,
        }, ensure_ascii=False, indent=2), encoding="utf-8")

    if not all_targets:
        removed = 0
        if is_full_run:
            removed = cleanup_deleted(index, current_rels)
        write_failures([])
        print(f"NO_FILES；清理失效 OCR 记录 {removed} 条")
        return

    targets = all_targets
    if args.incremental:
        candidates = {}
        for p, data, rel in index:
            item = candidates.setdefault(rel, {"hash": data.get("source_sha256"),
                                               "expected": data.get("source_page_count"),
                                               "paths": []})
            if item["hash"] != data.get("source_sha256"):
                item["hash"] = None
            item["paths"].append(p)
        known = {rel: item["hash"] for rel, item in candidates.items()
                 if item["hash"] and (item["expected"] is None or
                                      len(item["paths"]) == item["expected"])}
        targets = []
        for p in all_targets:
            rel = p.relative_to(root).as_posix()
            if known.get(rel) and known[rel] == file_sha(p):
                continue
            targets.append(p)
        if len(targets) < len(all_targets):
            print(f"增量模式：跳过 {len(all_targets) - len(targets)} 个未变更文件")
        if not targets:
            if is_full_run:
                removed = cleanup_deleted(index, current_rels)
                if removed:
                    print(f"清理已删除源文件的旧 OCR 记录 {removed} 条")
            write_failures([])
            print("无新增或变更文件，无需识别")
            return

    started = time.time()
    done = 0
    failures = []  # 红线：失败文件必须落盘，不得只打控制台

    if args.workers > 1 and len(targets) > 1:
        import multiprocessing as mp
        n = min(args.workers, len(targets))
        with mp.Pool(n, initializer=_worker_init) as pool:
            outcomes = pool.map(_worker_file, [str(p) for p in targets])
        mode = f"并行x{n}"
    else:
        from rapidocr_onnxruntime import RapidOCR
        engine = RapidOCR()
        outcomes = [(str(p), *run_one(p, engine)) for p in targets]
        mode = "串行"

    for path_str, page_results, err in outcomes:
        path = Path(path_str)
        rel = path.relative_to(root).as_posix()
        if err:
            print(f"ERROR\t{rel}\t{err}")
            try:
                digest = file_sha(path)
            except Exception:
                digest = None
            failures.append({"file": rel, "sha256": digest, "error": str(err)[:500]})
            # 源文件已变更或丢失时，旧 OCR 绝不能继续冒充当前版本。
            for old_path, old_data, old_rel in index:
                if old_rel == rel and old_path.exists() and old_data.get("source_sha256") != digest:
                    old_path.unlink()
            continue
        written = emit(out_dir, root, path, page_results)
        done += len(written)
        cleanup_stale(index, out_dir, rel, keep=written)
        print(f"OK\t{rel}\t{len(page_results)}页")

    if is_full_run:
        removed = cleanup_deleted(index, current_rels)
        if removed:
            print(f"清理已删除源文件的旧 OCR 记录 {removed} 条")
    write_failures(failures)
    failed = len(failures)
    print(f"\n[{mode}] 完成 {done} 页，失败 {failed} 个文件，"
          f"耗时 {time.time() - started:.1f}s，输出 {out_dir}")
    if failed:
        print(f"失败清单已写入 {fail_path}（build_output 会将其单列进交付报告）")


if __name__ == "__main__":
    main()
