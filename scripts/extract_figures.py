#!/usr/bin/env python
"""Render a paper PDF and locate its figures, including vector figures.

Raster-only extractors miss vector figures entirely (most CS/IEEE papers draw
their framework diagrams and plots as vectors), so figures are located by
clustering drawing objects that sit directly above each "Fig. N." caption.

    extract  --pdf P --workdir W     -> pages/, crops/, manifest.json
    crop     --image I --bbox x0 y0 x1 y1 --out O   (manual repair, pixels)
    recrop   --workdir W --fig N --bbox x0 y0 x1 y1 (manual repair, PDF points)
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    from PIL import Image
except Exception as exc:  # pragma: no cover
    raise SystemExit(f"Pillow is required: {exc}")

try:
    import pdfplumber
except Exception:  # pragma: no cover
    pdfplumber = None

NUM = re.compile(r'^(\d+)\.$')
DPI = 220
SCALE = DPI / 72.0


def fail(msg: str) -> None:
    raise SystemExit(msg)


def find_pdftoppm(explicit: str | None = None) -> str:
    found = explicit or shutil.which("pdftoppm") or shutil.which("pdftoppm.cmd")
    if not found:
        fail("pdftoppm (Poppler) was not found on PATH. Pass --pdftoppm.")
    return found


def render_pages(pdf: Path, pages_dir: Path, pdftoppm: str) -> list[Path]:
    pages_dir.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run([pdftoppm, "-r", str(DPI), "-png", str(pdf.resolve()),
                           str((pages_dir / "page").resolve())],
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode:
        fail(f"pdftoppm failed ({proc.returncode}):\n{proc.stderr}")
    pages = sorted(pages_dir.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
    if not pages:
        fail("pdftoppm produced no pages.")
    return pages


def columns(page) -> tuple[tuple[float, float], tuple[float, float]]:
    """Left/right column x-ranges for a 2-column layout, from the page width."""
    w = float(page.width)
    return (0.06 * w, 0.50 * w), (0.50 * w, 0.95 * w)


def anchors(page) -> list[dict]:
    """Caption anchors. Word-level, because a caption for a two-column-spanning
    figure shares its baseline with the other column's body text."""
    ws = sorted(page.extract_words(x_tolerance=1.5, y_tolerance=3),
                key=lambda w: (w['top'], w['x0']))
    left, right = columns(page)
    out = []
    for i, w in enumerate(ws[:-1]):
        if w['text'] not in ('Fig.', 'Figure'):
            continue
        nxt = ws[i + 1]
        if abs(nxt['top'] - w['top']) > 3:
            continue
        m = NUM.match(nxt['text'])
        if not m:
            continue
        x0, top = float(w['x0']), float(w['top'])
        colstart = left[0] if x0 < right[0] else right[0]
        line = sorted((v for v in ws if abs(v['top'] - top) <= 3), key=lambda v: v['x0'])
        # a word to the left inside the same column means this is a cross-reference
        if any(v['x1'] <= x0 - 1 and v['x0'] >= colstart - 8 for v in line):
            continue
        # walk right while spacing stays word-like, to bound the caption's own
        # text on a line it may share with the facing column
        run_end, first = float(nxt['x1']), None
        for v in line:
            if v['x0'] < run_end - 0.5:
                continue
            if v['x0'] - run_end > 12:
                break
            if first is None:
                first = v['text']
            run_end = float(v['x1'])
        if first is None:          # nothing follows the number -> cross-reference
            continue
        out.append({'fig': int(m.group(1)), 'x0': x0, 'top': top, 'run_end': run_end})
    seen, keep = set(), []
    for a in sorted(out, key=lambda a: a['top']):
        if a['fig'] in seen:
            continue
        seen.add(a['fig'])
        keep.append(a)
    return keep


def graphics(page) -> list[list[float]]:
    objs = []
    for kind in ('curves', 'rects', 'lines', 'images'):
        for o in getattr(page, kind):
            x0, x1, top, bot = float(o['x0']), float(o['x1']), float(o['top']), float(o['bottom'])
            if x1 - x0 <= 0 or bot - top <= 0:
                continue
            if x1 - x0 > 0.88 * float(page.width) and bot - top < 3:   # page rules
                continue
            objs.append([x0, top, x1, bot])
    return objs


def cluster_above(cand: list[list[float]], gap: float = 26) -> list[float] | None:
    """Union of drawing objects forming one contiguous block above a caption."""
    if not cand:
        return None
    cand = sorted(cand, key=lambda g: g[3], reverse=True)
    keep, frontier = [cand[0]], cand[0][1]
    for g in cand[1:]:
        if g[3] >= frontier - gap:
            keep.append(g)
            frontier = min(frontier, g[1])
    return [min(g[0] for g in keep), min(g[1] for g in keep),
            max(g[2] for g in keep), max(g[3] for g in keep)]


def spans_gutter(page, a: dict, prev: list[dict], gs: list[list[float]]) -> bool:
    """A left-column caption belongs to a spanning figure when the facing column
    holds drawing objects rather than body text over the same vertical band."""
    left, right = columns(page)
    if a['x0'] >= right[0]:
        return False
    floor = max([p['top'] + 6 for p in prev if p['top'] < a['top']] or [0.0])
    own = cluster_above([g for g in gs if g[3] <= a['top'] + 2 and g[1] >= floor
                         and g[0] >= left[0] - 6 and g[2] <= left[1] + 6])
    if own is None:
        return True
    band = (own[1], a['top'])
    for g in gs:
        if g[2] <= right[0]:
            continue
        h = min(g[3], band[1]) - max(g[1], band[0])
        if h > 0.5 * (g[3] - g[1]) and (g[2] - g[0]) * (g[3] - g[1]) > 200:
            return True
    return False


def caption_text(page, a: dict, span2: bool) -> str:
    left, right = columns(page)
    col = (left[0], right[1]) if span2 else (left if a['x0'] < right[0] else right)
    ws = page.extract_words(x_tolerance=1.5, y_tolerance=3)
    rows = {}
    for w in ws:
        rows.setdefault(round(w['top'] / 3), []).append(w)
    text, prev_bot = [], None
    for key in sorted(rows):
        grp = sorted(rows[key], key=lambda w: w['x0'])
        top = min(w['top'] for w in grp)
        if top < a['top'] - 1:
            continue
        if abs(top - a['top']) <= 3:       # anchor line: the caption's own run
            grp = [w for w in grp if w['x0'] >= a['x0'] - 0.5 and w['x1'] <= a['run_end'] + 0.5]
        else:                              # later lines: stay inside the column
            grp = [w for w in grp if w['x0'] >= col[0] - 8 and w['x1'] <= col[1] + 8]
        if not grp:
            break
        bot = max(w['bottom'] for w in grp)
        if prev_bot is not None and top - prev_bot > 7:
            break
        text.append(' '.join(w['text'] for w in grp))
        prev_bot = bot
        if len(text) > 12:
            break
    return ' '.join(text)


def save_crop(page_png: Path, bbox_pt: list[float], out: Path, pad: float = 4) -> list[int]:
    with Image.open(page_png) as im:
        x0, y0, x1, y1 = bbox_pt
        box = (max(0, int((x0 - pad) * SCALE)), max(0, int((y0 - pad) * SCALE)),
               min(im.width, int((x1 + pad) * SCALE)), min(im.height, int((y1 + pad) * SCALE)))
        crop = im.crop(box)
        out.parent.mkdir(parents=True, exist_ok=True)
        crop.save(out)
        return [crop.width, crop.height]


def extract_command(args) -> None:
    if pdfplumber is None:
        fail("pdfplumber is required for extraction.")
    pdf = Path(args.pdf).resolve()
    work = Path(args.workdir).resolve()
    work.mkdir(parents=True, exist_ok=True)
    pages = render_pages(pdf, work / "pages", find_pdftoppm(args.pdftoppm))

    figures = []
    with pdfplumber.open(str(pdf)) as doc:
        for pno, page in enumerate(doc.pages, 1):
            gs, ancs = graphics(page), anchors(page)
            if not gs or not ancs:
                continue
            left, right = columns(page)
            for idx, a in enumerate(ancs):
                span2 = spans_gutter(page, a, ancs[:idx], gs)
                col = (left[0], right[1]) if span2 else (left if a['x0'] < right[0] else right)
                floor = 0.0
                for p in ancs[:idx]:
                    if p['top'] >= a['top']:
                        continue
                    pcol = left if p['x0'] < right[0] else right
                    if span2 or (pcol[0] < col[1] and col[0] < pcol[1]):
                        floor = max(floor, p['top'] + 6)
                bbox = cluster_above([g for g in gs
                                      if g[3] <= a['top'] + 2 and g[1] >= floor
                                      and g[0] >= col[0] - 6 and g[2] <= col[1] + 6])
                if bbox is None:
                    continue
                out = work / "crops" / f"fig{a['fig']:02d}.png"
                px = save_crop(pages[pno - 1], bbox, out)
                figures.append({
                    'fig': a['fig'], 'page': pno, 'span2': span2,
                    'bbox_pt': [round(v, 1) for v in bbox], 'crop_px': px,
                    'aspect': round(px[0] / max(1, px[1]), 2),
                    'image_path': str(out),
                    'raw_caption': caption_text(page, a, span2),
                })
    figures.sort(key=lambda f: f['fig'])
    manifest = {
        'source_pdf': str(pdf),
        'workdir': str(work),
        'pages': [str(p) for p in pages],
        'figures': figures,
    }
    (work / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1),
                                        encoding='utf-8')
    for f in figures:
        print(f"Fig.{f['fig']:>2} p{f['page']:>2} span2={int(f['span2'])} "
              f"{f['crop_px'][0]:>4}x{f['crop_px'][1]:<4} aspect={f['aspect']:<5} "
              f"| {f['raw_caption'][:58]}")
    print(f"\n{len(figures)} figures -> {work / 'manifest.json'}")
    print("REVIEW EVERY CROP before building. Clustering cannot separate a figure "
          "from a table that sits closer to it than its own internal gaps; fix "
          "those with `recrop`.")


def recrop_command(args) -> None:
    work = Path(args.workdir).resolve()
    man = json.loads((work / "manifest.json").read_text(encoding='utf-8'))
    hit = [f for f in man['figures'] if f['fig'] == args.fig]
    if not hit:
        fail(f"Fig. {args.fig} is not in the manifest.")
    f = hit[0]
    page_png = Path(man['pages'][f['page'] - 1])
    f['bbox_pt'] = [float(v) for v in args.bbox]
    f['crop_px'] = save_crop(page_png, f['bbox_pt'], Path(f['image_path']))
    f['aspect'] = round(f['crop_px'][0] / max(1, f['crop_px'][1]), 2)
    (work / "manifest.json").write_text(json.dumps(man, ensure_ascii=False, indent=1),
                                        encoding='utf-8')
    print(f"Fig.{args.fig} -> {f['crop_px'][0]}x{f['crop_px'][1]} aspect={f['aspect']}")


def crop_command(args) -> None:
    with Image.open(Path(args.image).resolve()) as im:
        crop = im.crop(tuple(args.bbox))
        out = Path(args.out).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        crop.save(out)
    print(f"Wrote {out} ({crop.width}x{crop.height})")


def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='command', required=True)

    e = sub.add_parser('extract', help='Render pages and locate figures.')
    e.add_argument('--pdf', required=True)
    e.add_argument('--workdir', required=True)
    e.add_argument('--pdftoppm')
    e.set_defaults(func=extract_command)

    r = sub.add_parser('recrop', help='Replace one figure bbox, in PDF points.')
    r.add_argument('--workdir', required=True)
    r.add_argument('--fig', type=int, required=True)
    r.add_argument('--bbox', nargs=4, type=float, required=True, metavar=('X0', 'Y0', 'X1', 'Y1'))
    r.set_defaults(func=recrop_command)

    c = sub.add_parser('crop', help='Crop any rendered image, in pixels.')
    c.add_argument('--image', required=True)
    c.add_argument('--bbox', nargs=4, type=int, required=True)
    c.add_argument('--out', required=True)
    c.set_defaults(func=crop_command)

    args = ap.parse_args(argv)
    args.func(args)


if __name__ == '__main__':
    main()
