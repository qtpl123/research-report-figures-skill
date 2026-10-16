#!/usr/bin/env python
"""把 PDF 逐页渲染成 PNG，用于逐页检查排版（遮挡、超边距、断页、图注分离）。

用法:
    python pdf_to_png.py <input.pdf> <out_dir> [--scale 2.0]
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pypdfium2 as pdfium


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("out_dir")
    ap.add_argument("--scale", type=float, default=2.0)
    args = ap.parse_args()

    pdf_path = Path(args.pdf).resolve()
    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(str(pdf_path))
    total = len(pdf)
    for i in range(total):
        img = pdf[i].render(scale=args.scale).to_pil()
        out = out_dir / f"page-{i + 1}.png"
        img.save(out)
        print(f"page {i + 1}/{total} -> {out} ({img.width}x{img.height})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
