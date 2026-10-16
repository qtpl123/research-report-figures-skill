#!/usr/bin/env python
"""下载真实图片并登记来源信息（红线：不允许没有出处的图片进入报告）。

用法:
    python fetch_image.py --id fig-04 --url "https://..." --out assets/fig-04-device.png \
        --title 设备外观 --source "Cytiva 官网" --license "© Cytiva，仅作说明引用" \
        --author "Cytiva" --manifest assets/manifest.json

说明:
- 需要联网；无网络时会明确报错，不会伪造结果。
- 下载后会校验是否真的是图片，并记录 SHA256、尺寸、获取时间。
- 同名条目会被更新（不会重复登记）。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import ssl
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

from PIL import Image

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
ALLOWED_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}


def download(url: str, dest: Path, referer: str | None = None, timeout: int = 60) -> None:
    headers = {"User-Agent": UA, "Accept": "image/avif,image/webp,image/*,*/*;q=0.8"}
    if referer:
        headers["Referer"] = referer
    ctx = ssl.create_default_context()
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        data = resp.read()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(data)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--id", required=True)
    ap.add_argument("--url", required=True)
    ap.add_argument("--out", required=True, help="保存路径（相对报告目录，如 assets/fig-04.png）")
    ap.add_argument("--title", default="")
    ap.add_argument("--source", required=True, help="来源机构/文献/网站名称")
    ap.add_argument("--license", default="未标注/需确认")
    ap.add_argument("--author", default="")
    ap.add_argument("--referer", default=None)
    ap.add_argument("--manifest", default="assets/manifest.json")
    ap.add_argument("--base", default=".", help="报告目录（相对路径基于此解析）")
    ap.add_argument("--type", default="photo", choices=["photo"])
    args = ap.parse_args()

    base = Path(args.base).resolve()
    dest = (base / args.out).resolve()
    if dest.suffix.lower() not in ALLOWED_EXT:
        print(f"错误：输出文件需要图片扩展名，当前 {dest.suffix}", file=sys.stderr)
        return 2

    try:
        download(args.url, dest, referer=args.referer)
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as exc:
        print(f"下载失败：{exc}\n请确认网址可访问、或改用有明确授权/开放获取的来源。", file=sys.stderr)
        return 3

    try:
        with Image.open(dest) as im:
            im.verify()
        with Image.open(dest) as im:
            w, h = im.size
            fmt = im.format
    except Exception as exc:
        dest.unlink(missing_ok=True)
        print(f"下载内容不是有效图片，已删除：{exc}", file=sys.stderr)
        return 4

    sha = hashlib.sha256(dest.read_bytes()).hexdigest()

    manifest_path = (base / args.manifest).resolve()
    if manifest_path.exists():
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    else:
        data = {"figures": []}
    data.setdefault("figures", [])

    entry = {
        "id": args.id,
        "file": args.out.replace("\\", "/"),
        "type": "photo",
        "title": args.title,
        "origin": "real-photo",
        "source": args.source,
        "author": args.author,
        "url": args.url,
        "license": args.license,
        "accessed": datetime.now().strftime("%Y-%m-%d"),
        "sha256": sha,
        "pixels": f"{w}x{h}",
        "format": fmt,
    }
    data["figures"] = [f for f in data["figures"] if f.get("id") != args.id]
    data["figures"].append(entry)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"图片已保存：{dest}  ({w}x{h}, {fmt}, {dest.stat().st_size/1024:.0f} KB)")
    print(f"来源已登记：{manifest_path}  → 条目 {args.id}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
