"""生成演示报告的 assets/manifest.json（真实图片由脚本补全 SHA256 与像素信息）。"""
import hashlib
import json
from datetime import date
from pathlib import Path

from PIL import Image

BASE = Path(__file__).resolve().parent
ASSETS = BASE / "assets"
TODAY = date.today().strftime("%Y-%m-%d")


def real_photo(fid, filename, title, source, author, url, license_str):
    p = ASSETS / filename
    with Image.open(p) as im:
        w, h = im.size
        fmt = im.format
    return {
        "id": fid,
        "file": f"assets/{filename}",
        "type": "photo",
        "title": title,
        "origin": "real-photo",
        "source": source,
        "author": author,
        "url": url,
        "license": license_str,
        "accessed": TODAY,
        "sha256": hashlib.sha256(p.read_bytes()).hexdigest(),
        "pixels": f"{w}x{h}",
        "format": fmt,
    }


def generated(fid, filename, ftype, title, generator, source, note):
    return {
        "id": fid,
        "file": f"assets/{filename}",
        "type": ftype,
        "title": title,
        "generator": generator,
        "source": source,
        "note": note,
    }


figures = [
    generated(
        "fig-01", "fig-01-chain.png", "schematic",
        "钠离子电池产业链与正极路线分岔",
        "mermaid",
        "依据中国能源产业发展网、中研网公开行业资料整理",
        "本图为技术示意图，非真实照片",
    ),
    generated(
        "fig-02", "fig-02-system.png", "schematic",
        "钠离子电池材料体系构成",
        "graphviz",
        "依据中能建能源研究院《钠离子电池技术及市场发展趋势研究报告》整理",
        "本图为技术示意图，非真实照片",
    ),
    generated(
        "fig-03", "fig-03-cost.png", "chart",
        "三大正极路线 2026 年量产成本对比",
        "matplotlib",
        "星耀新能源《2026 钠离子电池正极材料行业动态》",
        "数据为该来源披露的 2026 年量产口径，不同机构统计口径存在差异",
    ),
    generated(
        "fig-04", "fig-04-energy.png", "chart",
        "三大正极路线量产能量密度对比",
        "matplotlib",
        "星耀新能源《2026 钠离子电池正极材料行业动态》",
        "数据为该来源披露的 2026 年量产口径，不同机构统计口径存在差异",
    ),
    generated(
        "fig-05", "fig-05-cycle.png", "chart",
        "循环寿命对比：聚阴离子路线与层状氧化物路线",
        "matplotlib",
        "西能能源行业分析、同花顺财经《钠离子电池产业化最新进展》",
        "两路线数据来自不同来源，仅作量级对比，非同一测试条件下的实测结果",
    ),
    real_photo(
        "fig-06", "fig-06-cells.jpg",
        "钠离子电芯实物（2.5 Ah）",
        "Wikimedia Commons",
        "Brihaspati",
        "https://commons.wikimedia.org/wiki/File:Sodium_ion_2.5_Ah_cells.jpg",
        "CC BY-SA 4.0",
    ),
    real_photo(
        "fig-07", "fig-07-systems.png",
        "钠离子电池体系构成（文献原图）",
        "Wikimedia Commons 收录的综述论文原图（Tang-Yeon Hwang 等）",
        "Tang-Yeon Hwang, Seung-Taek Myung, Yang-Kook Sun",
        "https://commons.wikimedia.org/wiki/File:Sodium-ion_battery_systems.png",
        "CC BY 3.0",
    ),
]

out = ASSETS / "manifest.json"
out.write_text(json.dumps({"figures": figures}, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"✓ 已写入 {out}（{len(figures)} 条）")
for f in figures:
    print(f"  {f['id']}  {f['type']:9s} {f['file']}")
