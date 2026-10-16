#!/usr/bin/env python
"""数据图表渲染器（matplotlib，300 dpi，中文字体已适配）。

用法:
    python chart_render.py <spec.json>

spec.json 格式（单张图）:
{
  "figure_id": "fig-03",
  "output": "assets/fig-03-hcp-clearance.png",
  "type": "bar",                       // bar | grouped_bar | stacked_bar | line | pie | scatter | box | chromatogram
  "title": "各层析步骤的 HCP 清除效果",   // 可为空字符串
  "categories": ["捕获层析", "精炼层析", "除病毒过滤"],
  "series": [{"name": "HCP 清除率 (%)", "values": [95.2, 99.1, 99.8]}],
  "xlabel": "工艺步骤",
  "ylabel": "清除率 (%)",
  "figsize": [8.0, 4.4],
  "dpi": 300,
  "value_labels": true,               // 是否在柱/点上标数值
  "y_min": 0, "y_max": 100,           // 可选
  "colors": ["#1F4E79", "#2E75B6", "#9DC3E6"]
}

也支持一次渲染多张图:
{ "figures": [ {...}, {...} ] }
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

# matplotlib 缓存放到项目本地，避免写系统目录被拒、并加快后续启动
_MPL_CACHE = Path(__file__).resolve().parents[1] / ".mplconfig"
_MPL_CACHE.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", str(_MPL_CACHE))

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager

# ---- 中文字体 ----------------------------------------------------------------
CJK_CANDIDATES = [
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/msyhbd.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "C:/Windows/Fonts/simsun.ttc",
]
DEFAULT_COLORS = ["#1F4E79", "#2E75B6", "#9DC3E6", "#C55A11", "#548235", "#7F7F7F"]


def setup_fonts() -> str:
    """注册系统中文字体，返回实际使用的字体名。"""
    chosen = None
    for path in CJK_CANDIDATES:
        p = Path(path)
        if p.exists():
            try:
                font_manager.fontManager.addfont(str(p))
                chosen = font_manager.FontProperties(fname=str(p)).get_name()
                break
            except Exception:
                continue
    if chosen is None:
        chosen = "DejaVu Sans"
    plt.rcParams["font.family"] = [chosen, "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.facecolor"] = "white"
    plt.rcParams["axes.facecolor"] = "white"
    return chosen


FONT_NAME = setup_fonts()


def _style_axes(ax, spec: dict) -> None:
    ax.set_title(spec.get("title") or "", fontsize=14, color="#1a1a1a", pad=12)
    if spec.get("xlabel"):
        ax.set_xlabel(spec["xlabel"], fontsize=11)
    if spec.get("ylabel"):
        ax.set_ylabel(spec["ylabel"], fontsize=11)
    ax.tick_params(labelsize=10)
    if spec.get("y_min") is not None or spec.get("y_max") is not None:
        ax.set_ylim(spec.get("y_min"), spec.get("y_max"))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", linestyle=":", color="#CCCCCC", linewidth=0.8)
    ax.set_axisbelow(True)


def render(spec: dict) -> Path:
    fig_type = spec.get("type", "bar")
    cats = spec.get("categories", [])
    series = spec.get("series", [])
    colors = spec.get("colors") or DEFAULT_COLORS
    show_labels = bool(spec.get("value_labels", False))
    legend_loc = spec.get("legend_loc", "best")
    dpi = int(spec.get("dpi", 300))
    figsize = tuple(spec.get("figsize", [8.0, 4.4]))
    out = Path(spec["output"]).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    x = np.arange(len(cats))

    if fig_type in ("bar", "grouped_bar"):
        n = max(len(series), 1)
        total_w = 0.72
        w = total_w / n
        for i, s in enumerate(series):
            offset = -total_w / 2 + w * i + w / 2
            bars = ax.bar(x + offset, s["values"], width=w * 0.92,
                          label=s.get("name", ""), color=colors[i % len(colors)])
            if show_labels:
                for b in bars:
                    ax.annotate(f"{b.get_height():g}", (b.get_x() + b.get_width() / 2, b.get_height()),
                                ha="center", va="bottom", fontsize=9, color="#333333",
                                xytext=(0, 2), textcoords="offset points")
        ax.set_xticks(x)
        ax.set_xticklabels(cats, rotation=0)
        if n > 1:
            ax.legend(fontsize=10, frameon=False, loc=legend_loc)

    elif fig_type == "stacked_bar":
        bottom = np.zeros(len(cats))
        for i, s in enumerate(series):
            vals = np.array(s["values"], dtype=float)
            ax.bar(x, vals, bottom=bottom, label=s.get("name", ""),
                   color=colors[i % len(colors)], width=0.6)
            if show_labels:
                for xi, v, b0 in zip(x, vals, bottom):
                    if v > 0:
                        ax.annotate(f"{v:g}", (xi, b0 + v / 2), ha="center", va="center",
                                    fontsize=9, color="white")
            bottom += vals
        ax.set_xticks(x)
        ax.set_xticklabels(cats)
        ax.legend(fontsize=10, frameon=False, loc=legend_loc)

    elif fig_type == "line":
        for i, s in enumerate(series):
            ax.plot(cats, s["values"], marker="o", markersize=5, linewidth=2.2,
                    label=s.get("name", ""), color=colors[i % len(colors)])
            if show_labels:
                for xi, v in zip(cats, s["values"]):
                    ax.annotate(f"{v:g}", (xi, v), ha="center", va="bottom",
                                fontsize=9, color="#333333", xytext=(0, 5),
                                textcoords="offset points")
        if len(series) > 1:
            ax.legend(fontsize=10, frameon=False, loc=legend_loc)
        plt.setp(ax.get_xticklabels(), rotation=0)

    elif fig_type == "pie":
        vals = series[0]["values"] if series else []
        ax.pie(vals, labels=cats, autopct="%1.1f%%", startangle=90,
               colors=colors[: len(vals)], textprops={"fontsize": 10},
               wedgeprops={"edgecolor": "white", "linewidth": 1.2})
        ax.set_title(spec.get("title") or "", fontsize=14, color="#1a1a1a", pad=12)
        ax.axis("equal")
        fig.savefig(out, dpi=dpi, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return out

    elif fig_type == "scatter":
        for i, s in enumerate(series):
            ax.scatter(s.get("x", cats), s.get("values", []), s=55,
                       label=s.get("name", ""), color=colors[i % len(colors)],
                       alpha=0.85, edgecolors="white", linewidths=0.8)
        if len(series) > 1:
            ax.legend(fontsize=10, frameon=False, loc=legend_loc)

    elif fig_type == "box":
        data = [s["values"] for s in series]
        labels = [s.get("name", "") for s in series]
        bp = ax.boxplot(data, labels=labels, patch_artist=True, widths=0.5,
                        medianprops={"color": "#C55A11", "linewidth": 2})
        for patch, c in zip(bp["boxes"], colors):
            patch.set_facecolor(c)
            patch.set_alpha(0.35)
            patch.set_edgecolor("#1F4E79")

    elif fig_type == "chromatogram":
        # 层析洗脱曲线示意图：由若干高斯峰叠加生成“示意形态”，不是实测数据。
        xr = spec.get("x_range", [0, 20])
        x = np.linspace(float(xr[0]), float(xr[1]), 900)
        y = np.full_like(x, float(spec.get("baseline", 5.0)))
        for pk in spec.get("peaks", []):
            y += float(pk.get("height", 100)) * np.exp(
                -0.5 * ((x - float(pk["center"])) / float(pk.get("width", 1.0))) ** 2
            )
        for band in spec.get("bands", []):
            ax.axvspan(band["x0"], band["x1"], color=band.get("color", "#FFF2CC"),
                       alpha=0.55, zorder=0)
            if band.get("label"):
                ax.annotate(band["label"], ((band["x0"] + band["x1"]) / 2, 0),
                            xytext=(0, 6), textcoords="offset points",
                            ha="center", fontsize=9, color="#7F6000")
        ax.plot(x, y, color=colors[0], linewidth=2.2, zorder=3)
        ax.fill_between(x, 0, y, color=colors[0], alpha=0.12, zorder=2)
        for pk in spec.get("peaks", []):
            if not pk.get("label"):
                continue
            top = float(spec.get("baseline", 5.0)) + float(pk.get("height", 100))
            ax.annotate(pk["label"], (pk["center"], top), ha="center", va="bottom",
                        fontsize=9, color="#1F4E79",
                        xytext=(0, 6), textcoords="offset points")
            ax.plot([pk["center"]], [top], marker="v", color="#C55A11", markersize=5)
        ax.set_xlabel(spec.get("xlabel", "洗脱体积（柱体积 CV）"), fontsize=11)
        ax.set_ylabel(spec.get("ylabel", "UV280 (mAU)"), fontsize=11)
        ax.set_xlim(float(xr[0]), float(xr[1]))
        ax.set_ylim(0, None)

    else:
        raise ValueError(f"不支持的图表类型: {fig_type}")

    _style_axes(ax, spec)
    fig.tight_layout()
    fig.savefig(out, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out


def main() -> int:
    if len(sys.argv) < 2:
        print("用法: python chart_render.py <spec.json>")
        return 2
    spec_path = Path(sys.argv[1]).resolve()
    data = json.loads(spec_path.read_text(encoding="utf-8"))
    figures = data.get("figures") if "figures" in data else [data]
    for spec in figures:
        out = render(spec)
        print(f"PNG: {out}  ({out.stat().st_size / 1024:.0f} KB, 字体: {FONT_NAME})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
