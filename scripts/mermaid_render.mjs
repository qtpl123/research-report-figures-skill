#!/usr/bin/env node
/**
 * Mermaid -> PNG/SVG 渲染器（本机 Chrome + Codex 自带 Playwright，不下载 Chromium）
 *
 * 用法:
 *   node mermaid_render.mjs <input.mmd> <output.png> [选项]
 *
 * 选项:
 *   --width <px>       输出图宽度，默认 1700（按这个宽度等比缩放）
 *   --scale <n>        像素密度倍数，默认 3（3 倍 ≈ 300dpi 打印质量）
 *   --svg <path>       同时导出矢量 SVG
 *   --font-size <px>   图中字号，默认 20
 *   --theme <name>     mermaid 主题: base|default|neutral|dark，默认 base
 *   --max-height <px>  限制图的最大高度（默认 2400 CSS px），过高时按比例缩小并提示
 *   --quiet            安静模式
 */

import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

// 依赖解析：优先用环境变量，其次按“脚本所在目录的上一级 /node/node_modules”约定查找。
// 可用环境变量覆盖：
//   CODEX_NODE_MODULES  依赖 playwright 的 node_modules 目录
//   MERMAID_JS          mermaid.min.js 的完整路径
const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(SCRIPT_DIR, "..");
const DEFAULT_NODE_MODULES =
  process.env.CODEX_NODE_MODULES || path.join(PROJECT_ROOT, "node", "node_modules");
const MERMAID_JS = process.env.MERMAID_JS ||
  path.join(DEFAULT_NODE_MODULES, "mermaid", "dist", "mermaid.min.js");

function loadPlaywright() {
  const roots = [
    process.env.CODEX_NODE_MODULES,
    DEFAULT_NODE_MODULES,
  ].filter(Boolean);
  const errors = [];
  for (const root of roots) {
    try {
      const req = createRequire(path.join(root, "noop.js"));
      return req("playwright");
    } catch (err) {
      errors.push(`${root}: ${err.message}`);
    }
  }
  try {
    return createRequire(import.meta.url)("playwright");
  } catch (err) {
    errors.push(`local: ${err.message}`);
  }
  throw new Error("找不到 playwright：\n" + errors.join("\n"));
}

function parseArgs(argv) {
  const opts = {
    input: null,
    output: null,
    width: 1700,
    maxHeight: 2400,
    scale: 3,
    svg: null,
    fontSize: 20,
    theme: "base",
    quiet: false,
  };
  const positional = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const next = () => argv[++i];
    if (a === "--width") opts.width = Number(next());
    else if (a === "--scale") opts.scale = Number(next());
    else if (a === "--svg") opts.svg = next();
    else if (a === "--font-size") opts.fontSize = Number(next());
    else if (a === "--theme") opts.theme = next();
    else if (a === "--max-height") opts.maxHeight = Number(next());
    else if (a === "--quiet") opts.quiet = true;
    else positional.push(a);
  }
  opts.input = positional[0];
  opts.output = positional[1];
  return opts;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (!opts.input || !opts.output) {
    console.error("用法: node mermaid_render.mjs <input.mmd> <output.png> [--width 1700] [--scale 3] [--svg out.svg]");
    process.exit(2);
  }
  const input = path.resolve(opts.input);
  const output = path.resolve(opts.output);
  if (!fs.existsSync(input)) {
    console.error(`找不到输入文件: ${input}`);
    process.exit(2);
  }
  if (!fs.existsSync(MERMAID_JS)) {
    console.error(`找不到 mermaid 库: ${MERMAID_JS}`);
    process.exit(2);
  }

  const code = fs.readFileSync(input, "utf8");
  const mermaidSource = fs.readFileSync(MERMAID_JS, "utf8");
  fs.mkdirSync(path.dirname(output), { recursive: true });

  const { chromium } = loadPlaywright();
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  let result;
  try {
    const context = await browser.newContext({
      viewport: { width: 2400, height: 1600 },
      deviceScaleFactor: opts.scale,
    });
    const page = await context.newPage();
    page.on("pageerror", (e) => console.error("[pageerror]", e.message));
    await page.setContent(
      '<!doctype html><html><head><meta charset="utf-8"></head>' +
        '<body style="margin:0;background:#ffffff">' +
        '<div id="stage" style="display:inline-block;background:#ffffff;padding:16px"></div>' +
        "</body></html>"
    );
    await page.addScriptTag({ content: mermaidSource });

    result = await page.evaluate(
      async ({ code, fontSize, theme, targetWidth, maxHeight }) => {
        const mermaid = window.mermaid;
        mermaid.initialize({
          startOnLoad: false,
          theme,
          securityLevel: "loose",
          fontFamily:
            '"Microsoft YaHei","Segoe UI","Helvetica Neue",Arial,"Noto Sans CJK SC",sans-serif',
          themeVariables: {
            fontFamily:
              '"Microsoft YaHei","Segoe UI","Helvetica Neue",Arial,"Noto Sans CJK SC",sans-serif',
            fontSize: `${fontSize}px`,
            primaryColor: "#DEEAF6",
            primaryBorderColor: "#1F4E79",
            primaryTextColor: "#1a1a1a",
            lineColor: "#1F4E79",
            tertiaryColor: "#F2F7FB",
          },
          // wrappingWidth 决定节点内文字换行的宽度（默认 200 会让中文标签碎成很多行）
          flowchart: {
            htmlLabels: true,
            curve: "basis",
            useMaxWidth: false,
            wrappingWidth: 420,
            padding: 14,
          },
          sequence: { useMaxWidth: false },
          gantt: { useMaxWidth: false },
          er: { useMaxWidth: false },
        });

        const { svg } = await mermaid.render("graph-" + Date.now(), code);
        const stage = document.getElementById("stage");
        stage.innerHTML = svg;
        const el = stage.querySelector("svg");
        if (!el) return { ok: false, error: "渲染结果中没有 SVG 元素" };

        const vb = el.getAttribute("viewBox");
        let naturalW = 0;
        let naturalH = 0;
        if (vb) {
          const parts = vb.split(/[\s,]+/).map(Number);
          naturalW = parts[2];
          naturalH = parts[3];
        }
        if (!naturalW || !naturalH) {
          const box = el.getBBox();
          naturalW = box.width;
          naturalH = box.height;
        }

        let width = Math.min(targetWidth, Math.max(naturalW, 400));
        let height = (naturalH / naturalW) * width;
        if (height > maxHeight) {
          width = width * (maxHeight / height);
          height = maxHeight;
        }
        el.setAttribute("width", String(width));
        el.setAttribute("height", String(height));
        el.style.width = `${width}px`;
        el.style.height = `${height}px`;
        el.style.maxWidth = "none";
        el.removeAttribute("style");
        el.setAttribute("style", `width:${width}px;height:${height}px;max-width:none`);

        return {
          ok: true,
          width: Math.round(width),
          height: Math.round(height),
          naturalWidth: Math.round(naturalW),
          naturalHeight: Math.round(naturalH),
          svg: el.outerHTML,
        };
      },
      { code, fontSize: opts.fontSize, theme: opts.theme,
        targetWidth: opts.width, maxHeight: opts.maxHeight }
    );

    if (!result.ok) throw new Error(result.error);

    const stage = page.locator("#stage");
    await stage.screenshot({ path: output, omitBackground: false });

    if (opts.svg) {
      const svgPath = path.resolve(opts.svg);
      fs.mkdirSync(path.dirname(svgPath), { recursive: true });
      fs.writeFileSync(
        svgPath,
        '<?xml version="1.0" encoding="UTF-8"?>\n' + result.svg,
        "utf8"
      );
      if (!opts.quiet) console.log(`SVG: ${svgPath}`);
    }
  } finally {
    await browser.close();
  }

  if (!opts.quiet) {
    const size = fs.statSync(output).size;
    const ratio = result.height / result.width;
    console.log(
      `PNG: ${output}\n  输出尺寸: ${result.width}x${result.height} CSS px @${opts.scale}x ` +
        `(${result.width * opts.scale}x${result.height * opts.scale} 像素, ${(size / 1024).toFixed(0)} KB)`
    );
    if (ratio > 3.2) {
      console.warn(
        `  ⚠ 图形过高（高宽比 ${ratio.toFixed(1)}:1），插入 Word 后会又窄又长、难以阅读。` +
          `建议改成左右布局（flowchart LR）或拆成两张图。`
      );
    } else if (result.width / result.height > 6) {
      console.warn(
        `  ⚠ 图形过宽（宽高比 ${(result.width / result.height).toFixed(1)}:1），` +
          `插入 Word 后文字会很小。建议改成上下布局（flowchart TB）或拆成两张图。`
      );
    }
  }
}

main().catch((err) => {
  console.error("渲染失败: " + (err && err.stack ? err.stack : err));
  process.exit(1);
});
