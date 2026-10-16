**English** · [中文](README.zh-CN.md)

# research-report-figures

An Agent Skill for generating **professional research reports with figures, delivered as Word documents**.

It turns the pipeline "topic → research → figures → Word report → layout check" into a repeatable process, so an AI assistant produces `.docx` files with a table of contents, figure numbers, captions, sources and references — then renders every page to verify the layout.

## The problem it solves

Getting an AI to *write* a report is easy. Getting it to **illustrate it credibly** is hard. The usual failure modes: a flow chart sketch passed off as a molecular structure, a plausible-looking fabricated citation, chart data with no traceable origin.

This Skill blocks those failures before generation, with three hard rules:

1. **Molecular structures, chemical structures, biological macromolecules (e.g. antibodies) and real experimental data must be cited from genuine sources** — self-drawn substitutes and flow-chart stand-ins are forbidden.
2. **Fabricating citations, data, or image sources is forbidden.** If a source cannot be found, write "no verifiable source found".
3. **Every real image must carry its original URL and license.** If either is missing, the build script refuses to generate the document.

## Three-tier figure strategy

| Tier | Applies to | Tooling |
| --- | --- | --- |
| Tier 1 · Technical schematics | Process flows, technical roadmaps, hierarchies, data comparisons | Mermaid / Graphviz / matplotlib |
| Tier 2 · Real images | Equipment photos, molecular structures, figures from literature, real spectra | Web retrieval + source registration |
| Tier 3 · AI-generated images | Covers and decorative illustrations only; off by default | Must be captioned "AI-generated" |

## Installation

Place this repository in your Agent's skills directory, for example:

```
~/.workbuddy/skills/research-report-figures/     # user-level skills dir (WorkBuddy / Claude Code style)
```

Directory layout:

```
research-report-figures/
├── SKILL.md                  # Main skill file (workflow and hard rules)
├── references/
│   └── content-syntax.md     # content.md syntax reference
├── scripts/                  # 8 utility scripts
└── examples/                 # Example report
```

## Dependencies

| Dependency | Used for | Notes |
| --- | --- | --- |
| Python 3 | Build and chart scripts | Requires `python-docx`, `Pillow`, `matplotlib`, `pypdfium2` |
| Node.js | Mermaid rendering | Requires `playwright`, `mermaid` |
| PowerShell 7+ | Graphviz rendering, Word-to-PDF export | Windows |
| Graphviz | Graph rendering | `dot` must be on PATH |
| Microsoft Word | Word-to-PDF export (QA step) | Only needed on Windows with Word installed |

Install the Python dependencies:

```bash
pip install python-docx Pillow matplotlib pypdfium2
```

## Path configuration (important)

`SKILL.md` uses `<TOOLS>` and `<PYTHON>` as placeholders. Before first use, replace them with real paths on your machine (or let the assistant substitute them at runtime):

- `<TOOLS>` → absolute path to this repository's `scripts/` directory
- `<PYTHON>` → path to a Python interpreter with the dependencies above installed
- `node` / `pwsh` → the corresponding executables on your machine

`scripts/mermaid_render.mjs` accepts environment variables for dependency locations:

| Variable | Meaning | Default |
| --- | --- | --- |
| `CODEX_NODE_MODULES` | `node_modules` directory containing playwright | `<TOOLS>/../node/node_modules` |
| `MERMAID_JS` | Full path to `mermaid.min.js` | `<CODEX_NODE_MODULES>/mermaid/dist/mermaid.min.js` |

## Usage

Trigger it directly in any Skill-capable AI assistant, for example:

> Write a "Sodium-ion battery technology roadmap" research report with figures — include a process flow diagram and data charts.

The assistant follows the standard workflow in `SKILL.md`: web research → figure planning → render and register sources → assemble the Word document → page-by-page layout check → delivery.

Output layout:

```
<working dir>/reports/<topic>/
├── content.md            # Report body
├── assets/               # Images, diagram sources, manifest.json
├── assets/qa/            # Layout QA artifacts (internal)
└── <report title>.docx   # Final deliverable
```

## Example

`examples/` contains a complete example report — see the output for yourself.

## Notes

- This Skill is **environment-agnostic**: all paths are configured via placeholders and environment variables, bound to no specific machine.
- The web research step requires your assistant to have network access, granted per-request under a sandbox.
- License details for real images are the user's responsibility to verify; check the original source's terms before republishing.

## License

[MIT](LICENSE)
