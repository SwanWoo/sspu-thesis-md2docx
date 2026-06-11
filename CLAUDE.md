# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SSPU graduation thesis tool. Converts Markdown thesis drafts from `input/*.md` into formatted `.docx` Word documents by injecting content into a university-provided Word template at the XML level.

Active code: `thesis2docx.py`. Input-side AI workflow: `.claude/skills/thesis-to-markdown.md` (Claude Code) / `AGENTS.md` (Codex).

## Input-side Workflow (IMPORTANT)

When the user wants to **create, convert, or restructure thesis content** (e.g. "写论文", "转 markdown", "帮我整理", "从 docx 转", "起草毕业论文", pastes raw text, or provides a .docx/.pdf/.txt path), **invoke the `thesis-to-markdown` skill FIRST** before doing anything else. That skill handles intent detection, cover-field collection, syntax normalization, and optional humanizer chaining. Do NOT hand-write Markdown without it — the parser is custom and unforgiving.

Codex agents: see `AGENTS.md` at repo root for the equivalent instructions.

## Commands

```bash
uv sync                    # Install dependencies (pymupdf)
uv run python thesis2docx.py   # Run conversion → outputs to output/ directory
```

No build step, no tests, no linting. Requires `xelatex` (TeX Live) on PATH for formula rendering; falls back to raw LaTeX text if absent.

## Architecture

```
final_paper.md → thesis2docx.py → output/*.docx
                                (reads word/document.xml as template)
```

The `.docx` template is a ZIP of OpenXML files. `thesis2docx.py` manipulates the XML directly with `xml.etree.ElementTree`. The template is fully extracted into `template/` (preserving the ZIP internal structure). The output ZIP is reassembled from these files plus the modified `document.xml`.

### Pipeline (main function)

1. `parse_markdown()` — splits `.md` into typed blocks (heading, paragraph, bullet, numbered, table, image, formula, codeblock, blockquote)
2. `inject_cover_data()` — fills P4-P13 with cover info (title, English title, student ID, name, etc.)
3. `replace_abstracts()` — cleans template formatting instructions from P31-P42, inserts Chinese/English abstract text and keywords
4. `build_toc_entries()` — generates PAGEREF-field-based TOC from body headings, replaces P44+ area
5. Body content loop — converts Markdown blocks to Word paragraphs and appends after the last section break
6. `_update_rels_file()` — patches `document.xml.rels` with image/formula relationship entries
7. ZIP assembly — writes modified XML into new `.docx`, patches `settings.xml` with `updateFields=true`

## Template Paragraph Positions (2026届新模板, hardcoded into document.xml)

```
P0-P3:    Cover page header/instructions/images
P4-P13:   Cover page fields (题目, 英文题目, 学号, 姓名, 班级, 专业, 学部(院), 入学时间, 指导教师, 日期)
          All use af2 (Title) style, theme fonts (minorHAnsi). No ddList dropdowns (WPS annotations instead).
P14:      Section break (cover → declaration, continuous)
P15-P29:  Declaration page
P30:      Section break (declaration → abstract, nextPage)
P31:      Chinese thesis title (黑体, sz=36)
P32:      "摘要" label (黑体, bold, sz=32)
P33-P35:  Chinese abstract body (sz=24, firstLine=480)
P36:      Chinese keywords (关键词：黑体, keywords 宋体, sz=24)
P37:      Formatting note (cleaned/emptied)
P38:      Formatting note (cleaned/emptied, no page break — CN/EN separated by sectPr at P43)
P39:      English title (af2, TNR+宋体, sz=36, bold, centered)
P40:      "ABSTRACT" (af2, TNR+宋体, sz=32, bold, centered)
P41:      English abstract body (af2, TNR+宋体, b=0, sz=24, jc=both)
P42:      English keywords (af2, TNR+宋体, sz=24, jc=both)
P43:      Section break (abstract → TOC, nextPage)
P44:      TOC title ("目录")
P45-P55:  TOC entries (TOC1/TOC2/TOC3 styles, replaced dynamically)
P65:      Section break (TOC → body, nextPage)
P66+:     Body content (all replaced from Markdown)
```

These indices are **fragile** — template changes require re-indexing.

## Style Mapping (2026届新模板, from word/styles.xml)

| Element | Style | Notes |
|---|---|---|
| H1 | `'12'` | basedOn heading 1 (style id `'1'`), centered, ea=黑体, sz=32 |
| H2 | `'afc'` | basedOn 12, left-aligned, ea=黑体, sz=28 |
| H3 | `'a'` / `'afa'+outlineLvl` | ea=黑体, sz=24, numPr(numId=1, ilvl=2) |
| Body | `'afa'` | basedOn a0, firstLineChars=200, firstLine=480 (2-char indent) |
| Table caption | `'a4'` | Auto-detected in `build_body_paragraph()` |
| Figure caption / blockquote | `'a4'` | Centered, 黑体, sz=21 |
| Image caption | `'a4'` | After downloaded image |
| Cover fields | `'af2'` | Title style, all P4-P13, theme fonts (minorHAnsi) |
| English abstract | `'af2'` | TNR+宋体, b=0 on body runs |
| Bullet | `'afa'` + numId=2 | |
| Numbered | `'afa'` + numId=1 | |
| Code block | `'afa'` | Times New Roman, no indent |
| TOC L1/L2/L3 | `'TOC1'`/`'TOC2'`/`'TOC3'` | 黑体 sz=28 / 宋体 sz=24 / 楷体 sz=24 |

**Important:** Style `af2` (Title) inherits bold from its definition. Runs that need to be non-bold must explicitly set `<w:b val="0"/>` (use `unbold=True` parameter in `make_text_run()`).

## Key Patterns

- **`_clean_paragraph_keep_ppr(p)`** — strips all children except `pPr` from a template paragraph, also removes illegal `<w:rPr>` elements that the template leaves inside `<w:pPr>`
- **Namespace registration order matters** — `register_namespaces()` registers 17 namespaces; `r:` prefix registered last for relationship references
- **`rels_counter` dict** — `{'count': 0, '_rels_entries': []}` passed mutably to track image/formula relationships; consumed by `_update_rels_file()`
- **Formula rendering** — XeLaTeX compiles LaTeX to PDF, PyMuPDF converts to PNG at 1200 DPI. Numbered formulas use center-tab + right-tab layout for `(2-1)` style numbering. Un-numbered formulas use `jc=center`.
- **Media files accumulate** — `template/word/media/` grows with each run (formula PNGs + downloaded images). Gitignored via `.gitignore`.
- **Image path resolution** — `build_image_block()` resolves paths in priority order: media cache → local file (relative to project root or `file://` URI) → HTTP(S) URL direct → proxy retry. Relative paths like `images/foo.png` are resolved against `os.path.dirname(os.path.abspath(__file__))`.
- **Bookmark counter** — global `_bookmark_counter`, scans template for max existing `_Toc` bookmark ID
- **Cover P5 (英文题目) and P10 (学部院)** — font sizes preserved from template (sz=28 for English title value, sz=28 for college value)
- **Cover font format** — 2026届新模板封面所有字段统一使用 theme font (asciiTheme/hAnsiTheme/cstheme = minorHAnsi)，不再使用 hAnsi='黑体'

## Unit Conventions

- Word half-points: sz=44=22pt, sz=36=18pt, sz=32=16pt, sz=28=14pt, sz=24=12pt, sz=21=10.5pt
- Indent: `firstLine=480` twips = 2 Chinese chars
- EMU for images: `pixels / DPI * 72 * 9525`

## Important Files

- `thesis2docx.py` — the converter (all active code)
- `input/*.md` — thesis source Markdown files (default: first `.md` alphabetically, or pass path as arg)
- `images/` — local image assets, named `图x-x-描述.png` by convention
- `.claude/skills/thesis-to-markdown.md` — input-side AI skill (Claude Code)
- `AGENTS.md` — input-side AI instructions (Codex, sync pair with above skill)
- `.claude/skills/humanizer-academic-zh.md` — de-AI-tone polishing skill
- `template/word/document.xml` — template body XML (read-only reference)
- `template/word/styles.xml` — template style definitions
- `pyproject.toml` — uv project config (dependency: `pymupdf>=1.24`)
