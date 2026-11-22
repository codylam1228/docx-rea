# docx-rea

**Zero-Format-Loss Word (`.docx`) Automation for AI Agents & Web LLMs**

`docx-rea` brings an **Inspect-First $\rightarrow$ Anchor-Verified Patching $\rightarrow$ Deep XML Clone** workflow (inspired by [`morluto/rea`](https://github.com/morluto/rea)) directly to Microsoft Word's underlying OpenXML (`.docx`) layer. Instead of generating documents from scratch—which destroys complex page margins, custom fonts, table column widths, borders, and headers/footers—or flooding an LLM's context window with tens of thousands of lines of raw XML, `docx-rea` exposes compact structural coordinates and performs surgical, format-preserving XML mutations.

---

## Key Features

* **100% Original Format Preservation**: Modifies existing `.docx` templates in-place at the OOXML level, preserving page geometry, styles, character run properties (`<w:rPr>`), cell widths (`<w:tcW>`), borders, shading, list numbering (`<w:numPr>`), headers, and footers.
* **Compact Structural Coordinates (>95% Token Reduction)**: Summarizes documents into clean node addresses (`p[12]`, `p[12].r[1]`, `t[0].r[2].c[1]`, `hf[0].header.p[0]`) with font/style metadata.
* **Atomic Anchor Verification**: Every edit, clone, or deletion requires both a target coordinate and an exact `anchor_text` substring. If indices shift, the operation aborts before mutating the file.
* **Embedded Media & Equation Protection**: Automatically isolates existing Word drawings/images (`<w:drawing>` $\rightarrow$ `[IMG:k]`) and native math (`<m:oMath>` $\rightarrow$ `[MATH:k]`) so text edits never corrupt embedded objects.
* **Native Word Equation Builder (`$LaTeX$` $\rightarrow$ `<m:oMath>`)**: Converts inline LaTeX expressions (fractions `\frac`, `\dfrac`, `\tfrac`, roots `\sqrt`, superscripts `^`, subscripts `_`, Greek letters, and operators) directly into native Word OMML equations.
* **Deep XML Node Cloning**: Clones table rows (`<w:tr>`) or styled paragraphs (`<w:p>`) via `copy.deepcopy()` and populates new cells/text while keeping all borders, column widths, and typography intact.

---

## Two Ways to Run

### Mode 1: Zero-Backend Web Studio (GitHub Pages + Pyodide WASM)

Run the entire Python OOXML engine (`lxml` + `python-docx`) directly inside the browser using **Pyodide WebAssembly** and an in-memory filesystem (`MEMFS`)—with **zero backend servers** and **zero API keys required**.

#### 1-Minute GitHub Pages Deployment
1. Upload the three static files from `docx-rea-ghpages` (`index.html`, `worker.js`, `xml_engine.py`) to the root of a GitHub repository.
2. Go to **Settings $\rightarrow$ Pages $\rightarrow$ Build and deployment $\rightarrow$ Source: Deploy from a branch (`main` / `/root`)** and save.
3. Open `https://<your-username>.github.io/<repo-name>/`.

#### Zero-API-Key Workflow
1. **Load Template**: Click **1. Load `.docx` Template** to write the document into the browser's Pyodide `MEMFS` sandbox and inspect its live layout via `docx-preview`.
2. **Copy Structured Prompt**: Enter your content instructions (e.g., *"Generate 10 new algebra transposition questions matching the existing layout"*) and click **Copy Complete Prompt**.
3. **Generate JSON in Any Web LLM**: Paste the prompt into **Gemini Web** (`gemini.google.com`) or **Grok Web** (`grok.com`). The model outputs a declarative JSON plan (`delete_targets`, `clones`, `edits`).
4. **Apply & Download**: Paste the JSON back into the studio, click **Apply JSON**, verify the live visual preview and Compact Diff, and download your new `.docx`.

*(Also registers `docx_apply_plan` via **WebMCP** `document.modelContext` when enabled in Chrome flags.)*

---

### Mode 2: Local Python CLI & FastMCP Server

Connect `docx-rea` directly to local MCP-capable AI agents (**Claude Code**, **Cursor**, **Gemini CLI**) using [`uv`](https://docs.astral.sh/uv/).

#### Installation
```bash
cd docx-rea
uv sync
```

#### CLI Usage
```bash
# Inspect compact structural coordinates
uv run docx-rea inspect template.docx

# Include run-level formatting coordinates and filter by prefix
uv run docx-rea inspect template.docx --runs --prefix "t[0]"

# Apply a JSON batch edit file
uv run docx-rea batch-edit template.docx edits.json -o output.docx

# Render visual preview pages (requires LibreOffice)
uv run docx-rea preview output.docx
```

#### MCP Server Configuration
Add the following entry to your MCP client configuration (`claude_desktop_config.json`, `.cursor/mcp.json`, or `~/.gemini/settings.json`):

```json
{
  "mcpServers": {
    "docx-rea": {
      "command": "uv",
      "args": [
        "--directory",
        "/ABSOLUTE/PATH/TO/docx-rea",
        "run",
        "docx-rea-mcp"
      ]
    }
  }
}
```

#### Exposed MCP Tools
| Tool | Description |
| :--- | :--- |
| `docx_inspect` | Returns the compact structural coordinate map, style summaries, and `[MATH:k]` / `[IMG:k]` markers. |
| `docx_batch_edit` | Atomically updates multiple paragraphs, runs, table cells, or headers/footers with anchor verification. |
| `docx_clone_node` | Deep-clones a table row (`t[0].r[1]`) or styled paragraph (`p[5]`) and populates new items (`after` or `replace`). |
| `docx_preview` | Converts `.docx` to PDF/PNG via LibreOffice headless for visual layout verification. |

---

## Coordinate Addressing Reference

| Coordinate Syntax | Target Element |
| :--- | :--- |
| `p[0]` | First body paragraph |
| `p[0].r[1]` | Second formatted text run inside `p[0]` |
| `t[0].r[2]` | Third row of the first table (used as a template for `clones` or `delete_targets`) |
| `t[0].r[2].c[1]` | Second cell of row 2 in table 0 |
| `t[0].r[2].c[1].p[1]` | Second paragraph inside a multi-paragraph table cell |
| `hf[0].header.p[0]` | First paragraph in section 0's header (also supports `footer`, `first_header`, `even_header`) |

---

## Unified JSON Plan Schema

Both the Web Studio and `xml_engine.apply_plan_docx()` accept a single declarative JSON payload:

```json
{
  "delete_targets": [
    {
      "target": "t[0].r[1]",
      "anchor_text": "Old row text to remove"
    }
  ],
  "clones": [
    {
      "target": "t[0].r[0]",
      "anchor_text": "Template row text",
      "mode": "replace",
      "items": [
        ["1. Make $x$ the subject of $\\frac{3x + a}{x - b} = k$."],
        ["2. Solve $x^2 - 5x + 6 = 0$."]
      ]
    }
  ],
  "edits": [
    {
      "target": "hf[0].header.p[0]",
      "anchor_text": "Worksheet 1",
      "new_text": "Worksheet 2 — Algebra Practice"
    }
  ]
}
```

---

## Repository Structure

```text
├── index.html          # Zero-backend Web Studio UI, Prompt Builder & WebMCP registration
├── worker.js           # Pyodide Web Worker managing MEMFS virtual disk & Python execution
├── xml_engine.py       # Pure-Python OOXML parser, anchor guard, deep cloner & LaTeX-to-OMML builder
├── pyproject.toml      # Python 3.11+ package & CLI/MCP entrypoints (local mode)
└── src/docx_rea/
    ├── xml_engine.py   # Core OOXML engine
    ├── mcp_server.py   # FastMCP server exposing docx_* tools
    └── cli.py          # Typer CLI for terminal inspection & patching
```