# DOCX Navigator

Merge multiple Word (`.docx`) files into a single document with:

* A **clickable menu** linking to each document’s start
* **“Back to menu”** links at the top of each section
* Automatic **page breaks** between merged documents
* Preserves original formatting, images, and styles

## Features

* **Auto-discovery**: Automatically finds all `.docx` files in current directory
* Groups files by the `<category>_<document name>.docx` pattern
* Internal bookmarks + hyperlinks (no MS Word required)
* Cross-platform (uses `python-docx` + `docxcompose`)

---

## Quickstart

> Requires: [uv](https://github.com/astral-sh/uv) installed.

```bash
# Automatically merges all .docx files in current directory
uv run --with . docx-navigator

# Or specify files explicitly
uv run --with . docx-navigator --inputs file1.docx file2.docx
```

No configuration needed! Just run the command in a directory with `.docx` files.
Output will be `all_documents.docx` by default.

---

## CLI Options

```bash
uv run --with . docx-navigator [OPTIONS]
```

**Main Options**

| Option           | Default              | Description                                                                            |
| ---------------- | -------------------- | -------------------------------------------------------------------------------------- |
| `--inputs`       | *auto-detect*        | Explicit list of `.docx` files to merge. If not provided, uses all `.docx` files in current directory. |
| `--output`       | `all_documents.docx` | Output file path/name.                                                                |
| `--menu-title`   | `Menu`               | Heading text for the clickable menu.                                                  |
| `--back-label`   | `Back to menu`       | Label for the backlink at the start of each section.                                  |
| `--category-sep` | `_`                  | Separator between category and document name in filenames.                            |
| `--dry-run`      | off                  | Preview what would be merged without writing output file.                             |

> The tool groups files by the `<category>_<document name>.docx` pattern using `--category-sep`. Anything before the first separator is treated as the category.

### Examples

**1) Basic usage - merge all .docx files in current directory**

```bash
uv run docx-navigator
```

**2) Merge specific files**

```bash
uv run docx-navigator \
  --inputs "Finance_Q1.docx" "Finance_Q2.docx" \
  --output "Finance_Reports.docx"
```

**3) Preview without creating output**

```bash
uv run docx-navigator --dry-run
```

**4) Customize menu and labels**

```bash
uv run docx-navigator \
  --menu-title "Document Index" \
  --back-label "Return to Index"
```

---

## Example

**Input files**

```
Finance_Quarterly Report Q1.docx
Finance_Quarterly Report Q2.docx
HR_Employee Handbook.docx
HR_Payroll Guidelines.docx
Marketing_Brand Guidelines.docx
Marketing_Campaign Plan 2025.docx
```

**Output menu structure**

```
Finance
  Quarterly Report Q1
  Quarterly Report Q2
HR
  Employee Handbook
  Payroll Guidelines
Marketing
  Brand Guidelines
  Campaign Plan 2025
```

* Clicking an entry in the menu jumps to the corresponding document section.
* Each section starts with a **Back to menu** link for quick navigation.
* Page breaks separate each appended document.

---

## Installation

Ensure you have [uv](https://github.com/astral-sh/uv) installed.
You can either install the CLI permanently or run it directly from the project directory.

### Option 1 — Run without installing

```bash
uv run --with . docx-navigator
```

The `--with .` flag tells `uv` to include the current project in the temporary environment before running.

### Option 2 — Install locally

```bash
uv pip install -e .
docx-navigator
```

This makes the `docx-navigator` command available anywhere in your environment.
