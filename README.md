# md2mdocx

Convert Markdown to manual-style Word documents (docx). Generates clean, professional documentation with cover pages, change history, and table of contents.

output sample： [md2docx-sample](./sample/sample.pdf)

## Features

- Auto-generates cover page, change history, and table of contents
- Headers and footers included
- Supports headings, lists, tables, code blocks, and images
- Inline markup (bold, italic, strikethrough, inline code)

## Usage

### Quick Start (Recommended)

No installation required. Use `npx` or `bunx`:

```bash
npx md2mdocx input.md output.docx
# or
bunx md2mdocx input.md output.docx
```

### Global Installation

```bash
npm install -g md2mdocx
# then
md2mdocx input.md output.docx
```

### Options

| Option | Description | Default |
|--------|-------------|---------|
| `--title` | Product name/title | 製品名 |
| `--subtitle` | Subtitle | マニュアル |
| `--doctype` | Document type | 操作マニュアル |
| `--version` | Version | 1.0.0 |
| `--date` | Creation date | Today's date |
| `--dept` | Department name | 技術開発部 |
| `--docnum` | Document number | DOC-001 |
| `--logo` | Logo image path | None |
| `--company` | Company name | サンプル株式会社 |
| `--theme` | Color theme | blue |
| `--config` | Config file path | (auto-detect) |

#### Config File (YAML)

Options can also be specified in a YAML file. By default, `input.yaml` (same path as the input `.md` file) is automatically loaded.

```yaml
title: "MyApp"
subtitle: "Manual"
doctype: "User Manual"
version: "2.0.0"
date: "January 1, 2025"
dept: "Development"
docnum: "DOC-001"
company: "ABC Corp"
theme: "blue"
# logo: "logo.png"
```

**Priority:** Command-line arguments > Config file > Defaults

#### Theme Options

| Theme | Description | Mermaid Theme |
|-------|-------------|---------------|
| `blue` | Blue accent (default) | default |
| `orange` | Orange accent | neutral |
| `green` | Green accent | forest |

The theme affects:
- Header border line color
- Change history table header background color
- Mermaid diagram color scheme

### Example

```bash
npx md2mdocx manual.md manual.docx \
  --title "MyApp" \
  --doctype "User Manual" \
  --version "2.0.0" \
  --company "ABC Corp" \
  --dept "Development" \
  --theme green
```

## Markdown Syntax

### Change History

You can include change history in your Markdown using HTML comments:

```markdown
<!-- CHANGELOG -->
| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | Jan 1, 2025 | Initial release |
| 1.1.0 | Feb 1, 2025 | Added features |
<!-- /CHANGELOG -->
```

### Images

```markdown
![Alt text](path/to/image.png)
```

You can also specify size using HTML img tags:

```html
<img src="icon.png" width="24" height="24">
```

### Line Breaks and Page Breaks

```markdown
This is line 1.<br>This is line 2.

<pagebreak>

This content appears on a new page.
```

| Syntax | Effect |
|--------|--------|
| `<br>` | Force line break within paragraph |
| `<pagebreak>` | Force page break |

### HTML Comment Controls

For use with existing Markdown files (e.g., README), you can use HTML comments to control parsing:

```markdown
# Project Title (skipped)

Badges and links here...

<!-- md2mdocx:start -->

# Documentation starts here

Content to be converted...

<!-- md2mdocx:pagebreak -->

# New page

<!-- md2mdocx:br -->

Empty line inserted above.

<!-- md2mdocx:end -->

## Appendix (skipped)
```

| Comment | Effect |
|---------|--------|
| `<!-- md2mdocx:start -->` | Start parsing from this line (skip content before) |
| `<!-- md2mdocx:end -->` | End parsing at this line (skip content after) |
| `<!-- md2mdocx:pagebreak -->` | Force page break |
| `<!-- md2mdocx:br -->` | Insert empty line |

### Mermaid Diagrams

Mermaid code blocks are automatically converted to PNG images using [Kroki](https://kroki.io) API:

~~~markdown
```mermaid
graph TD
    A[Start] --> B{Decision}
    B -->|Yes| C[Process A]
    B -->|No| D[Process B]
```
~~~

- Requires internet connection for rendering
- Images are displayed at original size, scaled down if exceeding page width
- If rendering fails, a warning message is shown in the document

## Notes

### Opening Generated Documents

When opening the generated Word document, you may see a dialog:

> "This document contains fields that may refer to other files. Do you want to update the fields in this document?"

This is because the document contains fields such as table of contents (TOC) and page numbers. Select **Yes** to update them to the latest state.

## License

MIT
