# md2mdocx

Convert Markdown to manual-style Word documents (docx). Generates clean, professional documentation with cover pages, change history, and table of contents.

## Features

- Auto-generates cover page, change history, and table of contents
- Headers and footers included
- Supports headings, lists, tables, code blocks, and images
- Inline markup (bold, italic, strikethrough, inline code)

## Installation

```bash
npm install -g md2mdocx
```

## Usage

### Basic

```bash
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

### Example

```bash
md2mdocx manual.md manual.docx \
  --title "MyApp" \
  --doctype "User Manual" \
  --version "2.0.0" \
  --company "ABC Corp" \
  --dept "Development"
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

## License

MIT
