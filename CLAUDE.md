# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

md2mdocx is a CLI tool that converts Markdown files to manual-style Word documents (docx). It generates professional documentation with cover pages, change history, and table of contents.

## Commands

```bash
# Run conversion
bun md2mdocx.js input.md output.docx

# Run with options
bun md2mdocx.js input.md output.docx --title "Title" --theme blue

# Install dependencies
bun install
```

## Architecture

Single-file architecture (`md2mdocx.js`):

1. **MarkdownParser class** - Parses Markdown into element objects (heading, paragraph, list, table, code, image, etc.)
2. **parseInlineMarkup()** - Processes inline formatting (bold, italic, code, images, line breaks)
3. **Section generators** - `createCoverSection()`, `createHistorySection()`, `createTOCSection()`
4. **convertElements()** - Transforms parsed elements into docx Paragraph/Table objects
5. **Mermaid support** - Uses Kroki API to render diagrams as PNG images

Key dependencies:
- `docx` - Word document generation
- `yaml` - Config file parsing

## Page Break Formats

The parser recognizes these page break syntaxes:
- `<pagebreak>` or `<pagebreak />`
- `<!-- md2mdocx:pagebreak -->`
- `<div style="page-break-before:always"></div>`
- `---` (when `--hr-pagebreak true`, default)

## Config

Options can be specified via:
1. Command-line arguments (highest priority)
2. YAML config file (auto-loads `input.yaml` from same path as input.md)
3. Default values

## Tooling

Use Bun instead of Node.js/npm:
- `bun` instead of `node`
- `bun install` instead of `npm install`
