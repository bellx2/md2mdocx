#!/usr/bin/env node
/**
 * Markdown to Word Converter
 * Manual Template Format
 *
 * Usage: bun md2docx.js input.md output.docx [options]
 * Options:
 *   --title "Product Name"
 *   --subtitle "Manual"
 *   --doctype "Operation Manual"
 *   --version "1.0.0"
 *   --date "January 1, 2024"
 *   --dept "Technical Development"
 *   --docnum "DOC-001"
 *   --logo "logo.png"
 *   --company "Company Name"
 *   --theme "blue|orange|green"  (accent color theme)
 *   --config "config.yaml"       (config file path, defaults to input.yaml)
 *
 * Config file (YAML):
 *   title: "Product Name"
 *   subtitle: "Manual"
 *   doctype: "Operation Manual"
 *   version: "1.0.0"
 *   date: "January 1, 2024"
 *   dept: "Technical Development"
 *   docnum: "DOC-001"
 *   logo: "logo.png"
 *   company: "Company Name"
 *   theme: "blue"
 *
 * Priority: Command line args > Config file > Default values
 *
 * HTML comment controls:
 *   <!-- md2mdocx:start -->  Start parsing from this line (skip file header)
 *   <!-- md2mdocx:end -->    End parsing at this line (skip file footer)
 *   <!-- md2mdocx:pagebreak --> Page break
 *   <!-- md2mdocx:br -->     Line break
 *
 * Other options:
 *   --hr-pagebreak true/false  Treat horizontal rules (---) as page breaks (default: true)
 *   --save-config "config.yaml" Save current settings (including defaults) to YAML file
 */

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
        AlignmentType, PageNumber, BorderStyle, WidthType, HeadingLevel, PageBreak,
        TableOfContents, ShadingType, LevelFormat, ImageRun } = require('docx');
const fs = require('fs');
const path = require('path');
const YAML = require('yaml');

// ===== 設定 =====
const PAGE_WIDTH = 11906;
const MARGIN = 1440;
const CONTENT_WIDTH = PAGE_WIDTH - (MARGIN * 2);
const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: "000000" };

// ===== テーマカラー設定 =====
const THEME_COLORS = {
  blue: {
    headerBorder: "2F4F76",    // 濃い青（ヘッダー下線）
    tableHeader: "538DD4",     // 明るい青（変更履歴テーブルヘッダー）
    mermaid: "default"         // Mermaidテーマ
  },
  orange: {
    headerBorder: "B45F06",    // 濃いオレンジ（ヘッダー下線）
    tableHeader: "F6B26B",     // 明るいオレンジ（変更履歴テーブルヘッダー）
    mermaid: "neutral"         // Mermaidテーマ（オレンジに近いグレー系）
  },
  green: {
    headerBorder: "38761D",    // 濃い緑（ヘッダー下線）
    tableHeader: "93C47D",     // 明るい緑（変更履歴テーブルヘッダー）
    mermaid: "forest"          // Mermaidテーマ
  }
};

/**
 * テーマカラーを取得
 * @param {string} theme - テーマ名 (blue, orange, green)
 * @returns {object} - カラー設定
 */
function getThemeColors(theme) {
  return THEME_COLORS[theme] || THEME_COLORS.blue;
}

// ===== Mermaid設定 =====
const MERMAID_IMAGE_WIDTH = 600;

/**
 * Get image dimensions from PNG binary
 * @param {Buffer} pngData - PNG binary data
 * @returns {{width: number, height: number}|null} - Image dimensions, null on failure
 */
function getPngDimensions(pngData) {
  // PNGシグネチャ確認
  if (pngData.length < 24) return null;
  if (pngData[0] !== 0x89 || pngData[1] !== 0x50 || pngData[2] !== 0x4E || pngData[3] !== 0x47) {
    return null;
  }
  // IHDRチャンクから幅と高さを取得（ビッグエンディアン）
  const width = pngData.readUInt32BE(16);
  const height = pngData.readUInt32BE(20);
  return { width, height };
}

/**
 * Convert Mermaid diagram to PNG image using Kroki API
 * @param {string} diagramSource - Mermaid diagram source code
 * @param {string} mermaidTheme - Mermaid theme name (default, neutral, dark, forest, base)
 * @returns {Promise<Buffer|null>} - PNG image binary data, null on failure
 */
async function renderMermaidToPng(diagramSource, mermaidTheme = 'default') {
  const https = require('https');

  // テーマ設定をダイアグラムの先頭に追加（既存のinit設定がない場合）
  let postData = diagramSource;
  if (!diagramSource.includes('%%{init:')) {
    postData = `%%{init: {'theme': '${mermaidTheme}'}}%%\n${diagramSource}`;
  }

  return new Promise((resolve) => {
    const options = {
      hostname: 'kroki.io',
      port: 443,
      path: '/mermaid/png',
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain',
        'Content-Length': Buffer.byteLength(postData)
      },
      timeout: 30000
    };

    const req = https.request(options, (res) => {
      if (res.statusCode !== 200) {
        console.warn(`Warning: Kroki API error (HTTP ${res.statusCode})`);
        resolve(null);
        return;
      }
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
    });

    req.on('error', (e) => {
      console.warn(`Warning: Kroki API connection error: ${e.message}`);
      resolve(null);
    });

    req.on('timeout', () => {
      req.destroy();
      console.warn('Warning: Kroki API timeout');
      resolve(null);
    });

    req.write(postData);
    req.end();
  });
}

/**
 * Pre-render all Mermaid diagrams
 * @param {Array} elements - Array of parsed elements
 * @param {string} mermaidTheme - Mermaid theme name
 * @returns {Promise<Map<string, Buffer>>} - Map of diagram source to PNG data
 */
async function prerenderMermaidDiagrams(elements, mermaidTheme = 'default') {
  const mermaidElements = elements.filter(
    el => el.type === 'code' && el.language === 'mermaid'
  );

  const renderedMap = new Map();

  for (const el of mermaidElements) {
    if (!renderedMap.has(el.content)) {
      console.log('Rendering Mermaid diagram...');
      const imageData = await renderMermaidToPng(el.content, mermaidTheme);
      renderedMap.set(el.content, imageData);
    }
  }

  return renderedMap;
}

// ===== Config file loading =====
/**
 * Load YAML config file
 * @param {string} configPath - Config file path
 * @returns {object} - Config object (empty object if file not found)
 */
function loadConfigFile(configPath) {
  if (!configPath || !fs.existsSync(configPath)) {
    return {};
  }
  try {
    const content = fs.readFileSync(configPath, 'utf-8');
    const config = YAML.parse(content);
    return config || {};
  } catch (e) {
    console.warn(`Warning: Failed to load config file: ${e.message}`);
    return {};
  }
}

/**
 * Save config to YAML file
 * @param {string} savePath - Destination file path
 * @param {object} options - Config object to save
 * @param {object} defaults - Default values object
 */
function saveConfigFile(savePath, options, defaults) {
  // Force .yaml extension
  if (!savePath.endsWith('.yaml') && !savePath.endsWith('.yml')) {
    savePath = savePath + '.yaml';
  }

  // Config items to save (exclude input, output, config, save-config, and null values)
  const configToSave = {};
  for (const key of Object.keys(defaults)) {
    if (options[key] !== null) {
      configToSave[key] = options[key];
    }
  }

  try {
    const yamlContent = YAML.stringify(configToSave, { lineWidth: 0 });
    fs.writeFileSync(savePath, yamlContent);
    console.log(`Config saved: ${savePath}`);
  } catch (e) {
    console.error(`Error: Failed to save config file: ${e.message}`);
    process.exit(1);
  }
}

// ===== Command line argument parsing =====
function parseArgs() {
  const args = process.argv.slice(2);

  // Default values
  const defaults = {
    title: "Product Name",
    subtitle: "Manual",
    doctype: "Operation Manual",
    version: "1.0.0",
    date: new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }),
    dept: "Technical Development",
    docnum: "DOC-001",
    logo: null,
    company: "Sample Corporation",
    theme: "blue",
    "hr-pagebreak": true
  };

  // Parse command line arguments
  const cliOptions = {
    input: null,
    output: null,
    config: null,
    "save-config": null
  };
  const cliValues = {};

  for (let i = 0; i < args.length; i++) {
    if (args[i].startsWith('--')) {
      let key, value;
      // Support --key=value format
      if (args[i].includes('=')) {
        const eqIndex = args[i].indexOf('=');
        key = args[i].slice(2, eqIndex);
        value = args[i].slice(eqIndex + 1);
      } else {
        key = args[i].slice(2);
        value = (i + 1 < args.length) ? args[++i] : undefined;
      }

      if (key === 'input' || key === 'output' || key === 'config' || key === 'save-config') {
        cliOptions[key] = value;
      } else if (defaults.hasOwnProperty(key) && value !== undefined) {
        // Convert boolean values
        if (value === 'true') {
          cliValues[key] = true;
        } else if (value === 'false') {
          cliValues[key] = false;
        } else {
          cliValues[key] = value;
        }
      }
    } else if (!cliOptions.input) {
      cliOptions.input = args[i];
    } else if (!cliOptions.output) {
      cliOptions.output = args[i];
    }
  }

  // Input file not required when only using --save-config
  if (!cliOptions.input && !cliOptions["save-config"]) {
    console.error('Usage: node md2docx.js input.md output.docx [options]');
    console.error('Options: --title, --subtitle, --doctype, --version, --date, --dept, --docnum, --logo, --company, --theme, --config');
    console.error('Theme: blue (default), orange, green');
    console.error('');
    console.error('Config file: Automatically loads input.yaml from the same path as input.md');
    console.error('             Can also be explicitly specified with --config option');
    console.error('');
    console.error('Save config: Use --save-config config.yaml to save current settings to YAML file');
    process.exit(1);
  }

  // Determine config file path
  let configPath = cliOptions.config;
  if (!configPath && cliOptions.input) {
    // Look for .yaml file in the same path as md file
    const inputPath = path.resolve(cliOptions.input);
    const yamlPath = inputPath.replace(/\.md$/, '.yaml');
    if (fs.existsSync(yamlPath)) {
      configPath = yamlPath;
      console.log(`Loading config file: ${yamlPath}`);
    }
  }

  // Load config file
  const fileConfig = loadConfigFile(configPath);

  // Priority: Command line args > Config file > Default values
  const options = {
    input: cliOptions.input,
    output: cliOptions.output,
    ...defaults
  };

  // Apply config file values
  for (const key of Object.keys(defaults)) {
    if (fileConfig[key] !== undefined) {
      options[key] = fileConfig[key];
    }
  }

  // Apply command line argument values (highest priority)
  for (const key of Object.keys(cliValues)) {
    options[key] = cliValues[key];
  }

  // Validate theme
  if (!THEME_COLORS[options.theme]) {
    console.warn(`Warning: Unknown theme "${options.theme}". Using default "blue".`);
    options.theme = "blue";
  }

  if (!options.output && options.input) {
    options.output = options.input.replace(/\.md$/, '.docx');
  }

  // If --save-config is specified, save settings and exit
  if (cliOptions["save-config"]) {
    saveConfigFile(cliOptions["save-config"], options, defaults);
    process.exit(0);
  }

  return options;
}

// ===== Markdown Parser =====
class MarkdownParser {
  constructor(markdown) {
    this.lines = markdown.split('\n');
    this.pos = 0;
    this.elements = [];
    this.images = [];
  }

  parse() {
    while (this.pos < this.lines.length) {
      const line = this.lines[this.pos];
      
      if (line.trim() === '') {
        this.pos++;
        continue;
      }

      // コードブロック
      if (line.startsWith('```')) {
        this.parseCodeBlock();
        continue;
      }

      // 強制改ページ（タグ形式）
      if (line.match(/^<pagebreak\s*\/?>$/i)) {
        this.elements.push({ type: 'pagebreak' });
        this.pos++;
        continue;
      }

      // 強制改ページ（HTMLコメント形式）
      if (line.match(/^<!--\s*md2mdocx:pagebreak\s*-->$/i)) {
        this.elements.push({ type: 'pagebreak' });
        this.pos++;
        continue;
      }

      // 強制改行（HTMLコメント形式）
      if (line.match(/^<!--\s*md2mdocx:br\s*-->$/i)) {
        this.elements.push({ type: 'br' });
        this.pos++;
        continue;
      }

      // 見出し
      const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
      if (headingMatch) {
        this.elements.push({
          type: 'heading',
          level: headingMatch[1].length,
          text: headingMatch[2]
        });
        this.pos++;
        continue;
      }

      // 画像（Markdown形式）
      const imageMatch = line.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
      if (imageMatch) {
        this.elements.push({
          type: 'image',
          alt: imageMatch[1],
          src: imageMatch[2]
        });
        this.pos++;
        continue;
      }

      // 画像（HTMLのimgタグ）
      const imgTagMatch = line.match(/<img\s+[^>]*src=["']([^"']+)["'][^>]*>/i);
      if (imgTagMatch) {
        const altMatch = line.match(/alt=["']([^"']*)["']/i);
        const widthMatch = line.match(/width=["']?(\d+)["']?/i);
        const heightMatch = line.match(/height=["']?(\d+)["']?/i);
        this.elements.push({
          type: 'image',
          alt: altMatch ? altMatch[1] : '',
          src: imgTagMatch[1],
          width: widthMatch ? parseInt(widthMatch[1]) : null,
          height: heightMatch ? parseInt(heightMatch[1]) : null
        });
        this.pos++;
        continue;
      }

      // テーブル
      if (line.includes('|') && this.pos + 1 < this.lines.length && 
          this.lines[this.pos + 1].match(/^\|?[\s\-:|]+\|?$/)) {
        this.parseTable();
        continue;
      }

      // 箇条書き（番号なし）
      if (line.match(/^[\s]*[-*+]\s+/)) {
        this.parseList('bullet');
        continue;
      }

      // 箇条書き（番号付き）
      if (line.match(/^[\s]*\d+\.\s+/)) {
        this.parseList('number');
        continue;
      }

      // 引用
      if (line.startsWith('>')) {
        this.parseBlockquote();
        continue;
      }

      // 水平線
      if (line.match(/^[-*_]{3,}$/)) {
        this.elements.push({ type: 'hr' });
        this.pos++;
        continue;
      }

      // 通常の段落
      this.parseParagraph();
    }

    return this.elements;
  }

  parseCodeBlock() {
    const startLine = this.lines[this.pos];
    const lang = startLine.slice(3).trim();
    this.pos++;
    
    const codeLines = [];
    while (this.pos < this.lines.length && !this.lines[this.pos].startsWith('```')) {
      codeLines.push(this.lines[this.pos]);
      this.pos++;
    }
    this.pos++; // 閉じ```をスキップ

    this.elements.push({
      type: 'code',
      language: lang,
      content: codeLines.join('\n')
    });
  }

  parseTable() {
    const rows = [];
    
    // ヘッダー行
    const headerLine = this.lines[this.pos];
    rows.push(this.parseTableRow(headerLine));
    this.pos++;
    
    // 区切り行をスキップ
    this.pos++;
    
    // データ行
    while (this.pos < this.lines.length && this.lines[this.pos].includes('|')) {
      rows.push(this.parseTableRow(this.lines[this.pos]));
      this.pos++;
    }

    this.elements.push({ type: 'table', rows });
  }

  parseTableRow(line) {
    return line.split('|')
      .map(cell => cell.trim())
      .filter((cell, idx, arr) => idx > 0 && idx < arr.length - 1 || cell !== '');
  }

  parseList(listType) {
    const items = [];
    const topPattern = listType === 'bullet' ? /^[-*+]\s+(.+)$/ : /^\d+\.\s+(.+)$/;

    while (this.pos < this.lines.length) {
      const line = this.lines[this.pos];
      const topMatch = line.match(topPattern);

      if (topMatch) {
        // トップレベルのリスト項目
        items.push({ text: topMatch[1], level: 0 });
        this.pos++;
      } else if (line.match(/^(\s+)[-*+]\s+(.+)$/)) {
        // ネストされた箇条書き（インデントの深さでレベルを決定）
        const match = line.match(/^(\s+)[-*+]\s+(.+)$/);
        const indent = match[1].length;
        // 2-3スペースでlevel 1、4-5スペースでlevel 2、以降も同様
        const level = Math.min(Math.floor((indent + 1) / 2), 3);
        items.push({ text: match[2], level: level });
        this.pos++;
      } else if (line.trim() === '') {
        this.pos++;
        break;
      } else {
        break;
      }
    }

    this.elements.push({ type: 'list', listType, items });
  }

  parseBlockquote() {
    const lines = [];
    while (this.pos < this.lines.length && this.lines[this.pos].startsWith('>')) {
      lines.push(this.lines[this.pos].replace(/^>\s?/, ''));
      this.pos++;
    }
    this.elements.push({ type: 'blockquote', text: lines.join('\n') });
  }

  parseParagraph() {
    const lines = [];
    while (this.pos < this.lines.length) {
      const line = this.lines[this.pos];
      if (line.trim() === '' || line.startsWith('#') || line.startsWith('```') ||
          line.match(/^[-*+]\s+/) || line.match(/^\d+\.\s+/) || line.startsWith('>')) {
        break;
      }
      lines.push(line);
      this.pos++;
    }
    if (lines.length > 0) {
      this.elements.push({ type: 'paragraph', text: lines.join(' ') });
    }
  }
}

// ===== Inline markup processing =====
function parseInlineMarkup(text, inputDir = null) {
  const runs = [];
  let remaining = text;

  // 正規表現パターン
  const patterns = [
    { regex: /\*\*\*(.+?)\*\*\*/, style: { bold: true, italics: true } },
    { regex: /\*\*(.+?)\*\*/, style: { bold: true } },
    { regex: /\*(.+?)\*/, style: { italics: true } },
    { regex: /__(.+?)__/, style: { bold: true } },
    { regex: /_(.+?)_/, style: { italics: true } },
    { regex: /~~(.+?)~~/, style: { strike: true } },
    { regex: /`(.+?)`/, style: { shading: { fill: "E8E8E8", type: ShadingType.CLEAR } } }
  ];

  // imgタグのパターン
  const imgPattern = /<img\s+[^>]*src=["']([^"']+)["'][^>]*>/i;
  // brタグのパターン
  const brPattern = /<br\s*\/?>/i;

  while (remaining.length > 0) {
    let earliest = { index: remaining.length, length: 0, text: remaining, style: {}, type: 'text' };

    // brタグをチェック
    const brMatch = remaining.match(brPattern);
    if (brMatch && brMatch.index < earliest.index) {
      earliest = {
        index: brMatch.index,
        length: brMatch[0].length,
        before: remaining.slice(0, brMatch.index),
        type: 'break'
      };
    }

    // imgタグをチェック
    const imgMatch = remaining.match(imgPattern);
    if (imgMatch && imgMatch.index < earliest.index) {
      const widthMatch = imgMatch[0].match(/width=["']?(\d+)["']?/i);
      const heightMatch = imgMatch[0].match(/height=["']?(\d+)["']?/i);
      earliest = {
        index: imgMatch.index,
        length: imgMatch[0].length,
        src: imgMatch[1],
        width: widthMatch ? parseInt(widthMatch[1]) : null,
        height: heightMatch ? parseInt(heightMatch[1]) : null,
        before: remaining.slice(0, imgMatch.index),
        type: 'image'
      };
    }

    // テキストパターンをチェック
    for (const p of patterns) {
      const match = remaining.match(p.regex);
      if (match && match.index < earliest.index) {
        earliest = {
          index: match.index,
          length: match[0].length,
          text: match[1],
          style: p.style,
          before: remaining.slice(0, match.index),
          type: 'text'
        };
      }
    }

    if (earliest.before) {
      runs.push(new TextRun({ text: earliest.before, font: "Meiryo", size: 22 }));
    }

    if (earliest.index < remaining.length) {
      if (earliest.type === 'break') {
        // 強制改行
        runs.push(new TextRun({ break: 1 }));
      } else if (earliest.type === 'image' && inputDir) {
        // 画像を埋め込み
        try {
          const imgPath = path.isAbsolute(earliest.src) ? earliest.src : path.join(inputDir, earliest.src);
          if (fs.existsSync(imgPath)) {
            const imgData = fs.readFileSync(imgPath);
            const ext = path.extname(imgPath).slice(1).toLowerCase();
            const typeMap = { jpg: 'jpg', jpeg: 'jpg', png: 'png', gif: 'gif', bmp: 'bmp' };
            const imgWidth = earliest.width || 24;
            const imgHeight = earliest.height || earliest.width || 24;
            runs.push(new ImageRun({
              type: typeMap[ext] || 'png',
              data: imgData,
              transformation: { width: imgWidth, height: imgHeight }
            }));
          } else {
            runs.push(new TextRun({ text: `[Image: ${earliest.src}]`, font: "Meiryo", size: 22, color: "FF0000" }));
          }
        } catch (e) {
          runs.push(new TextRun({ text: `[Image error]`, font: "Meiryo", size: 22, color: "FF0000" }));
        }
      } else if (earliest.type === 'image') {
        // Display as text when inputDir is not available
        runs.push(new TextRun({ text: `[Image]`, font: "Meiryo", size: 22 }));
      } else {
        runs.push(new TextRun({ text: earliest.text, font: "Meiryo", size: 22, ...earliest.style }));
      }
      remaining = remaining.slice(earliest.index + earliest.length);
    } else {
      if (remaining.length > 0) {
        runs.push(new TextRun({ text: remaining, font: "Meiryo", size: 22 }));
      }
      break;
    }
  }

  return runs.length > 0 ? runs : [new TextRun({ text: text, font: "Meiryo", size: 22 })];
}

// ===== Header generation =====
function createHeader(options) {
  const colors = getThemeColors(options.theme);
  const headerBorder = { style: BorderStyle.SINGLE, size: 24, color: colors.headerBorder };

  // 6:4の比率
  const leftWidth = Math.floor(CONTENT_WIDTH * 0.6);
  const rightWidth = CONTENT_WIDTH - leftWidth;

  return new Header({
    children: [
      new Table({
        columnWidths: [leftWidth, rightWidth],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: { top: {style: BorderStyle.NIL}, bottom: headerBorder, left: {style: BorderStyle.NIL}, right: {style: BorderStyle.NIL} },
                width: { size: leftWidth, type: WidthType.DXA },
                children: [new Paragraph({
                  children: [new TextRun({ text: `${options.title} ${options.doctype} Version ${options.version}`, size: 18, font: "Meiryo", color: "000000" })]
                })]
              }),
              new TableCell({
                borders: { top: {style: BorderStyle.NIL}, bottom: headerBorder, left: {style: BorderStyle.NIL}, right: {style: BorderStyle.NIL} },
                width: { size: rightWidth, type: WidthType.DXA },
                children: [new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: `文書管理番号: ${options.docnum}`, size: 18, font: "Meiryo", color: "000000" })]
                })]
              })
            ]
          })
        ]
      })
    ]
  });
}

// ===== Footer generation =====
function createFooter() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "- ", font: "Meiryo" }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Meiryo" }),
        new TextRun({ text: " -", font: "Meiryo" })
      ]
    })]
  });
}

// ===== Cover section =====
function createCoverSection(options, inputDir) {
  const children = [
    new Paragraph({ spacing: { before: 2400 }, children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `${options.title} ${options.subtitle}`, bold: true, color: "000000", font: "Meiryo", size: 48 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({ text: options.doctype, color: "000000", font: "Meiryo", size: 36 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 480, after: 480 },
      children: [new TextRun({ text: `Version ${options.version}`, font: "Meiryo", size: 28 })]
    }),
    new Paragraph({ spacing: { before: 1200 }, children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: options.date, color: "000000", font: "Meiryo", size: 24 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: options.company, color: "000000", font: "Meiryo", size: 24 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: options.dept, color: "000000", font: "Meiryo", size: 24 })]
    })
  ];

  // Add logo image
  if (options.logo) {
    try {
      const logoPath = path.isAbsolute(options.logo) ? options.logo : path.join(inputDir, options.logo);
      if (fs.existsSync(logoPath)) {
        const logoData = fs.readFileSync(logoPath);
        const ext = path.extname(logoPath).slice(1).toLowerCase();
        const typeMap = { jpg: 'jpg', jpeg: 'jpg', png: 'png', gif: 'gif', bmp: 'bmp' };
        children.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 480 },
          children: [new ImageRun({
            type: typeMap[ext] || 'png',
            data: logoData,
            transformation: { width: 150, height: 150 }
          })]
        }));
      } else {
        console.warn(`Warning: Logo image not found: ${logoPath}`);
      }
    } catch (e) {
      console.warn(`Warning: Failed to load logo image: ${e.message}`);
    }
  }

  return {
    properties: {
      page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, size: { width: 11906, height: 16838 } }
    },
    headers: { default: createHeader(options) },
    footers: { default: createFooter() },
    children: children
  };
}

// ===== Changelog extraction =====
function extractChangelog(markdown) {
  const changelogMatch = markdown.match(/<!--\s*CHANGELOG\s*-->([\s\S]*?)<!--\s*\/CHANGELOG\s*-->/);
  if (!changelogMatch) {
    return null;
  }

  const changelogContent = changelogMatch[1].trim();
  const lines = changelogContent.split('\n').filter(line => line.trim());

  // Skip header and separator rows, get data rows
  const dataRows = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    // Skip separator row (|---|---|---|)
    if (line.match(/^\|[\s\-:|]+\|$/)) continue;
    // Skip header row (first data row)
    if (i === 0 && (line.includes('Version') || line.includes('バージョン'))) continue;

    // Parse table row
    if (line.includes('|')) {
      const cells = line.split('|').map(c => c.trim()).filter((c, idx, arr) => idx > 0 && idx < arr.length - 1 || c !== '');
      if (cells.length >= 3) {
        dataRows.push({
          version: cells[0],
          date: cells[1],
          description: cells[2]
        });
      }
    }
  }

  return dataRows.length > 0 ? dataRows : null;
}

// ===== Change history section =====
function createHistorySection(options, changelog) {
  const colors = getThemeColors(options.theme);

  // Prepare change history data
  const historyData = changelog || [
    { version: options.version, date: options.date, description: 'Initial release' }
  ];

  // Generate data rows
  const dataRows = historyData.map(item => new TableRow({
    children: [
      new TableCell({
        borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
        width: { size: 1800, type: WidthType.DXA },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.version, font: "Meiryo", size: 20 })] })]
      }),
      new TableCell({
        borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
        width: { size: 2000, type: WidthType.DXA },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.date, font: "Meiryo", size: 20 })] })]
      }),
      new TableCell({
        borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
        width: { size: 5226, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: item.description, font: "Meiryo", size: 20 })] })]
      })
    ]
  }));

  return {
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: createHeader(options) },
    footers: { default: createFooter() },
    children: [
      new Paragraph({ children: [new TextRun({ text: "[Change History]", bold: true, font: "Meiryo", size: 22 })] }),
      new Table({
        columnWidths: [1800, 2000, 5226],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({
                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                width: { size: 1800, type: WidthType.DXA },
                shading: { fill: colors.tableHeader, type: ShadingType.CLEAR },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Version", bold: true, font: "Meiryo", size: 20 })] })]
              }),
              new TableCell({
                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                width: { size: 2000, type: WidthType.DXA },
                shading: { fill: colors.tableHeader, type: ShadingType.CLEAR },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Date", bold: true, font: "Meiryo", size: 20 })] })]
              }),
              new TableCell({
                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                width: { size: 5226, type: WidthType.DXA },
                shading: { fill: colors.tableHeader, type: ShadingType.CLEAR },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Description", bold: true, font: "Meiryo", size: 20 })] })]
              })
            ]
          }),
          ...dataRows
        ]
      })
    ]
  };
}

// ===== Table of contents section =====
function createTOCSection(options) {
  return {
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: createHeader(options) },
    footers: { default: createFooter() },
    children: [
      new Paragraph({ children: [new TextRun({ text: "Table of Contents", bold: true, font: "Meiryo", size: 28 })] }),
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" })
    ]
  };
}

// ===== Convert body elements to Word elements =====
function convertElements(elements, options, inputDir, mermaidRenderedMap = new Map()) {
  const children = [];
  let numberListRef = 0;
  let currentSectionIndent = 0; // Current section indent

  for (const el of elements) {
    switch (el.type) {
      case 'heading': {
        const headingLevel = el.level <= 3 ? el.level : 3;
        const sizeMap = { 1: 28, 2: 24, 3: 22 };
        const spacingMap = { 1: { before: 360, after: 240 }, 2: { before: 240, after: 180 }, 3: { before: 180, after: 120 } };
        const headingStyleMap = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
        // No indent for #, 360 for ## and below
        currentSectionIndent = el.level === 1 ? 0 : 360;
        children.push(new Paragraph({
          heading: headingStyleMap[headingLevel],
          spacing: spacingMap[headingLevel],
          indent: { left: currentSectionIndent },
          children: [new TextRun({ text: el.text, bold: true, font: "Meiryo", size: sizeMap[headingLevel] })]
        }));
        break;
      }

      case 'paragraph':
        children.push(new Paragraph({
          indent: { left: currentSectionIndent },
          children: parseInlineMarkup(el.text, inputDir)
        }));
        break;

      case 'list':
        // Count numbered lists
        if (el.listType === 'number') numberListRef++;
        const listRef = `number-${numberListRef}-indent${currentSectionIndent}`;

        for (const item of el.items) {
          const itemText = typeof item === 'string' ? item : item.text;
          const itemLevel = typeof item === 'string' ? 0 : item.level;
          const baseIndent = currentSectionIndent + 720;
          const itemIndent = baseIndent + itemLevel * 360;

          if (el.listType === 'number' && itemLevel === 0) {
            // Use numbering for numbered lists
            children.push(new Paragraph({
              numbering: { reference: listRef, level: 0 },
              children: parseInlineMarkup(itemText, inputDir)
            }));
          } else {
            // Add bullet symbol for bullet lists
            const bullet = el.listType === 'bullet' ? (itemLevel === 0 ? '・' : '-') : '-';
            children.push(new Paragraph({
              indent: { left: itemIndent, hanging: 360 },
              children: [
                new TextRun({ text: bullet + '\t', font: "Meiryo", size: 22 }),
                ...parseInlineMarkup(itemText, inputDir)
              ]
            }));
          }
        }
        break;

      case 'code':
        // Mermaidコードブロックの処理
        if (el.language === 'mermaid') {
          const mermaidKey = el.content;
          const imageData = mermaidRenderedMap.get(mermaidKey);

          if (imageData) {
            // Get PNG image dimensions
            const dimensions = getPngDimensions(imageData);
            let imgWidth = dimensions ? dimensions.width : MERMAID_IMAGE_WIDTH;
            let imgHeight = dimensions ? dimensions.height : Math.round(MERMAID_IMAGE_WIDTH * 0.6);

            // Scale down to fit page width/height (TWIP to pixels: 1 inch = 1440 TWIP, 96 dpi)
            const maxWidth = (CONTENT_WIDTH - currentSectionIndent) / 1440 * 96;
            const maxHeight = 600; // Maximum height (pixels)
            if (imgWidth > maxWidth || imgHeight > maxHeight) {
              const scaleW = maxWidth / imgWidth;
              const scaleH = maxHeight / imgHeight;
              const scale = Math.min(scaleW, scaleH);
              imgWidth = Math.round(imgWidth * scale);
              imgHeight = Math.round(imgHeight * scale);
            }

            // Embed as PNG image
            children.push(new Paragraph({
              indent: { left: currentSectionIndent },
              alignment: AlignmentType.CENTER,
              children: [new ImageRun({
                type: 'png',
                data: imageData,
                transformation: { width: imgWidth, height: imgHeight },
                altText: {
                  title: 'Mermaid Diagram',
                  description: 'Mermaid diagram',
                  name: 'mermaid-diagram'
                }
              })]
            }));
          } else {
            // Fallback: Display warning message
            children.push(new Paragraph({
              indent: { left: currentSectionIndent },
              shading: { fill: "FFF3CD", type: ShadingType.CLEAR },
              border: { left: { style: BorderStyle.SINGLE, size: 12, color: "FFC107" } },
              children: [
                new TextRun({
                  text: '[Mermaid diagram: Rendering failed - API connection error or offline]',
                  font: "Meiryo",
                  size: 20,
                  color: "856404"
                })
              ]
            }));
          }
          break;
        }

        // Display regular code blocks with background color
        const codeLines = el.content.split('\n');
        for (const codeLine of codeLines) {
          children.push(new Paragraph({
            shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
            indent: { left: currentSectionIndent + 360 },
            children: [new TextRun({ text: codeLine || ' ', font: "Meiryo", size: 20 })]
          }));
        }
        break;

      case 'table':
        const colCount = el.rows[0]?.length || 1;
        const tableWidth = CONTENT_WIDTH - currentSectionIndent;
        const colWidth = Math.floor(tableWidth / colCount);

        const tableRows = el.rows.map((row, rowIdx) => {
          return new TableRow({
            tableHeader: rowIdx === 0,
            children: row.map(cell => new TableCell({
              borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
              width: { size: colWidth, type: WidthType.DXA },
              shading: rowIdx === 0 ? { fill: "D9D9D9", type: ShadingType.CLEAR } : undefined,
              children: [new Paragraph({
                children: parseInlineMarkup(cell, inputDir)
              })]
            }))
          });
        });

        children.push(new Table({
          columnWidths: Array(colCount).fill(colWidth),
          rows: tableRows,
          width: { size: tableWidth, type: WidthType.DXA },
          indent: { size: currentSectionIndent, type: WidthType.DXA }
        }));
        break;

      case 'blockquote':
        children.push(new Paragraph({
          indent: { left: currentSectionIndent + 720 },
          border: { left: { style: BorderStyle.SINGLE, size: 24, color: "CCCCCC" } },
          children: [new TextRun({ text: el.text, italics: true, font: "Meiryo", size: 22 })]
        }));
        break;

      case 'image':
        try {
          const imgPath = path.isAbsolute(el.src) ? el.src : path.join(inputDir, el.src);
          if (fs.existsSync(imgPath)) {
            const imgData = fs.readFileSync(imgPath);
            const ext = path.extname(imgPath).slice(1).toLowerCase();
            const typeMap = { jpg: 'jpg', jpeg: 'jpg', png: 'png', gif: 'gif', bmp: 'bmp' };
            // Use width/height if specified, otherwise use defaults
            const imgWidth = el.width || 400;
            const imgHeight = el.height || (el.width ? Math.round(el.width * 0.75) : 300);
            children.push(new Paragraph({
              indent: { left: currentSectionIndent },
              alignment: AlignmentType.CENTER,
              children: [new ImageRun({
                type: typeMap[ext] || 'png',
                data: imgData,
                transformation: { width: imgWidth, height: imgHeight },
                altText: { title: el.alt || 'Image', description: el.alt || 'Image', name: el.alt || 'Image' }
              })]
            }));
          } else {
            children.push(new Paragraph({
              indent: { left: currentSectionIndent },
              children: [new TextRun({ text: `[Image: ${el.src}]`, font: "Meiryo", size: 22, color: "FF0000" })]
            }));
          }
        } catch (e) {
          children.push(new Paragraph({
            indent: { left: currentSectionIndent },
            children: [new TextRun({ text: `[Image load error: ${el.src}]`, font: "Meiryo", size: 22, color: "FF0000" })]
          }));
        }
        break;

      case 'hr':
        if (options["hr-pagebreak"]) {
          // Treat horizontal rule as page break
          children.push(new Paragraph({
            children: [new PageBreak()]
          }));
        } else {
          // Regular horizontal rule
          children.push(new Paragraph({
            indent: { left: currentSectionIndent },
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC" } },
            children: []
          }));
        }
        break;

      case 'pagebreak':
        children.push(new Paragraph({
          children: [new PageBreak()]
        }));
        break;

      case 'br':
        children.push(new Paragraph({
          indent: { left: currentSectionIndent },
          children: []
        }));
        break;
    }
  }

  return children;
}

// ===== Main processing =====
async function main() {
  const options = parseArgs();

  // Read Markdown file
  const markdownRaw = fs.readFileSync(options.input, 'utf-8');
  const inputDir = path.dirname(path.resolve(options.input));

  // Extract changelog
  const changelog = extractChangelog(markdownRaw);

  // Exclude CHANGELOG block from body
  let markdown = markdownRaw.replace(/<!--\s*CHANGELOG\s*-->[\s\S]*?<!--\s*\/CHANGELOG\s*-->\s*/, '');

  // Range specification using md2mdocx:start/end
  const startMatch = markdown.match(/<!--\s*md2mdocx:start\s*-->/i);
  const endMatch = markdown.match(/<!--\s*md2mdocx:end\s*-->/i);
  if (startMatch) {
    markdown = markdown.slice(startMatch.index + startMatch[0].length);
  }
  if (endMatch) {
    const endIdx = markdown.search(/<!--\s*md2mdocx:end\s*-->/i);
    if (endIdx !== -1) {
      markdown = markdown.slice(0, endIdx);
    }
  }

  // Parse
  const parser = new MarkdownParser(markdown);
  const elements = parser.parse();

  // Pre-render Mermaid diagrams (with theme applied)
  const colors = getThemeColors(options.theme);
  const mermaidRenderedMap = await prerenderMermaidDiagrams(elements, colors.mermaid);

  // Dynamically generate numbered list settings (no indent for #, 360 for ## and below)
  const numberConfigs = [];
  let listCount = 0;
  let currentIndent = 0;

  for (const el of elements) {
    if (el.type === 'heading') {
      currentIndent = el.level === 1 ? 0 : 360;
    }
    if (el.type === 'list' && el.listType === 'number') {
      listCount++;
      numberConfigs.push({
        reference: `number-${listCount}-indent${currentIndent}`,
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: currentIndent + 720, hanging: 360 } } } }]
      });
    }
  }

  // Convert to Word elements
  const contentChildren = convertElements(elements, options, inputDir, mermaidRenderedMap);

  // Generate document
  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Meiryo", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, color: "000000", font: "Meiryo" },
          paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, color: "000000", font: "Meiryo" },
          paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 22, bold: true, color: "000000", font: "Meiryo" },
          paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2 } }
      ]
    },
    numbering: { config: numberConfigs },
    sections: [
      createCoverSection(options, inputDir),
      createHistorySection(options, changelog),
      createTOCSection(options),
      {
        properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
        headers: { default: createHeader(options) },
        footers: { default: createFooter() },
        children: contentChildren
      }
    ]
  });

  // Save
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(options.output, buffer);
  console.log(`Done: ${options.output}`);
}

main().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
