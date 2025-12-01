#!/usr/bin/env node
/**
 * Markdown to Word Converter
 * マニュアルテンプレート形式
 *
 * Usage: bun md2docx.js input.md output.docx [options]
 * Options:
 *   --title "製品名"
 *   --subtitle "マニュアル"
 *   --doctype "操作マニュアル"
 *   --version "1.0.0"
 *   --date "2024年1月1日"
 *   --dept "技術開発部"
 *   --docnum "DOC-001"
 *   --logo "logo.png"
 *   --company "会社名"
 *   --theme "blue|orange|green"  (差し色テーマ)
 */

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
        AlignmentType, PageNumber, BorderStyle, WidthType, HeadingLevel, PageBreak,
        TableOfContents, ShadingType, LevelFormat, ImageRun } = require('docx');
const fs = require('fs');
const path = require('path');

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
 * PNGバイナリから画像サイズを取得
 * @param {Buffer} pngData - PNGバイナリデータ
 * @returns {{width: number, height: number}|null} - 画像サイズ、取得失敗時はnull
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
 * Kroki APIを使用してMermaid図をPNG画像に変換
 * @param {string} diagramSource - Mermaidダイアグラムのソースコード
 * @param {string} mermaidTheme - Mermaidテーマ名 (default, neutral, dark, forest, base)
 * @returns {Promise<Buffer|null>} - PNG画像のバイナリデータ、失敗時はnull
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
        console.warn(`警告: Kroki APIエラー (HTTP ${res.statusCode})`);
        resolve(null);
        return;
      }
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
    });

    req.on('error', (e) => {
      console.warn(`警告: Kroki API接続エラー: ${e.message}`);
      resolve(null);
    });

    req.on('timeout', () => {
      req.destroy();
      console.warn('警告: Kroki APIタイムアウト');
      resolve(null);
    });

    req.write(postData);
    req.end();
  });
}

/**
 * すべてのMermaid図を事前レンダリング
 * @param {Array} elements - パース済み要素の配列
 * @param {string} mermaidTheme - Mermaidテーマ名
 * @returns {Promise<Map<string, Buffer>>} - ダイアグラムソースをキーとするPNGデータのMap
 */
async function prerenderMermaidDiagrams(elements, mermaidTheme = 'default') {
  const mermaidElements = elements.filter(
    el => el.type === 'code' && el.language === 'mermaid'
  );

  const renderedMap = new Map();

  for (const el of mermaidElements) {
    if (!renderedMap.has(el.content)) {
      console.log('Mermaid図をレンダリング中...');
      const imageData = await renderMermaidToPng(el.content, mermaidTheme);
      renderedMap.set(el.content, imageData);
    }
  }

  return renderedMap;
}

// ===== コマンドライン引数パース =====
function parseArgs() {
  const args = process.argv.slice(2);
  const options = {
    input: null,
    output: null,
    title: "製品名",
    subtitle: "マニュアル",
    doctype: "操作マニュアル",
    version: "1.0.0",
    date: new Date().toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' }),
    dept: "技術開発部",
    docnum: "DOC-001",
    logo: null,
    company: "サンプル株式会社",
    theme: "blue"
  };

  for (let i = 0; i < args.length; i++) {
    if (args[i].startsWith('--')) {
      const key = args[i].slice(2);
      if (options.hasOwnProperty(key) && i + 1 < args.length) {
        options[key] = args[++i];
      }
    } else if (!options.input) {
      options.input = args[i];
    } else if (!options.output) {
      options.output = args[i];
    }
  }

  if (!options.input) {
    console.error('Usage: node md2docx.js input.md output.docx [options]');
    console.error('Options: --title, --subtitle, --doctype, --version, --date, --dept, --docnum, --logo, --company, --theme');
    console.error('Theme: blue (default), orange, green');
    process.exit(1);
  }

  // テーマの検証
  if (!THEME_COLORS[options.theme]) {
    console.warn(`警告: 不明なテーマ "${options.theme}"。デフォルトの "blue" を使用します。`);
    options.theme = "blue";
  }

  if (!options.output) {
    options.output = options.input.replace(/\.md$/, '.docx');
  }

  return options;
}

// ===== Markdownパーサー =====
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

      // 強制改ページ
      if (line.match(/^<pagebreak\s*\/?>$/i)) {
        this.elements.push({ type: 'pagebreak' });
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

// ===== インラインマークアップ処理 =====
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
            runs.push(new TextRun({ text: `[画像: ${earliest.src}]`, font: "Meiryo", size: 22, color: "FF0000" }));
          }
        } catch (e) {
          runs.push(new TextRun({ text: `[画像エラー]`, font: "Meiryo", size: 22, color: "FF0000" }));
        }
      } else if (earliest.type === 'image') {
        // inputDirがない場合はテキストで表示
        runs.push(new TextRun({ text: `[画像]`, font: "Meiryo", size: 22 }));
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

// ===== ヘッダー生成 =====
function createHeader(options) {
  const colors = getThemeColors(options.theme);
  const headerBorder = { style: BorderStyle.SINGLE, size: 24, color: colors.headerBorder };

  return new Header({
    children: [
      new Table({
        columnWidths: [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: { top: {style: BorderStyle.NIL}, bottom: headerBorder, left: {style: BorderStyle.NIL}, right: {style: BorderStyle.NIL} },
                width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA },
                children: [new Paragraph({
                  children: [new TextRun({ text: `${options.title} ${options.doctype} Version ${options.version}`, size: 20, font: "Meiryo", color: "000000" })]
                })]
              }),
              new TableCell({
                borders: { top: {style: BorderStyle.NIL}, bottom: headerBorder, left: {style: BorderStyle.NIL}, right: {style: BorderStyle.NIL} },
                width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA },
                children: [new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: `文書管理番号:　${options.docnum}`, size: 20, font: "Meiryo", color: "000000" })]
                })]
              })
            ]
          })
        ]
      })
    ]
  });
}

// ===== フッター生成 =====
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

// ===== 表紙セクション =====
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

  // ロゴ画像を追加
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
        console.warn(`警告: ロゴ画像が見つかりません: ${logoPath}`);
      }
    } catch (e) {
      console.warn(`警告: ロゴ画像の読み込みに失敗しました: ${e.message}`);
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

// ===== 変更履歴の抽出 =====
function extractChangelog(markdown) {
  const changelogMatch = markdown.match(/<!--\s*CHANGELOG\s*-->([\s\S]*?)<!--\s*\/CHANGELOG\s*-->/);
  if (!changelogMatch) {
    return null;
  }

  const changelogContent = changelogMatch[1].trim();
  const lines = changelogContent.split('\n').filter(line => line.trim());

  // ヘッダー行と区切り行をスキップしてデータ行を取得
  const dataRows = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    // 区切り行（|---|---|---|）をスキップ
    if (line.match(/^\|[\s\-:|]+\|$/)) continue;
    // ヘッダー行（最初のデータ行）をスキップ
    if (i === 0 && line.includes('バージョン')) continue;

    // テーブル行をパース
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

// ===== 変更履歴セクション =====
function createHistorySection(options, changelog) {
  const colors = getThemeColors(options.theme);

  // 変更履歴データの準備
  const historyData = changelog || [
    { version: options.version, date: options.date, description: '初版作成' }
  ];

  // データ行を生成
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
      new Paragraph({ children: [new TextRun({ text: "[変更履歴]", bold: true, font: "Meiryo", size: 22 })] }),
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
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "バージョン", bold: true, font: "Meiryo", size: 20 })] })]
              }),
              new TableCell({
                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                width: { size: 2000, type: WidthType.DXA },
                shading: { fill: colors.tableHeader, type: ShadingType.CLEAR },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "変更日付", bold: true, font: "Meiryo", size: 20 })] })]
              }),
              new TableCell({
                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                width: { size: 5226, type: WidthType.DXA },
                shading: { fill: colors.tableHeader, type: ShadingType.CLEAR },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "変更事由", bold: true, font: "Meiryo", size: 20 })] })]
              })
            ]
          }),
          ...dataRows
        ]
      })
    ]
  };
}

// ===== 目次セクション =====
function createTOCSection(options) {
  return {
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: createHeader(options) },
    footers: { default: createFooter() },
    children: [
      new Paragraph({ children: [new TextRun({ text: "目次", bold: true, font: "Meiryo", size: 28 })] }),
      new TableOfContents("目次", { hyperlink: true, headingStyleRange: "1-3" })
    ]
  };
}

// ===== 本文要素をWord要素に変換 =====
function convertElements(elements, options, inputDir, mermaidRenderedMap = new Map()) {
  const children = [];
  let numberListRef = 0;
  let currentSectionIndent = 0; // 現在のセクションのインデント

  for (const el of elements) {
    switch (el.type) {
      case 'heading': {
        const headingLevel = el.level <= 3 ? el.level : 3;
        const sizeMap = { 1: 28, 2: 24, 3: 22 };
        const spacingMap = { 1: { before: 360, after: 240 }, 2: { before: 240, after: 180 }, 3: { before: 180, after: 120 } };
        const headingStyleMap = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
        // #はインデントなし、##以降は360
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
        // 番号付きリストのカウント
        if (el.listType === 'number') numberListRef++;
        const listRef = `number-${numberListRef}-indent${currentSectionIndent}`;

        for (const item of el.items) {
          const itemText = typeof item === 'string' ? item : item.text;
          const itemLevel = typeof item === 'string' ? 0 : item.level;
          const baseIndent = currentSectionIndent + 720;
          const itemIndent = baseIndent + itemLevel * 360;

          if (el.listType === 'number' && itemLevel === 0) {
            // 番号付きリストはnumberingを使用
            children.push(new Paragraph({
              numbering: { reference: listRef, level: 0 },
              children: parseInlineMarkup(itemText, inputDir)
            }));
          } else {
            // 箇条書きはテキストで記号を追加
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
            // PNG画像のサイズを取得
            const dimensions = getPngDimensions(imageData);
            let imgWidth = dimensions ? dimensions.width : MERMAID_IMAGE_WIDTH;
            let imgHeight = dimensions ? dimensions.height : Math.round(MERMAID_IMAGE_WIDTH * 0.6);

            // ページ幅・高さに収まるよう縮小（TWIP→ピクセル: 1インチ=1440TWIP, 96dpi）
            const maxWidth = (CONTENT_WIDTH - currentSectionIndent) / 1440 * 96;
            const maxHeight = 600; // 最大高さ（ピクセル）
            if (imgWidth > maxWidth || imgHeight > maxHeight) {
              const scaleW = maxWidth / imgWidth;
              const scaleH = maxHeight / imgHeight;
              const scale = Math.min(scaleW, scaleH);
              imgWidth = Math.round(imgWidth * scale);
              imgHeight = Math.round(imgHeight * scale);
            }

            // PNG画像として埋め込み
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
            // フォールバック: 警告メッセージを表示
            children.push(new Paragraph({
              indent: { left: currentSectionIndent },
              shading: { fill: "FFF3CD", type: ShadingType.CLEAR },
              border: { left: { style: BorderStyle.SINGLE, size: 12, color: "FFC107" } },
              children: [
                new TextRun({
                  text: '[Mermaid図: レンダリング失敗 - API接続エラーまたはオフライン]',
                  font: "Meiryo",
                  size: 20,
                  color: "856404"
                })
              ]
            }));
          }
          break;
        }

        // 通常のコードブロックは背景色付きで表示
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
            // width/heightが指定されていればそれを使用、なければデフォルト
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
              children: [new TextRun({ text: `[画像: ${el.src}]`, font: "Meiryo", size: 22, color: "FF0000" })]
            }));
          }
        } catch (e) {
          children.push(new Paragraph({
            indent: { left: currentSectionIndent },
            children: [new TextRun({ text: `[画像読み込みエラー: ${el.src}]`, font: "Meiryo", size: 22, color: "FF0000" })]
          }));
        }
        break;

      case 'hr':
        children.push(new Paragraph({
          indent: { left: currentSectionIndent },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC" } },
          children: []
        }));
        break;

      case 'pagebreak':
        children.push(new Paragraph({
          children: [new PageBreak()]
        }));
        break;
    }
  }

  return children;
}

// ===== メイン処理 =====
async function main() {
  const options = parseArgs();
  
  // Markdownファイル読み込み
  const markdownRaw = fs.readFileSync(options.input, 'utf-8');
  const inputDir = path.dirname(path.resolve(options.input));

  // 変更履歴を抽出
  const changelog = extractChangelog(markdownRaw);

  // CHANGELOGブロックを本文から除外
  const markdown = markdownRaw.replace(/<!--\s*CHANGELOG\s*-->[\s\S]*?<!--\s*\/CHANGELOG\s*-->\s*/, '');

  // パース
  const parser = new MarkdownParser(markdown);
  const elements = parser.parse();

  // Mermaid図の事前レンダリング（テーマ適用）
  const colors = getThemeColors(options.theme);
  const mermaidRenderedMap = await prerenderMermaidDiagrams(elements, colors.mermaid);

  // 番号付きリストの設定を動的に生成（#はインデントなし、##以降は360）
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

  // Word要素に変換
  const contentChildren = convertElements(elements, options, inputDir, mermaidRenderedMap);

  // ドキュメント生成
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

  // 保存
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(options.output, buffer);
  console.log(`✓ 変換完了: ${options.output}`);
}

main().catch(err => {
  console.error('エラー:', err.message);
  process.exit(1);
});
