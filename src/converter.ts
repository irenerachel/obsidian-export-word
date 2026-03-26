import { App, TFile, normalizePath, requestUrl } from "obsidian";
import MarkdownIt from "markdown-it";
import markdownItFootnote from "markdown-it-footnote";
import {
  BorderStyle, Document, ExternalHyperlink, HeadingLevel, ImageRun,
  Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType,
  TableOfContents, StyleLevel, AlignmentType, PageBreak,
  Header, Footer, PageNumber, TextWrappingType, TextWrappingSide,
  ISectionOptions,
} from "docx";
import { getImageDimensions, guessImageType, guessImageTypeFromUrl } from "./image-utils";
import type { ExportWordSettings } from "./settings";

const IMAGE_EXTS = new Set(["png", "jpg", "jpeg", "gif", "webp", "svg", "bmp", "tif", "tiff"]);

const CALLOUT_COLORS: Record<string, string> = {
  note: "448AFF", info: "448AFF", tip: "00BFA5", hint: "00BFA5", important: "00BFA5",
  warning: "FF9100", caution: "FF9100", attention: "FF9100",
  danger: "FF5252", error: "FF5252", bug: "FF5252", failure: "FF5252", fail: "FF5252", missing: "FF5252",
  success: "00C853", check: "00C853", done: "00C853",
  question: "64DD17", help: "64DD17", faq: "64DD17",
  example: "7C4DFF", abstract: "00B0FF", summary: "00B0FF", tldr: "00B0FF",
  quote: "9E9E9E", cite: "9E9E9E",
};

interface ConvertContext {
  app: App;
  sourceFile: TFile;
  settings: ExportWordSettings;
  maxW: number;
  matched: number;
  warnings: string[];
  headings: { level: number; text: string }[];
}

/* ── Public entry ── */

export async function convertToDocx(
  markdown: string,
  title: string,
  app: App,
  sourceFile: TFile,
  settings: ExportWordSettings,
): Promise<{ buffer: ArrayBuffer; matched: number; warnings: string[] }> {

  const maxW = settings.imageSizing === "original" ? Infinity : settings.maxImageWidth;
  const ctx: ConvertContext = { app, sourceFile, settings, maxW, matched: 0, warnings: [], headings: [] };

  const normalized = await normalizeObsidianMarkdown(markdown, ctx);

  const md = new MarkdownIt({ html: true, linkify: true, typographer: false, breaks: false });
  md.enable("strikethrough");
  md.use(markdownItFootnote);

  const html = md.render(normalized);
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div id="r">${html}</div>`, "text/html");
  const root = doc.getElementById("r")!;
  const children = await convertBlocks(Array.from(root.childNodes), ctx);

  const fontSize = settings.fontSize * 2; // docx uses half-points
  const font = settings.defaultFont;
  const cjkFont = settings.enableSmartFont ? settings.cjkFont : font;

  // Build sections
  const sections: ISectionOptions[] = [];
  const pageProps = {
    page: {
      size: { width: 12240, height: 15840 },
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
    },
  };

  // Header
  const headerChildren: Paragraph[] = [];
  if (settings.headerText) {
    const headerStr = settings.headerText.replace(/\{title\}/g, title);
    headerChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: headerStr, size: 18, color: "999999", font })],
    }));
  }

  // Footer with page numbers
  const footerChildren: Paragraph[] = [];
  if (settings.enablePageNumbers) {
    footerChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "999999", font }),
        new TextRun({ text: " / ", size: 18, color: "999999", font }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "999999", font }),
      ],
    }));
  }

  const sectionHeaderFooter: any = {};
  if (headerChildren.length) sectionHeaderFooter.headers = { default: new Header({ children: headerChildren }) };
  if (footerChildren.length) sectionHeaderFooter.footers = { default: new Footer({ children: footerChildren }) };

  // Cover page section
  if (settings.enableCoverPage) {
    const now = new Date();
    const dateStr = `${now.getFullYear()} 年 ${now.getMonth() + 1} 月 ${now.getDate()} 日`;
    const coverChildren: Paragraph[] = [
      new Paragraph({ spacing: { before: 4000 } }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: title, bold: true, size: fontSize + 28, font: cjkFont })],
      }),
    ];
    if (settings.coverAuthor) {
      coverChildren.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: settings.coverAuthor, size: fontSize + 4, color: "666666", font: cjkFont })],
      }));
    }
    coverChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: dateStr, size: fontSize, color: "999999", font })],
    }));

    sections.push({
      properties: { ...pageProps, ...sectionHeaderFooter },
      children: coverChildren,
    });
  }

  // Main content section
  const docChildren: (Paragraph | Table | TableOfContents)[] = [];

  // Optional TOC
  if (settings.enableToc && ctx.headings.length > 0) {
    docChildren.push(
      new TableOfContents("目录", {
        hyperlink: true,
        headingStyleRange: "1-6",
        stylesWithLevels: [
          new StyleLevel("Heading1", 1), new StyleLevel("Heading2", 2), new StyleLevel("Heading3", 3),
          new StyleLevel("Heading4", 4), new StyleLevel("Heading5", 5), new StyleLevel("Heading6", 6),
        ],
      })
    );
    docChildren.push(new Paragraph({ spacing: { after: 300 } }));
  }
  docChildren.push(...children);

  sections.push({
    properties: { ...pageProps, ...sectionHeaderFooter },
    children: docChildren,
  });

  const wordDoc = new Document({
    creator: "Export to Word",
    title,
    description: "Exported from Obsidian",
    styles: {
      default: {
        document: {
          run: { font, size: fontSize },
          paragraph: { spacing: { after: 200 } },
        },
        heading1: { run: { font: cjkFont, size: fontSize + 16, bold: true } },
        heading2: { run: { font: cjkFont, size: fontSize + 12, bold: true } },
        heading3: { run: { font: cjkFont, size: fontSize + 8, bold: true } },
        heading4: { run: { font: cjkFont, size: fontSize + 4, bold: true } },
      },
    },
    features: { updateFields: true },
    sections,
  });

  const blob = await Packer.toBlob(wordDoc);
  const buffer = await blob.arrayBuffer();

  // Word/character count
  const plainText = markdown.replace(/^---[\s\S]*?---\n*/, "").replace(/[#*_~`>\[\]!|]/g, "");
  const charCount = plainText.replace(/\s/g, "").length;
  const wordCount = plainText.trim().split(/\s+/).filter(Boolean).length;

  return { buffer, matched: ctx.matched, warnings: ctx.warnings, charCount, wordCount };
}

/* ── Markdown normalization ── */

async function normalizeObsidianMarkdown(text: string, ctx: ConvertContext): Promise<string> {
  let out = text;

  // Strip frontmatter
  out = out.replace(/^---[\s\S]*?---\n*/, "");

  // Convert ![[embed]] syntax
  const embedRegex = /!\[\[([^\]]+)\]\]/g;
  const embedMatches = [...out.matchAll(embedRegex)];
  for (const m of embedMatches.reverse()) {
    const replacement = await processEmbed(m[1].trim(), ctx);
    out = out.slice(0, m.index!) + replacement + out.slice(m.index! + m[0].length);
  }

  // Convert [[wiki links]] to plain text
  out = out.replace(/\[\[([^\]]+)\]\]/g, (_, raw: string) => {
    const parts = raw.trim().split("|");
    const target = (parts[0] || "").trim();
    const alias = (parts[1] || "").trim();
    if (alias) return alias;
    const base = target.split("/").pop() || target;
    return base.replace(/\.[^.]+$/, "");
  });

  // Highlights ==text== → <mark>text</mark>
  out = out.replace(/==(.*?)==/g, "<mark>$1</mark>");

  // Tags #tag → inline code
  out = out.replace(/(^|\s)#([a-zA-Z\u4e00-\u9fff][\w\u4e00-\u9fff/-]*)/g, "$1`#$2`");

  // Callouts: pre-process before markdown-it
  if (ctx.settings.enableCallouts) {
    out = preprocessCallouts(out);
  }

  return out;
}

async function processEmbed(rawContent: string, ctx: ConvertContext): Promise<string> {
  const parts = rawContent.split("|");
  const target = (parts[0] || "").trim();
  const option = (parts[1] || "").trim();

  const file = ctx.app.metadataCache.getFirstLinkpathDest(target, ctx.sourceFile.path);
  if (!file) {
    ctx.warnings.push(`嵌入文件未找到：${target}`);
    return `**[缺失：${escMd(target)}]**`;
  }

  const ext = file.extension?.toLowerCase() || "";

  // Image embed
  if (IMAGE_EXTS.has(ext)) {
    const uri = encodeURI(file.path).replace(/\(/g, "%28").replace(/\)/g, "%29");
    if (/^\d+$/.test(option)) return `![${escMd(file.basename)}](${uri} "width=${option}")`;
    return `![${escMd(option || file.basename)}](${uri})`;
  }

  // Note embed: inline the content (with recursion guard)
  if (ext === "md") {
    try {
      const content = await ctx.app.vault.cachedRead(file);
      const stripped = content.replace(/^---[\s\S]*?---\n*/, "");
      return `\n\n${stripped}\n\n`;
    } catch {
      return `**[嵌入文件：${escMd(file.basename)}]**`;
    }
  }

  return `**[嵌入文件：${escMd(file.basename)}]**`;
}

/* ── Callout preprocessing ── */

function preprocessCallouts(text: string): string {
  const lines = text.split("\n");
  const result: string[] = [];
  let i = 0;

  while (i < lines.length) {
    const calloutMatch = lines[i].match(/^>\s*\[!([\w-]+)\]([+-]?)(?:\s+(.*))?$/);
    if (calloutMatch) {
      const type = calloutMatch[1].toLowerCase();
      const title = calloutMatch[3] || type.charAt(0).toUpperCase() + type.slice(1);
      const bodyLines: string[] = [];

      i++;
      while (i < lines.length && lines[i].match(/^>\s?/)) {
        bodyLines.push(lines[i].replace(/^>\s?/, ""));
        i++;
      }

      const color = CALLOUT_COLORS[type] || "448AFF";
      const body = bodyLines.join("\n").trim();
      result.push(`<div class="callout" data-type="${type}" data-color="${color}" data-title="${escHtml(title)}">${escHtml(body)}</div>`);
      result.push("");
    } else {
      result.push(lines[i]);
      i++;
    }
  }

  return result.join("\n");
}

function escHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/* ── Block conversion ── */

async function convertBlocks(nodes: ChildNode[], ctx: ConvertContext): Promise<(Paragraph | Table)[]> {
  const blocks: (Paragraph | Table)[] = [];

  for (const node of nodes) {
    if (node.nodeType === Node.TEXT_NODE && !node.textContent?.trim()) continue;
    if (node.nodeType !== Node.ELEMENT_NODE) continue;
    const el = node as HTMLElement;
    const tag = el.tagName;

    if (/^H[1-6]$/.test(tag)) {
      const levelNum = parseInt(tag.slice(1));
      const level = `HEADING_${levelNum}` as keyof typeof HeadingLevel;
      const text = el.textContent || "";
      ctx.headings.push({ level: levelNum, text });
      const headingInlines = await inlines(el.childNodes, ctx);
      // H1 page break: insert page break before H1 (except the first one)
      if (levelNum === 1 && ctx.settings.enableH1PageBreak && blocks.length > 0) {
        headingInlines.unshift(new TextRun({ children: [new PageBreak()] }));
      }
      blocks.push(new Paragraph({
        heading: HeadingLevel[level],
        spacing: { after: 220 },
        children: headingInlines,
      }));
    } else if (tag === "P") {
      blocks.push(new Paragraph({
        spacing: { after: 180 },
        children: await inlines(el.childNodes, ctx),
      }));
    } else if (tag === "UL" || tag === "OL") {
      blocks.push(...(await convertList(el, ctx, 0)));
    } else if (tag === "PRE") {
      const code = el.querySelector("code");
      blocks.push(new Paragraph({
        spacing: { after: 220 },
        border: { left: { style: BorderStyle.SINGLE, size: 10, color: "C5CAD4" } },
        shading: { fill: "F4F5F7" },
        indent: { left: 240, right: 120 },
        children: [new TextRun({
          text: (code ? code.textContent : el.textContent || "").replace(/\n$/, ""),
          font: "Courier New",
          size: (ctx.settings.fontSize - 1) * 2,
        })],
      }));
    } else if (tag === "BLOCKQUOTE") {
      blocks.push(new Paragraph({
        spacing: { after: 180 },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "4F6EF7" } },
        indent: { left: 280 },
        children: [new TextRun({ text: (el.textContent || "").trim(), italics: true, color: "4B5061" })],
      }));
    } else if (tag === "HR") {
      blocks.push(new Paragraph({
        spacing: { before: 140, after: 140 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "E2E4EA" } },
      }));
    } else if (tag === "TABLE") {
      blocks.push(await convertTable(el, ctx));
    } else if (tag === "DIV" && el.classList.contains("callout")) {
      blocks.push(...convertCallout(el, ctx));
    } else if (tag === "SECTION" && el.classList.contains("footnotes")) {
      blocks.push(...convertFootnotes(el, ctx));
    } else {
      // Recurse into unknown elements
      const sub = await convertBlocks(Array.from(el.childNodes), ctx);
      blocks.push(...sub);
    }
  }

  return blocks.length ? blocks : [new Paragraph(" ")];
}

/* ── Callout rendering ── */

function convertCallout(el: HTMLElement, ctx: ConvertContext): Paragraph[] {
  const color = el.getAttribute("data-color") || "448AFF";
  const title = el.getAttribute("data-title") || "";
  const body = el.textContent || "";

  const paragraphs: Paragraph[] = [];

  // Title line
  if (title) {
    paragraphs.push(new Paragraph({
      spacing: { after: 80 },
      border: { left: { style: BorderStyle.SINGLE, size: 14, color } },
      shading: { fill: color + "18" },
      indent: { left: 280 },
      children: [new TextRun({ text: title, bold: true, color })],
    }));
  }

  // Body
  if (body) {
    paragraphs.push(new Paragraph({
      spacing: { after: 180 },
      border: { left: { style: BorderStyle.SINGLE, size: 14, color } },
      shading: { fill: color + "08" },
      indent: { left: 280 },
      children: [new TextRun({ text: body, color: "4B5061" })],
    }));
  }

  return paragraphs;
}

/* ── Footnotes rendering ── */

function convertFootnotes(el: HTMLElement, ctx: ConvertContext): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  paragraphs.push(new Paragraph({
    spacing: { before: 400, after: 200 },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "E2E4EA" } },
  }));

  const items = el.querySelectorAll("li");
  let idx = 1;
  for (const li of Array.from(items)) {
    const text = li.textContent?.replace(/\s*↩︎?\s*$/, "").trim() || "";
    paragraphs.push(new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({ text: `${idx}. `, bold: true, size: 20, color: "6B7280" }),
        new TextRun({ text, size: 20, color: "6B7280" }),
      ],
    }));
    idx++;
  }

  return paragraphs;
}

/* ── Inline conversion ── */

interface InlineStyle {
  bold?: boolean;
  italics?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  font?: string;
  highlight?: string;
  superScript?: boolean;
}

async function inlines(
  nodeList: NodeListOf<ChildNode> | ChildNode[],
  ctx: ConvertContext,
  style: InlineStyle = {},
): Promise<(TextRun | ExternalHyperlink | ImageRun)[]> {
  const out: (TextRun | ExternalHyperlink | ImageRun)[] = [];

  for (const node of Array.from(nodeList)) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent) {
        out.push(new TextRun({
          text: node.textContent,
          bold: style.bold,
          italics: style.italics,
          underline: style.underline ? {} : undefined,
          strike: style.strike,
          color: style.color,
          font: style.font,
          highlight: style.highlight as any,
          superScript: style.superScript,
        }));
      }
      continue;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) continue;
    const el = node as HTMLElement;
    const tag = el.tagName;

    if (tag === "STRONG" || tag === "B") {
      out.push(...(await inlines(el.childNodes, ctx, { ...style, bold: true })));
    } else if (tag === "EM" || tag === "I") {
      out.push(...(await inlines(el.childNodes, ctx, { ...style, italics: true })));
    } else if (tag === "S" || tag === "DEL") {
      out.push(...(await inlines(el.childNodes, ctx, { ...style, strike: true })));
    } else if (tag === "MARK") {
      out.push(...(await inlines(el.childNodes, ctx, { ...style, highlight: "yellow" })));
    } else if (tag === "SUP") {
      out.push(...(await inlines(el.childNodes, ctx, { ...style, superScript: true })));
    } else if (tag === "CODE") {
      out.push(new TextRun({
        text: el.textContent || "",
        font: "Courier New",
        bold: style.bold,
        italics: style.italics,
      }));
    } else if (tag === "A") {
      const href = el.getAttribute("href") || "";
      // Skip footnote back-references
      if (href.startsWith("#") && el.classList.contains("footnote-backref")) continue;
      out.push(new ExternalHyperlink({
        link: href,
        children: await inlines(el.childNodes, ctx, { ...style, color: "4F6EF7", underline: true }),
      }));
    } else if (tag === "BR") {
      out.push(new TextRun({ break: 1 }));
    } else if (tag === "IMG") {
      out.push(await convertImage(el, ctx));
    } else {
      out.push(...(await inlines(el.childNodes, ctx, style)));
    }
  }

  return out.length ? out : [new TextRun("")];
}

/* ── Image conversion ── */

async function convertImage(el: HTMLElement, ctx: ConvertContext): Promise<TextRun | ImageRun> {
  const rawSrc = decodeURI(el.getAttribute("src") || "");

  // Web images
  if (rawSrc.startsWith("http://") || rawSrc.startsWith("https://")) {
    return await fetchWebImage(rawSrc, el, ctx);
  }

  // Vault images
  const file = resolveImage(rawSrc, ctx);
  if (!file) {
    ctx.warnings.push(`未匹配到图片：${rawSrc}`);
    return new TextRun({ text: `[缺失图片：${rawSrc}]`, color: "D94545", italics: true });
  }

  let buf: ArrayBuffer;
  try {
    buf = await ctx.app.vault.readBinary(file);
  } catch {
    ctx.warnings.push(`无法读取图片：${file.path}`);
    return new TextRun({ text: `[无法读取：${file.path}]`, color: "D94545", italics: true });
  }

  return buildImageRun(buf, file.name, el, ctx);
}

async function fetchWebImage(url: string, el: HTMLElement, ctx: ConvertContext): Promise<TextRun | ImageRun> {
  try {
    const response = await requestUrl({ url });
    const buf = response.arrayBuffer;
    const contentType = response.headers["content-type"] || "";
    const filename = url.split("?")[0].split("/").pop() || "image.png";
    return buildImageRun(buf, filename, el, ctx, contentType);
  } catch (err: any) {
    ctx.warnings.push(`网络图片下载失败：${url}`);
    return new TextRun({ text: `[网络图片失败：${url}]`, color: "D94545", italics: true });
  }
}

function buildImageRun(
  buf: ArrayBuffer, filename: string, el: HTMLElement, ctx: ConvertContext, contentType?: string,
): ImageRun {
  const dim = getImageDimensions(buf);
  const explicitW = extractExplicitWidth(el.getAttribute("title"));

  let width: number, height: number;
  if (explicitW !== null) {
    width = Math.min(explicitW, dim.width);
    height = Math.max(1, Math.round(dim.height * (width / dim.width)));
  } else if (ctx.maxW === Infinity) {
    width = dim.width;
    height = dim.height;
  } else {
    width = Math.min(ctx.maxW, dim.width);
    height = Math.max(1, Math.round(dim.height * (width / dim.width)));
  }

  const imageType = contentType ? guessImageTypeFromUrl(filename, contentType) : guessImageType(filename);
  const alt = el.getAttribute("alt") || filename;

  ctx.matched++;
  return new ImageRun({
    data: buf,
    type: imageType,
    transformation: { width, height },
    altText: { title: alt, description: alt, name: filename },
  });
}

function resolveImage(rawSrc: string, ctx: ConvertContext): TFile | null {
  const direct = ctx.app.vault.getAbstractFileByPath(normalizePath(rawSrc));
  if (direct instanceof TFile) return direct;

  const linked = ctx.app.metadataCache.getFirstLinkpathDest(rawSrc, ctx.sourceFile.path);
  if (linked instanceof TFile) return linked;

  const sourceDir = ctx.sourceFile.parent?.path || "";
  if (sourceDir) {
    const relative = ctx.app.vault.getAbstractFileByPath(normalizePath(`${sourceDir}/${rawSrc}`));
    if (relative instanceof TFile) return relative;
  }

  const basename = rawSrc.split("/").pop() || rawSrc;
  if (basename !== rawSrc) {
    const byName = ctx.app.metadataCache.getFirstLinkpathDest(basename, ctx.sourceFile.path);
    if (byName instanceof TFile) return byName;
  }

  return null;
}

/* ── List conversion ── */

async function convertList(el: HTMLElement, ctx: ConvertContext, level: number): Promise<Paragraph[]> {
  const out: Paragraph[] = [];
  const ordered = el.tagName === "OL";
  let idx = parseInt(el.getAttribute("start") || "1", 10);

  for (const li of Array.from(el.children)) {
    if (li.tagName !== "LI") continue;
    const direct: ChildNode[] = [], nested: HTMLElement[] = [];
    for (const c of Array.from(li.childNodes)) {
      if (c.nodeType === Node.ELEMENT_NODE && ((c as HTMLElement).tagName === "UL" || (c as HTMLElement).tagName === "OL")) {
        nested.push(c as HTMLElement);
      } else {
        direct.push(c);
      }
    }
    const prefix = ordered ? `${idx}. ` : "\u2022 ";
    out.push(new Paragraph({
      spacing: { after: 100 },
      indent: { left: 240 + level * 280 },
      children: [new TextRun({ text: prefix, bold: true }), ...(await inlines(direct, ctx))],
    }));
    for (const n of nested) out.push(...(await convertList(n, ctx, level + 1)));
    idx++;
  }
  return out;
}

/* ── Table conversion ── */

async function convertTable(el: HTMLElement, ctx: ConvertContext): Promise<Table> {
  const rows: TableRow[] = [];
  for (const tr of Array.from(el.querySelectorAll("tr"))) {
    const cells: TableCell[] = [];
    for (const td of Array.from(tr.children)) {
      cells.push(new TableCell({
        children: [new Paragraph({ spacing: { after: 80 }, children: await inlines(td.childNodes, ctx) })],
      }));
    }
    rows.push(new TableRow({
      children: cells,
      tableHeader: tr.parentElement?.tagName === "THEAD",
    }));
  }
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
}

/* ── Helpers ── */

function escMd(v: string): string { return v.replace(/[[\]()]/g, "\\$&"); }

function extractExplicitWidth(title: string | null): number | null {
  const m = (title || "").match(/width=(\d+)/i);
  return m ? Number(m[1]) : null;
}
