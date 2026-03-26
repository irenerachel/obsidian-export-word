import { App, TFile, normalizePath } from "obsidian";
import MarkdownIt from "markdown-it";
import {
  BorderStyle, Document, ExternalHyperlink, HeadingLevel, ImageRun,
  Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType,
} from "docx";
import { getImageDimensions, guessImageType } from "./image-utils";
import type { ExportWordSettings } from "./settings";

const IMAGE_EXTS = new Set(["png", "jpg", "jpeg", "gif", "webp", "svg", "bmp", "tif", "tiff"]);

interface ConvertContext {
  app: App;
  sourceFile: TFile;
  settings: ExportWordSettings;
  maxW: number;
  matched: number;
  warnings: string[];
}

/* ── Public entry ── */

export async function convertToDocx(
  markdown: string,
  title: string,
  app: App,
  sourceFile: TFile,
  settings: ExportWordSettings,
): Promise<{ buffer: Buffer; matched: number; warnings: string[] }> {

  const maxW = settings.imageSizing === "original" ? Infinity : settings.maxImageWidth;
  const ctx: ConvertContext = { app, sourceFile, settings, maxW, matched: 0, warnings: [] };

  const normalized = normalizeObsidianMarkdown(markdown, ctx);
  const md = new MarkdownIt({ html: false, linkify: true, typographer: false, breaks: false });
  const html = md.render(normalized);

  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div id="r">${html}</div>`, "text/html");
  const root = doc.getElementById("r")!;
  const children = await convertBlocks(Array.from(root.childNodes), ctx);

  const wordDoc = new Document({
    creator: "Obsidian Export Word",
    title,
    description: "Exported from Obsidian",
    sections: [{ properties: {}, children }],
  });

  const buffer = await Packer.toBuffer(wordDoc);
  return { buffer, matched: ctx.matched, warnings: ctx.warnings };
}

/* ── Markdown normalization ── */

function normalizeObsidianMarkdown(text: string, ctx: ConvertContext): string {
  // Convert ![[embed]] syntax
  let out = text.replace(/!\[\[([^\]]+)\]\]/g, (_, raw: string) => {
    const parts = raw.trim().split("|");
    const target = (parts[0] || "").trim();
    const option = (parts[1] || "").trim();

    // Resolve via vault
    const file = ctx.app.metadataCache.getFirstLinkpathDest(target, ctx.sourceFile.path);
    if (!file) {
      ctx.warnings.push(`嵌入文件未找到：${target}`);
      return `**[缺失：${escMd(target)}]**`;
    }

    const ext = file.extension?.toLowerCase() || "";
    if (!IMAGE_EXTS.has(ext)) {
      return `**[嵌入文件：${escMd(file.basename)}]**`;
    }

    // Use vault-relative path as src; we'll resolve it in convertImage
    const uri = encodeURI(file.path).replace(/\(/g, "%28").replace(/\)/g, "%29");
    if (/^\d+$/.test(option)) {
      return `![${escMd(file.basename)}](${uri} "width=${option}")`;
    }
    return `![${escMd(option || file.basename)}](${uri})`;
  });

  // Convert [[wiki links]] to plain text
  out = out.replace(/\[\[([^\]]+)\]\]/g, (_, raw: string) => {
    const parts = raw.trim().split("|");
    const target = (parts[0] || "").trim();
    const alias = (parts[1] || "").trim();
    if (alias) return alias;
    const base = target.split("/").pop() || target;
    return base.replace(/\.[^.]+$/, "");
  });

  return out;
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
      const level = `HEADING_${tag.slice(1)}` as keyof typeof HeadingLevel;
      blocks.push(new Paragraph({
        heading: HeadingLevel[level],
        spacing: { after: 220 },
        children: await inlines(el.childNodes, ctx),
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
        })],
      }));
    } else if (tag === "BLOCKQUOTE") {
      blocks.push(new Paragraph({
        spacing: { after: 180 },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "4F6EF7" } },
        indent: { left: 280 },
        children: [new TextRun({
          text: (el.textContent || "").trim(),
          italics: true,
          color: "4B5061",
        })],
      }));
    } else if (tag === "HR") {
      blocks.push(new Paragraph({
        spacing: { before: 140, after: 140 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "E2E4EA" } },
      }));
    } else if (tag === "TABLE") {
      blocks.push(await convertTable(el, ctx));
    }
  }

  return blocks.length ? blocks : [new Paragraph(" ")];
}

/* ── Inline conversion ── */

interface InlineStyle {
  bold?: boolean;
  italics?: boolean;
  underline?: boolean;
  color?: string;
  font?: string;
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
          color: style.color,
          font: style.font,
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
    } else if (tag === "CODE") {
      out.push(new TextRun({
        text: el.textContent || "",
        font: "Courier New",
        bold: style.bold,
        italics: style.italics,
      }));
    } else if (tag === "A") {
      out.push(new ExternalHyperlink({
        link: el.getAttribute("href") || "",
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

  const imageType = guessImageType(file.name);
  const alt = el.getAttribute("alt") || file.basename;

  ctx.matched++;
  return new ImageRun({
    data: buf,
    type: imageType,
    transformation: { width, height },
    altText: { title: alt, description: alt, name: file.name },
  });
}

function resolveImage(rawSrc: string, ctx: ConvertContext): TFile | null {
  // 1. Try as vault-relative path directly (this is what our normalizeObsidianMarkdown produces)
  const direct = ctx.app.vault.getAbstractFileByPath(normalizePath(rawSrc));
  if (direct instanceof TFile) return direct;

  // 2. Try metadataCache link resolution (handles short names, attachments folder, etc.)
  const linked = ctx.app.metadataCache.getFirstLinkpathDest(rawSrc, ctx.sourceFile.path);
  if (linked instanceof TFile) return linked;

  // 3. Try relative to source file's directory
  const sourceDir = ctx.sourceFile.parent?.path || "";
  if (sourceDir) {
    const relative = ctx.app.vault.getAbstractFileByPath(normalizePath(`${sourceDir}/${rawSrc}`));
    if (relative instanceof TFile) return relative;
  }

  // 4. Try just the filename
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
    const direct: ChildNode[] = [];
    const nested: HTMLElement[] = [];

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
        children: [new Paragraph({
          spacing: { after: 80 },
          children: await inlines(td.childNodes, ctx),
        })],
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
