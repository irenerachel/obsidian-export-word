import { Plugin, Notice, TFile, TFolder, Menu, TAbstractFile, Platform, normalizePath } from "obsidian";
import { convertToDocx } from "./converter";
import { ExportWordSettingTab, DEFAULT_SETTINGS, type ExportWordSettings } from "./settings";

export default class ExportWordPlugin extends Plugin {
  settings: ExportWordSettings = DEFAULT_SETTINGS;

  async onload() {
    await this.loadSettings();
    this.addSettingTab(new ExportWordSettingTab(this.app, this));

    // 1. Command: export single note
    this.addCommand({
      id: "export-current-note-to-word",
      name: "Export current note to Word",
      icon: "file-down",
      checkCallback: (checking) => {
        const file = this.app.workspace.getActiveFile();
        if (!file || file.extension !== "md") return false;
        if (!checking) this.exportNote(file);
        return true;
      },
    });

    // 2. Command: batch export folder
    this.addCommand({
      id: "export-folder-to-word",
      name: "Batch export folder to Word",
      icon: "folder-down",
      callback: () => {
        const file = this.app.workspace.getActiveFile();
        if (file?.parent) {
          this.exportFolder(file.parent);
        } else {
          new Notice("请先打开某个文件夹中的笔记");
        }
      },
    });

    // 3. Ribbon icon
    this.addRibbonIcon("file-down", "导出当前笔记为 Word", () => {
      const file = this.app.workspace.getActiveFile();
      if (!file || file.extension !== "md") {
        new Notice("请先打开一篇 Markdown 笔记");
        return;
      }
      this.exportNote(file);
    });

    // 4. File explorer right-click
    this.registerEvent(
      this.app.workspace.on("file-menu", (menu: Menu, file: TAbstractFile) => {
        if (file instanceof TFile && file.extension === "md") {
          menu.addItem((item) => {
            item.setTitle("导出为 Word").setIcon("file-down")
              .onClick(() => this.exportNote(file));
          });
        }
        if (file instanceof TFolder) {
          menu.addItem((item) => {
            item.setTitle("批量导出为 Word").setIcon("folder-down")
              .onClick(() => this.exportFolder(file));
          });
        }
      })
    );

    // 5. Editor right-click
    this.registerEvent(
      this.app.workspace.on("editor-menu", (menu: Menu) => {
        const file = this.app.workspace.getActiveFile();
        if (!file || file.extension !== "md") return;
        menu.addItem((item) => {
          item.setTitle("导出为 Word").setIcon("file-down")
            .onClick(() => this.exportNote(file));
        });
      })
    );
  }

  onunload() {}

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  /* ── Single note export ── */

  async exportNote(file: TFile, silent = false): Promise<boolean> {
    if (!silent) new Notice("正在导出 Word...");

    try {
      const markdown = await this.app.vault.cachedRead(file);
      const title = this.resolveTitle(file, markdown);

      const { buffer, matched, warnings, charCount, wordCount } = await convertToDocx(
        markdown, title, this.app, file, this.settings,
      );

      const outputName = await this.saveDocx(file, title, buffer);

      if (!silent) {
        const readMin = Math.max(1, Math.round(charCount / 500));
        let msg = `导出成功：${outputName}\n🖼 ${matched} 张图片 · 📝 ${charCount} 字 · ⏱ 约 ${readMin} 分钟`;
        if (warnings.length) {
          msg += `\n⚠️ ${warnings.length} 个警告，详见控制台`;
          console.warn("[export-word] warnings:", warnings);
        }
        new Notice(msg, 6000);
      }
      return true;
    } catch (err: any) {
      console.error("[export-word]", err);
      if (!silent) new Notice(`导出失败：${err.message}`, 8000);
      return false;
    }
  }

  /* ── Batch export ── */

  async exportFolder(folder: TFolder) {
    const mdFiles = folder.children.filter(
      (f): f is TFile => f instanceof TFile && f.extension === "md"
    );

    if (mdFiles.length === 0) {
      new Notice("该文件夹中没有 Markdown 笔记");
      return;
    }

    new Notice(`开始批量导出 ${mdFiles.length} 篇笔记...`);
    let success = 0;
    let fail = 0;

    for (const file of mdFiles) {
      const ok = await this.exportNote(file, true);
      if (ok) success++;
      else fail++;
    }

    let msg = `批量导出完成：${success} 篇成功`;
    if (fail > 0) msg += `，${fail} 篇失败`;
    new Notice(msg, 8000);
  }

  /* ── Save docx ── */

  private async saveDocx(file: TFile, title: string, arrayBuffer: ArrayBuffer): Promise<string> {
    const safe = title.replace(/[\\/:*?"<>|]/g, "_");
    const filename = `${safe}.docx`;
    const location = this.settings.outputLocation;

    // Mobile or "same-folder": save into vault
    if (Platform.isMobileApp || location === "same-folder") {
      const parentDir = file.parent?.path || "";
      const vaultPath = normalizePath(parentDir ? `${parentDir}/${filename}` : filename);
      const existing = this.app.vault.getAbstractFileByPath(vaultPath);
      if (existing instanceof TFile) {
        await this.app.vault.modifyBinary(existing, arrayBuffer);
      } else {
        await this.app.vault.createBinary(vaultPath, arrayBuffer);
      }
      return vaultPath;
    }

    // Desktop: use Node.js fs for external paths
    const path = require("path") as typeof import("path");
    const os = require("os") as typeof import("os");
    const fs = require("fs") as typeof import("fs");

    let dir: string;
    switch (location) {
      case "desktop":
        dir = path.join(os.homedir(), "Desktop");
        break;
      case "downloads":
        dir = path.join(os.homedir(), "Downloads");
        break;
      case "custom":
        dir = this.settings.customOutputPath || path.join(os.homedir(), "Desktop");
        break;
      default:
        dir = path.join(os.homedir(), "Desktop");
    }

    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const outputPath = path.join(dir, filename);
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer));

    // Return human-readable location
    const home = os.homedir();
    if (outputPath.startsWith(path.join(home, "Desktop"))) return `桌面/${filename}`;
    if (outputPath.startsWith(path.join(home, "Downloads"))) return `下载/${filename}`;
    return outputPath;
  }

  /* ── Title resolution ── */

  private resolveTitle(file: TFile, markdown: string): string {
    if (this.settings.titleSource === "first-heading") {
      const match = markdown.match(/^#{1,6}\s+(.+)$/m);
      if (match) return match[1].trim();
      return file.basename;
    }

    if (this.settings.titleSource === "custom") {
      const fmt = this.settings.customTitleFormat || "{filename}";
      const now = new Date();
      const heading = markdown.match(/^#{1,6}\s+(.+)$/m)?.[1]?.trim() || file.basename;
      const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
      const timeStr = `${String(now.getHours()).padStart(2, "0")}-${String(now.getMinutes()).padStart(2, "0")}`;
      return fmt
        .replace(/\{filename\}/g, file.basename)
        .replace(/\{heading\}/g, heading)
        .replace(/\{date\}/g, dateStr)
        .replace(/\{time\}/g, timeStr);
    }

    return file.basename;
  }
}
