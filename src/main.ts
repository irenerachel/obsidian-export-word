import { Plugin, Notice, TFile, Menu, TAbstractFile } from "obsidian";
import * as path from "path";
import * as os from "os";
import * as fs from "fs";
import { convertToDocx } from "./converter";
import { ExportWordSettingTab, DEFAULT_SETTINGS, type ExportWordSettings } from "./settings";

const ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/><path d="M9 13l1.5 5 1.5-4 1.5 4 1.5-5"/></svg>`;

export default class ExportWordPlugin extends Plugin {
  settings: ExportWordSettings = DEFAULT_SETTINGS;

  async onload() {
    await this.loadSettings();
    this.addSettingTab(new ExportWordSettingTab(this.app, this));

    // 1. Command palette (Cmd+P)
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

    // 2. Ribbon icon (left sidebar, one click)
    this.addRibbonIcon("file-down", "导出当前笔记为 Word", () => {
      const file = this.app.workspace.getActiveFile();
      if (!file || file.extension !== "md") {
        new Notice("请先打开一篇 Markdown 笔记");
        return;
      }
      this.exportNote(file);
    });

    // 3. Right-click menu in file explorer
    this.registerEvent(
      this.app.workspace.on("file-menu", (menu: Menu, file: TAbstractFile) => {
        if (!(file instanceof TFile) || file.extension !== "md") return;
        menu.addItem((item) => {
          item.setTitle("导出为 Word")
            .setIcon("file-down")
            .onClick(() => this.exportNote(file));
        });
      })
    );

    // 4. Right-click menu in editor
    this.registerEvent(
      this.app.workspace.on("editor-menu", (menu: Menu) => {
        const file = this.app.workspace.getActiveFile();
        if (!file || file.extension !== "md") return;
        menu.addItem((item) => {
          item.setTitle("导出为 Word")
            .setIcon("file-down")
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

  private async exportNote(file: TFile) {
    new Notice("正在导出 Word...");

    try {
      const markdown = await this.app.vault.cachedRead(file);
      const title = this.resolveTitle(file, markdown);

      const { buffer, matched, warnings } = await convertToDocx(
        markdown, title, this.app, file, this.settings,
      );

      const outputPath = this.resolveOutputPath(file, title);
      const dir = path.dirname(outputPath);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(outputPath, Buffer.from(buffer));

      const basename = path.basename(outputPath);
      const locationLabel = this.getLocationLabel(outputPath);
      let msg = `导出成功：${basename}\n📍 已保存到：${locationLabel}\n🖼 ${matched} 张图片`;
      if (warnings.length) {
        msg += `\n⚠️ ${warnings.length} 个警告，详见控制台`;
        console.warn("[export-word] warnings:", warnings);
      }
      new Notice(msg, 8000);
    } catch (err: any) {
      console.error("[export-word]", err);
      new Notice(`导出失败：${err.message}`, 8000);
    }
  }

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

  private getLocationLabel(outputPath: string): string {
    const home = os.homedir();
    const desktop = path.join(home, "Desktop");
    const downloads = path.join(home, "Downloads");
    if (outputPath.startsWith(desktop)) return `桌面/${path.basename(outputPath)}`;
    if (outputPath.startsWith(downloads)) return `下载/${path.basename(outputPath)}`;
    return outputPath;
  }

  private resolveOutputPath(file: TFile, title: string): string {
    const safe = title.replace(/[\\/:*?"<>|]/g, "_");
    switch (this.settings.outputLocation) {
      case "desktop":
        return path.join(os.homedir(), "Desktop", `${safe}.docx`);
      case "downloads":
        return path.join(os.homedir(), "Downloads", `${safe}.docx`);
      case "same-folder": {
        const vaultBase = (this.app.vault.adapter as any).basePath as string;
        const parentDir = file.parent?.path || "";
        return path.join(vaultBase, parentDir, `${safe}.docx`);
      }
      case "custom": {
        const customDir = this.settings.customOutputPath || path.join(os.homedir(), "Desktop");
        return path.join(customDir, `${safe}.docx`);
      }
      default:
        return path.join(os.homedir(), "Desktop", `${safe}.docx`);
    }
  }
}
