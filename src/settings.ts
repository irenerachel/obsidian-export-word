import { App, PluginSettingTab, Setting } from "obsidian";
import type ExportWordPlugin from "./main";

export interface ExportWordSettings {
  outputLocation: "desktop" | "downloads" | "same-folder" | "custom";
  customOutputPath: string;
  imageSizing: "original" | "max-width";
  maxImageWidth: number;
  titleSource: "filename" | "first-heading" | "custom";
  customTitleFormat: string;
}

export const DEFAULT_SETTINGS: ExportWordSettings = {
  outputLocation: "desktop",
  customOutputPath: "",
  imageSizing: "original",
  maxImageWidth: 600,
  titleSource: "filename",
  customTitleFormat: "{filename}",
};

export class ExportWordSettingTab extends PluginSettingTab {
  plugin: ExportWordPlugin;

  constructor(app: App, plugin: ExportWordPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    /* ── 导出位置 ── */
    new Setting(containerEl)
      .setName("导出位置")
      .setDesc("生成的 .docx 文件保存到哪里")
      .addDropdown((d) =>
        d
          .addOption("desktop", "桌面")
          .addOption("downloads", "下载文件夹")
          .addOption("same-folder", "和笔记同目录")
          .addOption("custom", "自定义路径")
          .setValue(this.plugin.settings.outputLocation)
          .onChange(async (v) => {
            this.plugin.settings.outputLocation = v as ExportWordSettings["outputLocation"];
            await this.plugin.saveSettings();
            this.display();
          })
      );

    if (this.plugin.settings.outputLocation === "custom") {
      const pathSetting = new Setting(containerEl)
        .setName("自定义导出路径")
        .setDesc(this.plugin.settings.customOutputPath || "尚未选择，请点击右侧按钮选择文件夹")
        .addText((t) =>
          t
            .setPlaceholder("点击右侧按钮选择，或手动输入路径")
            .setValue(this.plugin.settings.customOutputPath)
            .onChange(async (v) => {
              this.plugin.settings.customOutputPath = v.trim();
              await this.plugin.saveSettings();
            })
        )
        .addButton((btn) =>
          btn
            .setButtonText("选择文件夹")
            .onClick(async () => {
              try {
                const electron = require("electron");
                const result = await electron.remote.dialog.showOpenDialog({
                  properties: ["openDirectory", "createDirectory"],
                  title: "选择导出文件夹",
                  defaultPath: this.plugin.settings.customOutputPath || undefined,
                });
                if (!result.canceled && result.filePaths.length > 0) {
                  this.plugin.settings.customOutputPath = result.filePaths[0];
                  await this.plugin.saveSettings();
                  this.display();
                }
              } catch {
                new (require("obsidian").Notice)("无法打开文件夹选择器，请手动输入路径");
              }
            })
        );

      // Make the text input wider
      const inputEl = pathSetting.controlEl.querySelector("input");
      if (inputEl) (inputEl as HTMLElement).style.width = "260px";
    }

    /* ── 图片尺寸 ── */
    new Setting(containerEl)
      .setName("图片尺寸")
      .setDesc("导出时图片的默认尺寸处理方式")
      .addDropdown((d) =>
        d
          .addOption("original", "原图大小（不缩放）")
          .addOption("max-width", "限制最大宽度")
          .setValue(this.plugin.settings.imageSizing)
          .onChange(async (v) => {
            this.plugin.settings.imageSizing = v as ExportWordSettings["imageSizing"];
            await this.plugin.saveSettings();
            this.display();
          })
      );

    if (this.plugin.settings.imageSizing === "max-width") {
      new Setting(containerEl)
        .setName("最大宽度（像素）")
        .setDesc("图片宽度超过此值时将等比缩放")
        .addText((t) =>
          t
            .setValue(String(this.plugin.settings.maxImageWidth))
            .onChange(async (v) => {
              const n = parseInt(v, 10);
              if (n > 0) {
                this.plugin.settings.maxImageWidth = n;
                await this.plugin.saveSettings();
              }
            })
        );
    }

    /* ── 文档标题 ── */
    new Setting(containerEl)
      .setName("文档标题来源")
      .setDesc("Word 文档的文件名从哪里取")
      .addDropdown((d) =>
        d
          .addOption("filename", "使用文件名")
          .addOption("first-heading", "使用笔记第一个标题")
          .addOption("custom", "自定义格式")
          .setValue(this.plugin.settings.titleSource)
          .onChange(async (v) => {
            this.plugin.settings.titleSource = v as ExportWordSettings["titleSource"];
            await this.plugin.saveSettings();
            this.display();
          })
      );

    if (this.plugin.settings.titleSource === "custom") {
      new Setting(containerEl)
        .setName("自定义标题格式")
        .setDesc(
          "可用变量：{filename} 文件名、{heading} 第一个标题、{date} 今天日期（2026-03-27）、{time} 当前时间（14-30）。" +
          "例如：{filename} - {date}"
        )
        .addText((t) =>
          t
            .setPlaceholder("{filename} - {date}")
            .setValue(this.plugin.settings.customTitleFormat)
            .onChange(async (v) => {
              this.plugin.settings.customTitleFormat = v;
              await this.plugin.saveSettings();
            })
        );
    }
  }
}
