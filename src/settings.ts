import { App, Platform, PluginSettingTab, Setting } from "obsidian";
import type ExportWordPlugin from "./main";

export interface ExportWordSettings {
  outputLocation: "desktop" | "downloads" | "same-folder" | "custom";
  customOutputPath: string;
  imageSizing: "original" | "max-width";
  maxImageWidth: number;
  titleSource: "filename" | "first-heading" | "custom";
  customTitleFormat: string;
  enableToc: boolean;
  enableCallouts: boolean;
  enableCoverPage: boolean;
  coverAuthor: string;
  enablePageNumbers: boolean;
  headerText: string;
  enableH1PageBreak: boolean;
  enableSmartFont: boolean;
  defaultFont: string;
  cjkFont: string;
  fontSize: number;
  enableWatermark: boolean;
  watermarkText: string;
}

export const DEFAULT_SETTINGS: ExportWordSettings = {
  outputLocation: Platform.isMobileApp ? "same-folder" : "desktop",
  customOutputPath: "",
  imageSizing: "original",
  maxImageWidth: 600,
  titleSource: "filename",
  customTitleFormat: "{filename}",
  enableToc: false,
  enableCallouts: true,
  enableCoverPage: false,
  coverAuthor: "",
  enablePageNumbers: true,
  headerText: "",
  enableH1PageBreak: false,
  enableSmartFont: true,
  defaultFont: "Calibri",
  cjkFont: "PingFang SC",
  fontSize: 12,
  enableWatermark: false,
  watermarkText: "",
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

    containerEl.createEl("h3", { text: "导出设置" });

    /* ── 导出位置 ── */
    const locationSetting = new Setting(containerEl)
      .setName("导出位置")
      .setDesc("生成的 .docx 文件保存到哪里")
      .addDropdown((d) => {
        if (Platform.isDesktopApp) {
          d.addOption("desktop", "桌面");
          d.addOption("downloads", "下载文件夹");
        }
        d.addOption("same-folder", "和笔记同目录");
        if (Platform.isDesktopApp) {
          d.addOption("custom", "自定义路径");
        }
        d.setValue(this.plugin.settings.outputLocation)
          .onChange(async (v) => {
            this.plugin.settings.outputLocation = v as ExportWordSettings["outputLocation"];
            await this.plugin.saveSettings();
            this.display();
          });
      });

    if (this.plugin.settings.outputLocation === "custom" && Platform.isDesktopApp) {
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
          btn.setButtonText("选择文件夹").onClick(async () => {
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
              // Fallback: user types manually
            }
          })
        );
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
          t.setValue(String(this.plugin.settings.maxImageWidth)).onChange(async (v) => {
            const n = parseInt(v, 10);
            if (n > 0) { this.plugin.settings.maxImageWidth = n; await this.plugin.saveSettings(); }
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
        .setDesc("可用变量：{filename} {heading} {date} {time}。例如：{filename} - {date}")
        .addText((t) =>
          t.setPlaceholder("{filename} - {date}").setValue(this.plugin.settings.customTitleFormat)
            .onChange(async (v) => { this.plugin.settings.customTitleFormat = v; await this.plugin.saveSettings(); })
        );
    }

    /* ── 页面布局 ── */
    containerEl.createEl("h3", { text: "页面布局" });

    new Setting(containerEl)
      .setName("封面页")
      .setDesc("在文档开头生成带标题、作者、日期的封面页")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enableCoverPage).onChange(async (v) => {
          this.plugin.settings.enableCoverPage = v; await this.plugin.saveSettings(); this.display();
        })
      );

    if (this.plugin.settings.enableCoverPage) {
      new Setting(containerEl)
        .setName("封面作者名")
        .setDesc("留空则不显示作者")
        .addText((t) =>
          t.setPlaceholder("阿真Irene").setValue(this.plugin.settings.coverAuthor)
            .onChange(async (v) => { this.plugin.settings.coverAuthor = v; await this.plugin.saveSettings(); })
        );
    }

    new Setting(containerEl)
      .setName("页码")
      .setDesc("在页脚显示页码")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enablePageNumbers).onChange(async (v) => {
          this.plugin.settings.enablePageNumbers = v; await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("页眉文字")
      .setDesc("留空则不显示页眉。支持 {title} 变量")
      .addText((t) =>
        t.setPlaceholder("").setValue(this.plugin.settings.headerText)
          .onChange(async (v) => { this.plugin.settings.headerText = v; await this.plugin.saveSettings(); })
      );

    new Setting(containerEl)
      .setName("H1 前自动分页")
      .setDesc("每个一级标题前插入分页符，适合多章节长文")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enableH1PageBreak).onChange(async (v) => {
          this.plugin.settings.enableH1PageBreak = v; await this.plugin.saveSettings();
        })
      );

    /* ── 高级功能 ── */
    containerEl.createEl("h3", { text: "高级功能" });

    new Setting(containerEl)
      .setName("自动生成目录")
      .setDesc("在文档开头插入目录（Table of Contents）")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enableToc).onChange(async (v) => {
          this.plugin.settings.enableToc = v; await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("渲染 Callout")
      .setDesc("将 > [!note] 等 Callout 语法转为带颜色标记的引用块")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enableCallouts).onChange(async (v) => {
          this.plugin.settings.enableCallouts = v; await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("水印")
      .setDesc("在文档中添加文字水印")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enableWatermark).onChange(async (v) => {
          this.plugin.settings.enableWatermark = v; await this.plugin.saveSettings(); this.display();
        })
      );

    if (this.plugin.settings.enableWatermark) {
      new Setting(containerEl)
        .setName("水印文字")
        .addText((t) =>
          t.setPlaceholder("CONFIDENTIAL").setValue(this.plugin.settings.watermarkText)
            .onChange(async (v) => { this.plugin.settings.watermarkText = v; await this.plugin.saveSettings(); })
        );
    }

    /* ── 样式 ── */
    containerEl.createEl("h3", { text: "Word 样式" });

    new Setting(containerEl)
      .setName("中英文智能字体")
      .setDesc("自动为中文和英文分别应用不同字体")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enableSmartFont).onChange(async (v) => {
          this.plugin.settings.enableSmartFont = v; await this.plugin.saveSettings(); this.display();
        })
      );

    new Setting(containerEl)
      .setName("英文/默认字体")
      .addText((t) =>
        t.setValue(this.plugin.settings.defaultFont).onChange(async (v) => {
          this.plugin.settings.defaultFont = v.trim() || "Calibri"; await this.plugin.saveSettings();
        })
      );

    if (this.plugin.settings.enableSmartFont) {
      new Setting(containerEl)
        .setName("中文字体")
        .setDesc("用于中日韩文字")
        .addText((t) =>
          t.setValue(this.plugin.settings.cjkFont).onChange(async (v) => {
            this.plugin.settings.cjkFont = v.trim() || "PingFang SC"; await this.plugin.saveSettings();
          })
        );
    }

    new Setting(containerEl)
      .setName("正文字号")
      .setDesc("Word 文档的正文字号（pt）")
      .addText((t) =>
        t.setValue(String(this.plugin.settings.fontSize)).onChange(async (v) => {
          const n = parseInt(v, 10);
          if (n > 0 && n <= 72) { this.plugin.settings.fontSize = n; await this.plugin.saveSettings(); }
        })
      );
  }
}
