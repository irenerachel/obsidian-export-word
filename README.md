# Export to Word — Obsidian 插件

一键把 Obsidian 笔记导出为 Word（.docx）文档，图片自动嵌入，无需安装 Pandoc。

> Export Obsidian notes to Word (.docx) with embedded images. No Pandoc required.

---

## 功能特性

- **一键导出** — 侧栏图标、右键菜单、命令面板，三种方式随你选
- **图片自动嵌入** — 支持 `![[图片.png]]`、`![[图片.png|320]]`、`![](路径)` 三种写法
- **原图不压缩** — 默认保持原图大小，像素数据原封不动写入 Word
- **自定义导出路径** — 桌面、下载文件夹、笔记同目录、或任意自定义路径（支持文件夹选择器）
- **自定义文件名格式** — 支持 `{filename}`、`{heading}`、`{date}`、`{time}` 变量组合
- **完整 Markdown 支持** — 标题、粗体、斜体、链接、代码块、引用、列表（含嵌套）、表格、分隔线
- **零依赖** — 不需要 Pandoc 或任何外部工具

## 安装

### 从 Obsidian 社区插件安装（审核中）

1. 打开 Obsidian → Settings → Community plugins → Browse
2. 搜索 **Export to Word**
3. 点击 Install → Enable

### 手动安装

1. 从 [Releases](https://github.com/irenerachel/obsidian-export-word/releases) 下载 `main.js` 和 `manifest.json`
2. 在 Obsidian 仓库的 `.obsidian/plugins/` 下新建文件夹 `obsidian-export-word`
3. 把下载的两个文件放进去
4. 重启 Obsidian → Settings → Community plugins → 启用 Export to Word

## 使用方法

### 导出笔记

打开一篇笔记，然后任选一种方式：

| 方式 | 操作 |
|------|------|
| 侧栏图标 | 点击左侧栏的下载图标 |
| 右键菜单 | 在文件列表或编辑器里右键 →「导出为 Word」|
| 命令面板 | `Cmd/Ctrl+P` → 输入 `Export current note to Word` |

导出完成后会弹出通知，显示文件名和保存路径。

### 设置

打开 Settings → Export to Word：

| 设置项 | 说明 |
|--------|------|
| 导出位置 | 桌面 / 下载文件夹 / 笔记同目录 / 自定义路径 |
| 图片尺寸 | 原图大小（默认）/ 限制最大宽度 |
| 文档标题 | 文件名 / 第一个标题 / 自定义格式 |

自定义标题格式支持变量：

| 变量 | 示例 |
|------|------|
| `{filename}` | 我的笔记 |
| `{heading}` | 笔记的第一个标题 |
| `{date}` | 2026-03-27 |
| `{time}` | 14-30 |

例如 `{filename} - {date}` → `我的笔记 - 2026-03-27.docx`

## 支持的语法

| Obsidian 语法 | Word 效果 |
|---------------|-----------|
| `# ~ ######` | 标题 1-6 级 |
| `**粗体**` `*斜体*` | 粗体、斜体 |
| `[链接](url)` | 可点击的超链接 |
| `` `代码` `` / ` ```代码块``` ` | 等宽字体 / 带背景代码段 |
| `> 引用` | 左侧蓝色竖线引用块 |
| `- 列表` / `1. 列表` | 无序/有序列表（支持嵌套） |
| `| 表格 |` | Word 表格 |
| `---` | 分隔线 |
| `![[图片.png]]` | 嵌入原图 |
| `![[图片.png\|320]]` | 限宽 320px，高度等比缩放 |
| `![alt](path)` | 标准 Markdown 图片 |
| `[[链接]]` / `[[链接\|别名]]` | 转为纯文本 |

## 开发

```bash
npm install
npm run build    # 生产构建
npm run dev      # 开发模式（watch）
```

## License

MIT
