<p align="center">
  <img
    src="assets/brand/office-viewer-mark.svg"
    width="96"
    alt="Office Viewer"
  >
</p>

<h1 align="center">Office Viewer</h1>

<p align="center">
  <strong>在 Obsidian 中本地阅读和搜索 PowerPoint 演示文稿与 Word 文档。</strong>
</p>

<p align="center">
  本地阅读 · 搜索 PPTX 与 DOCX · 精确引用幻灯片 · 源文件保持不变
</p>

Office Viewer 可在桌面版 Obsidian 中直接打开本地 `.pptx` 演示文稿和 `.docx`
Word 文档，无需转 PDF、无需上传。PPTX 保留幻灯片引用与嵌入能力；DOCX 聚焦连续
阅读和当前文档搜索。

当前发布版本为 **0.2.5**。GitHub Release 会先更新；Obsidian 社区插件目录可能稍后刷新。

完整说明、开发与发布流程以英文 [README.md](README.md) 为准。本文是面向中文用户的摘要。

![在 Obsidian 中阅读本地 PPTX，含缩略图条](assets/readme/hero-reading-view.png)

## 安装

**Obsidian 社区插件（推荐）**

1. 打开 **设置 → 社区插件**。
2. 如需先启用社区插件，再打开 **浏览**。
3. 搜索 **Office Viewer**，安装并启用。

**GitHub Release**

从
[GitHub Releases](https://github.com/jerry4pan/obsidian-office-viewer-plugin/releases/latest)
下载 `main.js`、`manifest.json`、`styles.css` 到
`<Vault>/.obsidian/plugins/office-viewer/`，重载 Obsidian 后启用
**Office Viewer**。

## 反馈与支持

- **报告问题**：使用
  [缺陷报告表单](https://github.com/jerry4pan/obsidian-office-viewer-plugin/issues/new?template=bug-report.yml)。
- **提出改进建议**：使用
  [功能建议表单](https://github.com/jerry4pan/obsidian-office-viewer-plugin/issues/new?template=feature-request.yml)。
- **分享真实工作流**：发布一条
  [工作流故事](https://github.com/jerry4pan/obsidian-office-viewer-plugin/discussions/categories/workflow-stories)。
- **咨询使用问题**：前往
  [GitHub Discussions 问答区](https://github.com/jerry4pan/obsidian-office-viewer-plugin/discussions/categories/q-a)。
- **报告安全漏洞**：按 `SECURITY.md` 使用私密报告渠道。

报告缺陷前请使用最新版本，并在可用时附上 **复制诊断摘要**。不要上传机密演示文稿或
Word 文档、包含敏感内容的截图、文件名、路径、正文原文或图像。

## 功能概要

- 在库中打开 `.pptx`，按页阅读；支持按钮、页码跳转与方向键。
- 在库中打开 `.docx`，连续阅读最终版本的正文标题、段落、列表、表格、内嵌图片
  与安全链接；页眉、页脚、批注、文本框、删除文字和隐藏文字不在 DOCX 视图范围内。
- 常见位图、可本地转换的 Windows 图元文件（WMF/EMF），以及带可用缓存数据的 Office
  图表，会以预览图显示；不保证与 Word 打印版式一致。
- 在 DOCX 中按 `Cmd+F` 或 `Ctrl+F` 打开文档搜索面板并聚焦输入框，`Escape` 关闭。
  搜索仅覆盖当前文档可见正文；结果按段落显示（含匹配次数与摘要），并跳转到当前
  会话高亮。Word 默认使用完整排版。在标签页或嵌入预览中点击工具栏的**简化视图**图标，
  可在两种视图之间切换。无法呈现的内容会显示占位与文档级提示。
- DOCX 暂不提供稳定段落引用、伴生笔记、阅读位置持久化或全库索引。
- 缩略图条可滚动、渐进渲染，并可拖拽调整宽度。
- 支持全屏阅读；预览不够时可用 **在默认应用中打开**（如 PowerPoint、Keynote、Word、Pages）。
- 可复制稳定的**幻灯片引用**，即使演示文稿重排，也能按原生幻灯片 ID 返回同一页。
- 可复制**幻灯片嵌入**，在 Markdown 阅读视图中显示来源驱动的单页内容；独占一行的规范嵌入在 Live Preview 中也会原地渲染。
- 选中当前幻灯片中可见的文字后，可使用**复制所选文字并附来源**；普通
  `Cmd+C` / `Ctrl+C` 保持不变，显式操作会在所选文字后附上稳定的幻灯片引用。

![Markdown 阅读视图中的单页幻灯片嵌入](assets/readme/slide-embed.png)

- 用 `![[演示文稿.pptx]]` 或 `![[文档.docx]]` 把整份文件嵌入笔记。PPTX 保留原始幻灯片比例，可在嵌入卡片内翻页；DOCX 按原始纸张宽度、页边距和文档记录的分页信息预览正文。点击右上角的外部打开图标可交给系统默认应用。
- 在 Live Preview 中，请让整份文件嵌入或规范 PPTX 单页嵌入独占一行。光标或选区触及嵌入时可编辑原始语法；Source 模式始终显示语法。
- 被引用页已删除或源演示文稿缺失时会明确提示，不会静默退回旧页码。
- 可展开当前页**讲者备注**面板阅读作者备注，并可连同稳定幻灯片引用一并复制。

![演示文稿内容搜索与匹配结果](assets/readme/presentation-search.png)

- 在打开的 PPTX 中按 `Cmd+F` 或 `Ctrl+F`，可搜索当前演示文稿自己编写的标题、正文、文本框、形状文字、表格单元格和讲者备注；有备注时支持全部 / 幻灯片 / 讲者备注范围筛选。
- 搜索查询、源文字、片段和结果仅存在于当前视图会话，不会持久化。图片、母版/版式文字、图表和 SmartArt 不在搜索范围内。
- 默认记住阅读位置（仅存库内相对路径、文件大小、修改时间、页码与时间戳）。关闭该设置不会清除已认领的伴生笔记路径对。
- 可从已打开的 `.pptx` 查看器显式**打开伴生笔记**：同目录同名 Markdown；不存在则创建最小标题与来源链接，已存在则原样认领。普通打开 PPTX 不会写库；源 PPTX 保持只读。
- **诊断摘要** 默认关闭；开启后可显示可检测的兼容性提示，并复制不含内容的诊断信息。
- 界面跟随 Obsidian 语言，支持英文、简体中文、繁体中文；其他语言回退到英文。
- 全程本地；源 PPTX 与 DOCX 只读、不上传、不收集遥测。伴生笔记仅在显式操作时写入 Markdown；插件数据只保存认领关系的两条库内相对路径，不保存 PPTX/DOCX 搜索内容。

## 当前边界

- 正式版解析 `.pptx` 与 `.docx`，旧版文件可交给系统默认应用打开。
  macOS 开发版可用本机 LibreOffice 预览 Office 97–2003 的 `.ppt` 和 `.doc`，
  包括 `![[演示文稿.ppt]]` 和 `![[文档.doc]]`。
- 只读、本地；不回写源文件。
- 仅桌面版 Obsidian；不支持手机与平板。
- 正式版在 Obsidian 中直接解析 OOXML。开发版的旧格式转换在本地临时目录进行。
- 正常阅读不会上传演示或文档内容、不会自动跟随外部关系、不会执行宏/脚本、不会发网络请求。
- 渲染是可读预览，不是像素级 PowerPoint 或 Word 还原；嵌入 SVG、不受支持的图表类型等
  可能降级或显示占位。预览不可信时请用 **在默认应用中打开**。
- 可检测的不支持媒体与缺失字体，仅在开启 **诊断摘要** 时显示警告；仍可能存在未知的 Office 差异。
- 编辑、保存、动画、OCR、全库搜索、主画面搜索高亮、DOCX 段落引用、与正文混排或多嵌入同行的 Live Preview 渲染、遥测、账号、授权与云服务均不在范围内。

隐私与安全细节见 `PRIVACY.md` 与 `SECURITY.md`。

Office Viewer 是 [Jerry Pan](https://github.com/jerry4pan) 的开源项目。
