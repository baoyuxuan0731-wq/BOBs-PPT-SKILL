# BOBs PPT Skill

一套把论文或实验材料做成学术组会幻灯片的流程。适用于 Codex，也可以把提示词复制给其他 AI 助手。每次汇报的材料和输出放到独立的本地目录。

| 步骤 | 提供什么 | 得到什么 |
| --- | --- | --- |
| 1. 确定要求 | 论文或实验资料、视觉模板、听众、时长、页数 | 逐页执行稿与图表来源 |
| 2. 生成 HTML | 已确认的执行稿 | 可在浏览器里翻页的完整预览 |
| 3. 评审和调整 | 要改的页码或总体反馈 | 确认版 HTML |
| 4. 转成 SVG | 确认版 HTML | 按页编号、可独立打开的 SVG |
| 5. 制作 PowerPoint | SVG、可选的本地逐页讲稿 | 可编辑 PPTX 与逐页检查结果 |

在 Codex 中可把本仓库作为技能目录，入口是 [SKILL.md](SKILL.md)。逐步推进时可复制 [PROMPTS.md](PROMPTS.md) 的相应提示词。用纯文本指导已有工作流时，脚本是可选的。

## 转换工具

### HTML 转 SVG

用 Chrome 或 Edge 打开自己制作的本地 HTML。默认画布为 1280 × 720，每页是一个 `[data-slide]` 元素。幻灯片 CSS 应以页面自身的 class 为作用域。默认导出布局为 `flex`，用 `data-export-display="block"` 等属性指定其他布局。图片使用内嵌 PNG/JPEG。页面控件放在画布之外，或用 `data-export-ignore` 标记。

在开发者工具的 Sources → Snippets 中创建本地脚本，粘贴 [Export-Slides.js](Export-Slides.js) 正文并运行。页面会出现逐页下载链接。下载后用相同尺寸对照 HTML。刷新页面可以关闭下载面板。

脚本针对固定画布、常见文本、纯色背景和矩形边框。渐变、变换、伪元素绘图、复杂裁剪和任意 SVG 图形应先简化，或由专业工具单独导出。对科学图表可直接嵌入已有 PNG，保留图例、条件和坐标轴。

### SVG 转可编辑 PowerPoint

需要 Windows 和桌面版 Microsoft PowerPoint。把 SVG 放到独立目录，按 `01.svg`、`02.svg` 等零填充编号。使用 PowerShell 执行，参数按本机路径填写：

```powershell
.\Convert-SvgToPowerPoint.ps1 -SvgDirectory '.\private\svg' -Output '.\private\report.pptx'
```

可加 `-NotesJson` 指定 UTF-8 JSON 文件，内容为与页数相同的字符串数组。用 `-PreviewDirectory` 指定目录时，脚本会重新打开成品并导出逐页 PNG。

转换器支持导出脚本生成的纯色矩形、平面文本和内嵌 PNG/JPEG。PPTX 中的文字和形状可单独编辑，图片可替换。柱状图若由矩形与标签构成，可编辑这些对象。需要可改数据的 PowerPoint 图表时，从原始数值在 PowerPoint 中另建图表。其他 SVG 可在桌面 PowerPoint 中按页插入并尝试「转换为形状」，文字能否继续作为文本编辑取决于转换结果。

## 发布范围

公开仓库只存流程、通用脚本和提示词。论文、数据、模板、实际汇报、讲稿和中间截图留在本地。提交前只暂存明确要公开的文件，并检查差异。忽略规则用于减少误提交，还应检查新文件和生成物是否已被追踪。
