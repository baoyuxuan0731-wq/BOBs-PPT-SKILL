# BOBs PPT Skill

将论文或实验材料制作成可评审的学术汇报幻灯片。**给出仓库链接并要求做汇报时，默认先做第 1 步。** 本仓库的执行入口是 [SKILL.md](SKILL.md)。论文任务使用内置的 [nature-paper2ppt](vendor/nature-skills/skills/nature-paper2ppt/SKILL.md) 组织证据与图表，交付顺序遵循下表。

## 新对话这样开始

把下面一句连同论文发给助手。有模板或参考截图时一并提供。用途、听众、时长和重点可另加一句，页数有要求才填写。

> 按这个仓库的 SKILL.md 做文献汇报，加载内置 nature-paper2ppt 并实际查看版式截图。先给逐页执行稿，页数按内容。等我确认后，再一次生成整套 HTML。

只发送网页链接时，助手需要能读取仓库并把它用于当前任务。在 Codex 中，把仓库作为项目打开可使用 [AGENTS.md](AGENTS.md) 引导项目内任务。也可以在把完整目录纳入可用技能后，显式选用 bob-ppt-skill。直接复制的完整话术见 [PROMPTS.md](PROMPTS.md)。

## 按阶段交付

| 步骤 | 本轮交付 | 何时推进 |
| --- | --- | --- |
| 1. 核实材料与版式 | 逐页执行稿，写明关键原理、图源、实际参考的截图与用时 | 用户确认执行稿 |
| 2. 生成 HTML | 一次交付完整的 16:9 HTML 翻页预览 | 用户审阅 |
| 3. 审阅与修订 | 有反馈时修订 HTML 并交付整套新版 | 用户确认最新 HTML |
| 4. 导出 SVG | 可独立打开的逐页编号 SVG | 用户确认 SVG |
| 5. 制作 PowerPoint | 可编辑 PPTX、逐页讲稿和桌面 PowerPoint 验证结果 | 交付终稿 |
| 6. 沉淀技能（可选） | 通用流程、提示词或脚本 | 用户另行要求 |

用户回复“继续”且没有未处理的反馈时，只进入下一阶段。初次要求“做 PPT”也先交执行稿。用户明确指定阶段或跳过某步时，按指定范围执行。

## 模板与配套技能

从 [模板库](templates/) 选择一套 PPTX 源模板，实际查看适用的 [版式截图](templates/screenshots/README.md)，在执行稿中写明参考了哪张图、用在哪一页、借鉴什么。根据证据和推导安排内容，避免为了凑页数而拆散论点。图表和正文来自本次研究材料。

论文汇报按仓库内 [nature-paper2ppt](vendor/nature-skills/skills/nature-paper2ppt/SKILL.md) 的分类、叙事、选图与检查指导执行。它的上游入口、固定版本及许可见 [上游说明](vendor/nature-skills/UPSTREAM.md)。第三方文件保持与固定版本一致。

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

公开仓库包含流程、通用脚本、提示词和配套的第三方技能。模板库可以收录已确认可公开的截图和源模板。论文、数据、用户提供的私有模板、实际汇报、讲稿和中间截图留在本地。提交前只暂存明确要公开的文件，并检查差异。忽略规则用于减少误提交，还应检查新文件和生成物是否已被追踪。
