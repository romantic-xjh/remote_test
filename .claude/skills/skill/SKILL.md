---
name: ppt-generator
description: 根据用户提供的内容和需求，生成专业的 PowerPoint (PPTX) 演示文稿文件。
tags: [ppt, presentation, slides, office]
---

# PPT 生成器 (PPT Generator)

你是一个专业的 PowerPoint 演示文稿生成器。当用户要求生成 PPT、幻灯片、演示文稿时，使用此技能。

## 工作流程

### 第 1 步：收集需求

通过对话了解用户需求。尽可能一次问清楚以下信息：

1. **主题/标题** — 演示文稿的主题是什么？
2. **页数和内容** — 需要多少页幻灯片？每页的内容（文字、要点、数据）是什么？如果用户没有详细内容，帮用户根据主题自动生成内容大纲。
3. **风格偏好** — 颜色主题（默认专业蓝、商务灰、清新绿、活力橙等）、是否需要配图占位符。
4. **输出路径** — 保存位置，默认保存到桌面的 `presentation.pptx`。

### 第 2 步：生成 PPT

使用 **python-pptx** 库编写并执行 Python 脚本。以下是核心 API 参考：

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

prs = Presentation()
# 设置 16:9 宽屏尺寸
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# === 添加幻灯片 (layout: 0=title, 1=content, 2=section, 6=blank) ===
slide_layout = prs.slide_layouts[layout_index]  # 0-10
slide = prs.slides.add_slide(slide_layout)

# === 设置背景色 ===
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(0x1A, 0x2B, 0x4C)

# === 文本框 ===
from pptx.util import Inches, Pt
left, top, width, height = Inches(1), Inches(2), Inches(11), Inches(3)
textbox = slide.shapes.add_textbox(left, top, width, height)
tf = textbox.text_frame
tf.word_wrap = True

# 添加段落
p = tf.paragraphs[0]
p.text = "标题文字"
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
p.alignment = PP_ALIGN.CENTER

# 添加更多段落
p2 = tf.add_paragraph()
p2.text = "副标题"
p2.font.size = Pt(24)
p2.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
p2.alignment = PP_ALIGN.CENTER
p2.space_before = Pt(20)

# === 图片 ===
# slide.shapes.add_picture('image.png', left, top, width, height)

# === 表格 ===
from pptx.util import Inches
rows, cols = 3, 4
table_shape = slide.shapes.add_table(rows, cols, Inches(2), Inches(2), Inches(9), Inches(4))
table = table_shape.table
table.cell(0, 0).text = "Header"
# 设置表格列宽
table.columns[0].width = Inches(2)

# === 形状/矩形 ===
from pptx.enum.shapes import MSO_SHAPE
shape = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(11), Inches(0.5)
)
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0x00, 0x78, 0xD4)
shape.line.fill.background()  # 无边框

# === 保存 ===
prs.save('output.pptx')
```

### 第 3 步：设计原则

生成 PPT 时始终遵循以下设计原则：

1. **对比度** — 文字和背景必须有足够对比度。深色背景配浅色文字，浅色背景配深色文字。
2. **统一性** — 所有幻灯片使用一致的配色方案、字体和排版。
3. **简洁性** — 每页要点不超过 5-7 条，避免大段文字。
4. **视觉层次** — 标题 > 副标题 > 正文，用字号和颜色区分。
5. **专业配色** — 使用一套主题色贯穿始终。默认推荐：
   - 主色: #1A73E8 (Google Blue)
   - 深色: #1A2B4C (Dark Navy)
   - 浅色: #F8F9FA (Light Gray)
   - 强调: #FF6B35 (Orange accent)

### 第 4 步：确认交付

生成完成后：
1. 告诉用户文件保存路径
2. 简要描述生成的 PPT 内容（共几页，每页主题）
3. 询问是否需要修改（如调整颜色、增加页面、修改内容等）

## 示例交互

**用户：** "帮我做一个关于人工智能发展的 PPT，大概 10 页"

**助手：**
1. 确认主题 → 了解用户想覆盖的具体方面（历史、技术、应用、未来等）
2. 生成大纲 → 10 页幻灯片的具体结构
3. 编写 Python 脚本 → 使用 python-pptx 创建
4. 运行脚本 → 生成 PPTX 文件
5. 告知结果 → 路径、页数、每页概要

## 注意事项

- 始终先和用户确认 PPT 大纲再生成代码，避免返工。
- 如果用户提供的内容不够丰富，帮用户扩展和润色。
- 字体选择：中文优先使用 `微软雅黑`（Windows）或 `PingFang SC`（Mac），英文使用 `Arial` 或 `Calibri`。
- 生成的文件路径如果包含中文，确保 Python 脚本使用正确的编码处理。
