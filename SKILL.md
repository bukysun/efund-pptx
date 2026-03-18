---
name: efund-pptx
description: Use when creating PowerPoint presentations using the E Fund (易方达) branded template. Triggers on any request to generate, fill, or modify slides that must follow E Fund's VI standards, font rules, color system, and layout conventions.
---

# 易方达 PPT Skill

Template: `assets/template.pptx` (10.00" × 5.63", 8 layouts)
Full rules: `efund_ppt_rules.md`
Disclaimers: `disclaimers/`

---

## Quick Start

```python
from pptx import Presentation
import shutil, os

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE   = os.path.join(SKILL_DIR, "assets/template.pptx")

prs = Presentation(TEMPLATE)

# ALWAYS delete all existing slides first
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[0]

# Add slides, fill content, save
prs.save("output.pptx")
```

---

## Layout Index

| idx | Name | Use For |
|-----|------|---------|
| 0 | 中文封面页 | 中文封面 |
| 1 | 中文目录页，仅供目录页使用 | 目录页 |
| 2 | 标题和内容 | 正文内容页 |
| 3 | 图形内容1 | 图文结合页 |
| 4 | 图形内容2 | 图文结合页 |
| 5 | 结尾页 | 中文封底 |
| 6 | 1_结尾页 | 英文封底 |
| 7 | 首页 | 英文/中英文封面 |

---

## Placeholder Map

### Layout 0 — 中文封面页
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 1 | OBJECT(7) | y=1.39" | 标题+副标题文本区 |
| 10 | OBJECT(7) | y=3.49" | 姓名/日期区 |

### Layout 1 — 目录页
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 0 | TITLE(1) | y=1.57" | 目录标题（固定为"目录"） |

> 目录内容填入页面内已有的 TABLE shape，不使用占位符。

### Layout 2 — 标题和内容
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 10 | OBJECT(7) | y=1.42", h=2.71" | 正文内容区 |

> 此布局无 TITLE 占位符，页面标题通过母版中 TITLE 形状呈现（位于 y=0.23"）。

### Layout 3 / 4 — 图形内容
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 0 | TITLE(1) | y=0.23" | 幻灯片标题 |
| 10 | OBJECT(7) | y=1.42", w=4.88" | 左侧文字/内容区 |

### Layout 5 — 结尾页（中文）
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 10 | OBJECT(7) | y=1.30" | 答谢词 |
| 11 | BODY(2) | y=2.74" | 联系信息 |

### Layout 7 — 首页（英文/中英文封面）
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 1 | OBJECT(7) | y=1.39" | 标题区 |
| 10 | OBJECT(7) | y=3.49" | 姓名/日期区 |

---

## Content Area Boundaries

```
正文页安全内容区（布局2/3/4）：
  左边距:  0.40"
  右边距:  0.40"（右侧内容终止于 ~9.60"）
  内容顶部: 1.42"（占位符 ph idx=10 的 top）
  内容底部: ~4.13"（1.42" + 2.71"）

左文右图分割线（大约）：
  左侧文字区: x=0.40", width≈4.73"
  右侧图表区: x=5.40", width≈4.34"
```

---

## Filling Content

### 封面页（中文/英文/中英文）

```python
slide = prs.slides.add_slide(prs.slide_layouts[0])  # 或 layout[7]

for shape in slide.shapes:
    if not hasattr(shape, 'placeholder_format'):
        continue
    try:
        idx = shape.placeholder_format.idx
    except:
        continue
    if idx == 1:
        tf = shape.text_frame
        tf.paragraphs[0].text = "易方达的历史和文化"      # 大标题
        p2 = tf.add_paragraph()
        p2.text = "2024年新员工培训"                       # 副标题（可选）
    elif idx == 10:
        tf = shape.text_frame
        tf.paragraphs[0].text = "XXX（姓名）"
        p2 = tf.add_paragraph()
        p2.text = "2024年3月"
```

### 正文页（有内文标题）

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor

slide = prs.slides.add_slide(prs.slide_layouts[2])

# 填写内容占位符（ph idx=10）
for shape in slide.shapes:
    try:
        if shape.placeholder_format.idx == 10:
            tf = shape.text_frame
            # 清空默认段落
            for para in tf.paragraphs:
                para.clear()

            content = [
                ("投资理念", 0),        # 一级标题，level=0
                ("客户至上：提供专业解决方案", 1),  # 一级文字，level=1
                ("风险管理", 0),
                ("全流程风险管理", 1),
            ]

            tf.paragraphs[0].text  = content[0][0]
            tf.paragraphs[0].level = content[0][1]
            for text, level in content[1:]:
                p = tf.add_paragraph()
                p.text  = text
                p.level = level
    except:
        pass
```

> **不要手动添加项目符号字符**，模板 master 已通过 level 自动处理缩进和符号。

### 目录页

```python
slide = prs.slides.add_slide(prs.slide_layouts[1])

# 目录内容写入页面内已有的 TABLE shape
for shape in slide.shapes:
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    if shape.shape_type == 19:  # TABLE
        table = shape.table
        # 根据实际目录项数填写表格单元格
        # table.cell(row, col).text = "..."
        break
```

### 封底页

```python
slide = prs.slides.add_slide(prs.slide_layouts[5])  # 中文封底

for shape in slide.shapes:
    try:
        idx = shape.placeholder_format.idx
        if idx == 10:
            shape.text_frame.paragraphs[0].text = "谢谢"
        elif idx == 11:
            tf = shape.text_frame
            tf.paragraphs[0].text = "联系我们："
            p = tf.add_paragraph()
            p.text = "XXX   Tel：+86(20)8510 ----"
            p = tf.add_paragraph()
            p.text = "        Email：XXX@efunds.com.cn"
    except:
        pass
```

### 风险提示页（读取 disclaimers/ 文件）

```python
import os

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))

def load_disclaimer(name: str) -> str:
    """name: mutual_fund_cn / mutual_fund_en / private_fund_cn /
             private_fund_en / annuity_cn / annuity_en /
             non_product_cn / non_product_en"""
    path = os.path.join(SKILL_DIR, "disclaimers", f"{name}.txt")
    with open(path, encoding="utf-8") as f:
        return f.read()

slide = prs.slides.add_slide(prs.slide_layouts[2])
for shape in slide.shapes:
    try:
        if shape.placeholder_format.idx == 10:
            shape.text_frame.paragraphs[0].text = load_disclaimer("mutual_fund_cn")
    except:
        pass
```

---

## VI Color Reference

```python
from pptx.dml.color import RGBColor

DEEP_BLUE   = RGBColor(0,   80,  150)   # 标题、主色
BRIGHT_BLUE = RGBColor(30,  185, 225)   # 目录序号、表头背景
DARK_GRAY   = RGBColor(60,  60,  60)    # 正文文字
MID_GRAY    = RGBColor(150, 150, 150)   # 连接线
LIGHT_GRAY  = RGBColor(204, 204, 204)   # 虚线、网格、灰化目录项
PALE_GRAY   = RGBColor(242, 242, 242)   # 表格背景
WHITE       = RGBColor(255, 255, 255)   # 目录主标题、表头字体
```

---

## Font Reference

| 元素 | 中文字体 | 英文/数字字体 | 字号 | 颜色 |
|------|---------|-------------|------|------|
| 封面大标题 | 华文黑体 | Arial | 28pt | DEEP_BLUE |
| 封面副标题 | 华文黑体 加粗 | Arial 加粗 | 22pt | DEEP_BLUE |
| 封面姓名/日期 | 华文黑体 | Arial | 14pt | DEEP_BLUE |
| 目录主标题 | 华文黑体 加粗 | — | 23pt | WHITE |
| 目录序号 | — | Arial | 28pt | BRIGHT_BLUE |
| 目录正文 | 华文黑体 | — | 18pt | DEEP_BLUE |
| 正文（少量文字） | 华文黑体 | — | 15pt | DARK_GRAY |
| 正文（大量文字） | 华文黑体 | — | 12pt | DARK_GRAY |
| 内文一级标题 | 华文黑体 加粗 | — | 15pt | DEEP_BLUE |
| 内文一级文字 | 华文黑体 | — | 12pt | DARK_GRAY |
| 内文二级标题 | 华文黑体 加粗 | — | 12pt | DEEP_BLUE |
| 内文二级文字 | 华文黑体 | — | 10pt | DARK_GRAY |
| 图表标题 | 华文黑体 | — | 10pt | DEEP_BLUE |
| 备注/数据来源 | 华文黑体 | Arial | 7pt | DARK_GRAY |
| 表头文字 | 华文黑体 加粗 | — | 适当 | WHITE |
| 封底答谢词 | 华文黑体 加粗 | — | 45pt | DEEP_BLUE |

---

## Key Rules (Summary)

**封面：**
- 所有文字靠左对齐，只能在设定文本框内编辑
- 免责声明底色、Logo 底色不得修改；封面图片不得替换
- 合作方 Logo 放在标题与日期之间，第一个与标题左侧对齐，高度 ≤ 大标题字高

**目录：**
- 内容必须填入给定表格，不得添加图片，格式不能自行调整
- 当前章节正常显示，其他章节灰化（LIGHT_GRAY）

**正文：**
- "上文下图" 或 "左文右图"
- 正文页下方横线、公司名称、页码不得删除
- 不建议出现三级标题

**图表：**
- 纯白底色，无边框，无渐变/阴影/3D 效果
- 网格线二选一（横或纵）
- 模型图连接线 1.5 磅 MID_GRAY，内容框使用直角方框

**表格：**
- 表头背景 BRIGHT_BLUE + 白色加粗字
- 内容居左，行高"中部对齐"
- 禁止使用规定色系外颜色（如红色）

**风险提示：** 对外使用时必须保留，不得删除。选择对应版本：
- 公募基金产品推介 → `mutual_fund_cn/en`
- 专户/养老金 → `private_fund_cn/en`
- 年金 → `annuity_cn/en`
- 非产品推介 → `non_product_cn/en`

---

## QA Checklist

```bash
# 1. 内容检查
python3 -c "
from pptx import Presentation
from markitdown import MarkItDown
md = MarkItDown()
print(md.convert('output.pptx').text_content)
"

# 检查残留占位符
python3 -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum"

# 2. 转图检查
python3 scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
ls slide-*.jpg
```

**视觉检查要点：**
- 文字是否超出内容区边界
- 标题/页脚/页码是否被内容遮挡
- 左右两侧留白是否保留
- 颜色是否符合 VI 规范
