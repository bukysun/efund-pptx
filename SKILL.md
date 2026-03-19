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
import copy, os

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE   = os.path.join(SKILL_DIR, "assets/template.pptx")

prs = Presentation(TEMPLATE)

# ⚠️ 目录页的表格不在 layout 中，必须在删除 slides 之前先提取 XML
toc_table_xml = None
for slide in prs.slides:
    if slide.slide_layout.name == '中文目录页，仅供目录页使用':
        for shape in slide.shapes:
            if shape.shape_type == 19:            # TABLE
                toc_table_xml = copy.deepcopy(shape._element)
                break
        break

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
# ── 中文封面（布局 0） ──────────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[0])

for shape in slide.shapes:
    if not hasattr(shape, 'placeholder_format'):
        continue
    idx = shape.placeholder_format.idx
    tf  = shape.text_frame

    if idx == 1:
        # 大标题：华文黑体 28pt，DEEP_BLUE，靠左
        add_text(tf, "易方达的历史和文化", first=True, size=28)
        # 副标题（可选）：华文黑体 加粗 22pt，DEEP_BLUE
        add_text(tf, "2024年新员工培训", size=22, bold=True)

    elif idx == 10:
        # 姓名行：华文黑体 14pt / 数字 Arial 14pt（apply_font 同时设两套字体）
        add_text(tf, "汇报人：XXX", first=True, size=14)
        add_text(tf, "2024年3月",   size=14)


# ── 英文封面（布局 7，纯英文） ───────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[7])

for shape in slide.shapes:
    if not hasattr(shape, 'placeholder_format'):
        continue
    idx = shape.placeholder_format.idx
    tf  = shape.text_frame

    if idx == 1:
        # 大标题：Arial 28pt，DEEP_BLUE
        add_text(tf, "E Fund Annual Report 2024",
                 first=True, cn_font=None, size=28)
        # 副标题（可选）：Arial 加粗 21pt
        add_text(tf, "Strategic Overview",
                 cn_font=None, size=21, bold=True)

    elif idx == 10:
        # 姓名行：Arial 14pt（纯英文，cn_font=None）
        add_text(tf, "Presenter: John Smith", first=True, cn_font=None, size=14)
        add_text(tf, "March 2024, Shanghai",  cn_font=None, size=14)


# ── 中英文封面（布局 7，首行中文） ──────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[7])

for shape in slide.shapes:
    if not hasattr(shape, 'placeholder_format'):
        continue
    idx = shape.placeholder_format.idx
    tf  = shape.text_frame

    if idx == 1:
        # 首行中文标题：华文黑体 28pt
        add_text(tf, "易方达年度报告 2024", first=True, size=28)
        # 第二行英文副标题：Arial 加粗 21pt（cn_font=None 只设英文字体）
        add_text(tf, "Annual Report", cn_font=None, size=21, bold=True)

    elif idx == 10:
        # 中文在上：华文黑体 加粗 14pt
        add_text(tf, "汇报人：XXX", first=True, size=14, bold=True)
        # 英文在下：Arial 12pt（cn_font=None）
        add_text(tf, "Presenter: XXX", cn_font=None, size=12)
        # 日期（仅英文）：Arial 加粗 12pt
        add_text(tf, "March 2024", cn_font=None, size=12, bold=True)
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

> **⚠️ 目录表格不在 layout 中，需从模板幻灯片克隆。**
> 必须在删除所有 slides 之前先提取 `toc_table_xml`（见 Quick Start）。
> 表格结构：7行×2列，col0=序号（Arial 28pt 亮蓝），col1=标题（华文黑体 18pt）。

```python
from pptx.oxml.ns import qn
from lxml import etree
import copy


def _set_toc_row(tr, number: str, title: str, color_hex: str):
    """更新一行目录：序号文字 + 标题文字 + 标题颜色（操作原始 XML tr 元素）。"""
    cells = tr.findall(qn('a:tc'))

    # Col 0：序号
    for rEl in cells[0].findall('.//' + qn('a:r')):
        t = rEl.find(qn('a:t'))
        if t is not None:
            t.text = number

    # Col 1：标题文字 + 颜色
    for rEl in cells[1].findall('.//' + qn('a:r')):
        t = rEl.find(qn('a:t'))
        if t is not None:
            t.text = title
        rPr = rEl.find(qn('a:rPr'))
        if rPr is None:
            rPr = etree.SubElement(rEl, qn('a:rPr'))
        old = rPr.find(qn('a:solidFill'))
        if old is not None:
            rPr.remove(old)
        # ⚠️ OOXML 要求 solidFill 必须在 latin/ea/cs 之前
        # 用 insert(0,...) 而不是 SubElement（SubElement 追加到末尾会被 PowerPoint 忽略）
        fill = etree.Element(qn('a:solidFill'))
        srgb = etree.SubElement(fill, qn('a:srgbClr'))
        srgb.set('val', color_hex)
        rPr.insert(0, fill)


def fill_toc_table(tbl, chapters: list[str], active_idx: int):
    """
    填写目录表格内容，自动调整行数，高亮当前章节，灰化其余章节。

    Args:
        tbl        : shape.table 对象
        chapters   : 章节标题列表（任意长度，无上限）
        active_idx : 当前章节的 0-based 索引（显示为 DEEP_BLUE，其余灰化）

    行数处理规则：
        章节数 < 模板行数 → 删除多余行（不留空行）
        章节数 > 模板行数 → 克隆最后一行样式向下扩展
    """
    tbl_xml  = tbl._tbl
    tr_list  = tbl_xml.findall(qn('a:tr'))
    n_have   = len(tr_list)
    n_need   = len(chapters)

    if n_need < n_have:
        # 删除多余行
        for tr in tr_list[n_need:]:
            tbl_xml.remove(tr)

    elif n_need > n_have:
        # 克隆最后一行，补足所需行数
        style_row = tr_list[-1]
        for _ in range(n_need - n_have):
            tbl_xml.append(copy.deepcopy(style_row))

    # 重新获取（行数可能已变化）
    tr_list = tbl_xml.findall(qn('a:tr'))

    for ri, tr in enumerate(tr_list):
        color_hex = '005096' if ri == active_idx else 'CCCCCC'
        _set_toc_row(tr, f'{ri + 1:02d}.', chapters[ri], color_hex)


def add_toc_slide(prs, toc_table_xml, chapters: list[str], active_idx: int):
    """
    添加目录页：克隆模板表格 XML → 填入内容 → 返回 slide。

    Args:
        prs           : Presentation 对象
        toc_table_xml : Quick Start 阶段提取的 deepcopy XML 元素
        chapters      : 章节标题列表（最多 7 项）
        active_idx    : 当前高亮章节的 0-based 索引
    """
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    # 将克隆的表格 XML 挂到 slide 的 shape tree
    cloned = copy.deepcopy(toc_table_xml)
    slide.shapes._spTree.append(cloned)

    # 找到刚插入的表格并填写内容
    for shape in slide.shapes:
        if shape.shape_type == 19:
            fill_toc_table(shape.table, chapters, active_idx)
            break

    return slide


# ── 示例调用 ────────────────────────────────────────────────
chapters = [
    "公司介绍",
    "投资策略",
    "产品线概览",
    "风险管理",
    "业绩回顾",
    "展望与规划",
    "风险提示及免责声明",
]

# 在"投资策略"这一章时，active_idx=1（第2章高亮，其余灰化）
toc_slide = add_toc_slide(prs, toc_table_xml, chapters, active_idx=1)
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

### 免责声明（读取 disclaimers/ 文件）

免责声明位于**封面页底部蓝色横幅**，固定文本，不得修改。

```python
import os

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))

def load_disclaimer(lang: str = "cn") -> str:
    """lang: 'cn' → 中文版封面用, 'en' → 英文版封面用"""
    path = os.path.join(SKILL_DIR, "disclaimers", f"disclaimer_{lang}.txt")
    with open(path, encoding="utf-8") as f:
        return f.read()

# 中文/中英文封面用中文声明，英文封面用英文声明
disclaimer_text = load_disclaimer("cn")   # 或 load_disclaimer("en")
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

## Font Utils

> ⚠️ **核心陷阱：`run.font.name` 只设置 Latin（西文）字体，不影响中文字符。**
> 中文字符走的是 XML 中的 `a:ea`（East Asian）属性，必须单独写入，
> 否则中文即使指定了"华文黑体"也会 fallback 到母版默认字体。

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

CN_FONT = "华文黑体"
EN_FONT = "Arial"


def apply_font(run, cn_font=CN_FONT, en_font=EN_FONT,
               size=None, bold=None, color=None):
    """
    同时设置 Latin（西文/数字）与 East Asian（中文）两套字体。

    python-pptx 的 run.font.name 只写 a:latin，中文字符需额外写 a:ea，
    本函数统一处理，确保中英混排时字体正确。

    用法：
        run = para.add_run()
        run.text = "易方达 2024"
        apply_font(run, size=28, bold=False, color=DEEP_BLUE)
    """
    rPr = run._r.get_or_add_rPr()

    if en_font:                                   # Latin：英文、数字
        el = rPr.find(qn('a:latin'))
        if el is None:
            el = etree.SubElement(rPr, qn('a:latin'))
        el.set('typeface', en_font)

    if cn_font:                                   # East Asian：中文
        el = rPr.find(qn('a:ea'))
        if el is None:
            el = etree.SubElement(rPr, qn('a:ea'))
        el.set('typeface', cn_font)

    if size  is not None: run.font.size      = Pt(size)
    if bold  is not None: run.font.bold      = bold
    if color is not None: run.font.color.rgb = color


def add_text(tf, text, *, first=False, align=PP_ALIGN.LEFT,
             cn_font=CN_FONT, en_font=EN_FONT,
             size=None, bold=None, color=DEEP_BLUE):
    """
    向 text_frame 追加一个段落并应用字体。

    Args:
        tf    : shape.text_frame
        text  : 段落文字
        first : True → 复用 tf.paragraphs[0]（清空已有内容），
                False → 新增段落（默认）
        align : 对齐方式，VI 规范默认靠左
        color : 默认 DEEP_BLUE（R0,G80,B150），可覆盖
    Returns:
        para  : 设置完成的 paragraph 对象
    """
    para = tf.paragraphs[0] if first else tf.add_paragraph()
    if first:
        para.clear()
    run = para.add_run()
    run.text = text
    apply_font(run, cn_font=cn_font, en_font=en_font,
               size=size, bold=bold, color=color)
    para.alignment = align
    return para
```

---

## Font Reference

| 元素 | 中文字体 | 英文/数字字体 | 字号 | 颜色 |
|------|---------|-------------|------|------|
| **[中文封面] 大标题** | 华文黑体 | — | 28pt | DEEP_BLUE |
| **[中文封面] 副标题** | 华文黑体 加粗 | — | 22pt | DEEP_BLUE |
| **[中文封面] 姓名/部门/日期** | 华文黑体 | Arial | 14pt | DEEP_BLUE |
| **[中英文封面] 首行中文时-首行** | 华文黑体 | — | 28pt | DEEP_BLUE |
| **[中英文封面] 首行中文时-第二行** | — | Arial 加粗 | 21pt | DEEP_BLUE |
| **[中英文封面] 首行英文时-首行** | — | Arial | 28pt | DEEP_BLUE |
| **[中英文封面] 首行英文时-第二行** | 华文黑体 加粗 | — | 22pt | DEEP_BLUE |
| **[中英文封面] 中文在上姓名行** | 华文黑体 加粗 | Arial | 14pt / 12pt | DEEP_BLUE |
| **[中英文封面] 英文在上姓名行** | 华文黑体 | Arial 加粗 | 12pt / 14pt | DEEP_BLUE |
| **[中英文封面] 日期/地点(仅英文)** | — | Arial 加粗 | 12pt | DEEP_BLUE |
| **[英文封面] 大标题** | — | Arial | 28pt（≥21pt） | DEEP_BLUE |
| **[英文封面] 副标题** | — | Arial 加粗 | 21pt | DEEP_BLUE |
| **[英文封面] 姓名/部门/日期** | — | Arial | 14pt | DEEP_BLUE |
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

**免责声明：** 位于封面底部蓝色横幅，固定文本，不得删除、不得修改、底色不得修改。
- 中文/中英文封面 → `disclaimers/disclaimer_cn.txt`
- 英文封面 → `disclaimers/disclaimer_en.txt`

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
