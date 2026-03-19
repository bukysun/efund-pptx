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
        # 大标题：华文黑体_易方达 28pt，DEEP_BLUE，靠左
        add_text(tf, "易方达的历史和文化", first=True, size=28)
        # 副标题（可选）：华文黑体_易方达 加粗 22pt，DEEP_BLUE
        add_text(tf, "2024年新员工培训", size=22, bold=True)

    elif idx == 10:
        # 姓名行：华文黑体_易方达 14pt / 数字 Arial 14pt（apply_font 同时设两套字体）
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
        # 首行中文标题：华文黑体_易方达 28pt
        add_text(tf, "易方达年度报告 2024", first=True, size=28)
        # 第二行英文副标题：Arial 加粗 21pt（cn_font=None 只设英文字体）
        add_text(tf, "Annual Report", cn_font=None, size=21, bold=True)

    elif idx == 10:
        # 中文在上：华文黑体_易方达 加粗 14pt
        add_text(tf, "汇报人：XXX", first=True, size=14, bold=True)
        # 英文在下：Arial 12pt（cn_font=None）
        add_text(tf, "Presenter: XXX", cn_font=None, size=12)
        # 日期（仅英文）：Arial 加粗 12pt
        add_text(tf, "March 2024", cn_font=None, size=12, bold=True)
```

### 正文页工具函数

```python
from lxml import etree

# ── Layout 2 无 title placeholder，需注入 ───────────────────
_TITLE_SP_XML = '''<p:sp
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="9999" name="TitleInjected"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="323850" y="211138"/><a:ext cx="8382375" cy="268287"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:t></a:t></a:r></a:p>
  </p:txBody>
</p:sp>'''
# ⚠️ 必须包含显式 <a:xfrm> 坐标；否则会继承 master 的窄标题框（约 2" 宽），导致文字被截断


def set_slide_title(slide, text, size=23, bold=True, color=WHITE):
    """
    为任意布局设置标题，自动处理 Layout 2 无 title placeholder 的问题。

    Layout 3/4 直接写 ph idx=0；
    Layout 2 的标题继承自 slide master，新建幻灯片时不存在于 slide XML，
    需注入一个带坐标的 <p:sp type="title"> 元素后才能填写文字。

    默认样式：WHITE 23pt 加粗（标题位于顶部蓝色横幅内，必须用白色）。
    """
    title_shape = slide.shapes.title
    if title_shape is None:
        slide.shapes._spTree.insert(2, etree.fromstring(_TITLE_SP_XML))
        title_shape = slide.shapes.title
    if title_shape:
        add_text(title_shape.text_frame, text, first=True,
                 cn_font=CN_FONT, en_font=None,
                 size=size, bold=bold, color=color)


# ── 多级正文 ─────────────────────────────────────────────────
BODY_STYLES = {
    0: dict(size=15, bold=True,  color=DEEP_BLUE, cn_font=CN_FONT, en_font=None),  # 一级标题
    1: dict(size=12, bold=False, color=DARK_GRAY,  cn_font=CN_FONT, en_font=None),  # 一级文字
    2: dict(size=10, bold=False, color=DARK_GRAY,  cn_font=CN_FONT, en_font=None),  # 二级文字
}

# 内容区可用高度（ph idx=10, h=2.71"）
CONTENT_HEIGHT_PT = 2.71 * 72  # ≈ 195 pt


def _set_para_spacing(para, spc_before_pt=None, line_spc_pct=None):
    """设置段落间距（直接操作 XML）。
    spc_before_pt : 段前距，单位 pt
    line_spc_pct  : 行距百分比，如 160 = 160%
    """
    pPr = para._p.get_or_add_pPr()
    if line_spc_pct is not None:
        lnSpc = pPr.find(qn('a:lnSpc'))
        if lnSpc is None:
            lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
        lnSpc.clear()
        etree.SubElement(lnSpc, qn('a:spcPct')).set('val', str(int(line_spc_pct * 1000)))
    if spc_before_pt is not None:
        spcBef = pPr.find(qn('a:spcBef'))
        if spcBef is None:
            spcBef = etree.SubElement(pPr, qn('a:spcBef'))
        spcBef.clear()
        etree.SubElement(spcBef, qn('a:spcPts')).set('val', str(int(spc_before_pt * 100)))


def add_body_content(tf, items, available_pt=CONTENT_HEIGHT_PT):
    """
    填写多级正文，段前空间动态分配：
      - 计算所有条目基础行高之和
      - 剩余空间按权重分配为段前间距：一级标题(非首条)权重3，其余权重1
      - 内容多时间距自动压缩（最小0），内容少时均匀撑开（最大20pt）

    Args:
        tf           : shape.text_frame（ph idx=10 的文本框）
        items        : list of (text, level) tuples
                       level 0 → 一级标题：华文黑体_易方达 加粗 15pt DEEP_BLUE
                       level 1 → 一级文字：华文黑体_易方达 12pt DARK_GRAY
                       level 2 → 二级文字：华文黑体_易方达 10pt DARK_GRAY
        available_pt : 内容区高度（pt），默认 195pt（Layout 2/3/4 的 ph idx=10 高度）

    Example:
        add_body_content(tf, [
            ('核心投资理念', 0),
            ('客户至上：为客户提供有针对性的产品', 1),
            ('价值导向，研究驱动', 1),
            ('风险管理', 0),
            ('全流程风险管理体系', 1),
        ])
    """
    _MIN_L0_SPC = 6.0   # 节标题（非首条）最小段前间距，pt
    _TARGET_LNS = {0: 140, 1: 160, 2: 160}   # 目标行距 %
    _FONT_SZ    = {0: 15,  1: 12,  2: 10}    # 各 level 字号

    n_l0_gaps  = sum(1 for i, (_, lv) in enumerate(items) if lv == 0 and i > 0)
    avail_text = available_pt - n_l0_gaps * _MIN_L0_SPC

    # 计算行距缩放比例：优先保持目标行距，溢出时等比压缩到 100%（单倍）
    base_target = sum(_FONT_SZ.get(lv,12) * _TARGET_LNS.get(lv,160)/100 for _,lv in items)
    base_min    = sum(_FONT_SZ.get(lv,12)                                 for _,lv in items)

    if base_target <= avail_text:
        scale     = 1.0
        remaining = avail_text - base_target
    elif base_min <= avail_text:
        scale     = avail_text / base_target   # 等比缩减，保底 100%
        remaining = 0.0
    else:
        scale     = 100 / max(_TARGET_LNS.values())
        remaining = 0.0

    actual_lns = {lv: max(100, int(_TARGET_LNS[lv] * scale)) for lv in [0, 1, 2]}

    # 剩余空间按权重分配段前间距（节标题权重 3，正文权重 1）
    weights = [0 if i == 0 else (3 if lv == 0 else 1)
               for i, (_, lv) in enumerate(items)]
    total_w = sum(weights)
    unit    = remaining / total_w if total_w > 0 else 0

    for i, (text, level) in enumerate(items):
        para = add_text(tf, text, first=(i == 0),
                        **BODY_STYLES.get(level, BODY_STYLES[1]))
        if i == 0:
            spc_before = 0.0
        elif level == 0:
            spc_before = _MIN_L0_SPC + min(20.0 - _MIN_L0_SPC, 3 * unit)
        else:
            spc_before = min(20.0, unit)
        _set_para_spacing(para, spc_before_pt=spc_before, line_spc_pct=actual_lns[level])
```

### 正文页（Layout 2 纯文字）

```python
slide = prs.slides.add_slide(prs.slide_layouts[2])

set_slide_title(slide, '投资理念与策略框架')

for shape in slide.shapes:
    try:
        if shape.placeholder_format.idx == 10:
            add_body_content(shape.text_frame, [
                ('核心投资理念', 0),
                ('客户至上：为客户提供有针对性的产品与专业解决方案', 1),
                ('价值导向：以基本面研究驱动投资决策', 1),
                ('风险管理', 0),
                ('全流程风险管理，覆盖投前、投中、投后各环节', 1),
                ('独立的风控团队，与投资团队形成制衡机制', 1),
            ])
    except:
        pass
```

### 正文页（Layout 3 左文右图）

```python
from pptx.util import Inches

slide = prs.slides.add_slide(prs.slide_layouts[3])

# 标题（ph idx=0，正常设置）
set_slide_title(slide, '资产管理规模增长')

# 左侧文字（ph idx=10）
for shape in slide.shapes:
    try:
        if shape.placeholder_format.idx == 10:
            add_body_content(shape.text_frame, [
                ('规模概况', 0),
                ('截至2024年底，管理规模突破2万亿元', 1),
                ('公募基金规模行业前三', 1),
                ('主要产品线', 0),
                ('权益类基金：占比约40%', 1),
                ('固收类基金：占比约45%', 1),
            ])
    except:
        pass

# 右侧区域（x≈5.40"，w≈4.34"）：手动添加图表/图片/文本框
# txBox = slide.shapes.add_textbox(Inches(5.4), Inches(1.42), Inches(4.34), Inches(2.71))
# 或通过 slide.shapes.add_chart(...) / slide.shapes.add_picture(...) 添加
```

> **不要手动添加项目符号字符**，模板 master 已通过 level 自动处理缩进和符号。

### 目录页

> **⚠️ 目录表格不在 layout 中，需从模板幻灯片克隆。**
> 必须在删除所有 slides 之前先提取 `toc_table_xml`（见 Quick Start）。
> 表格结构：7行×2列，col0=序号（Arial 28pt 亮蓝），col1=标题（华文黑体_易方达 18pt）。

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

> ⚠️ **核心陷阱：中文字体必须使用 `a:latin charset="-122"`，而非 `a:ea`。**
> 模板中实际的中文 run 使用的是 `<a:latin typeface="华文黑体_易方达" charset="-122"/>`（GBK 字符集），
> 用 `a:ea` 设置的字体在 PowerPoint 中不生效，中文会 fallback 到 Arial。

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

CN_FONT = "华文黑体_易方达"   # ⚠️ 公司定制字体，必须使用完整名称
EN_FONT = "Arial"


def apply_font(run, cn_font=CN_FONT, en_font=EN_FONT,
               size=None, bold=None, color=None):
    """
    设置中文/英文/混排字体。

    ⚠️ 关键规则：中文字符必须通过 a:latin charset="-122"（GBK 中文字符集）设置，
    这是模板 slide 中实际使用的方式。a:ea 字段不被 PowerPoint 识别为中文字体覆盖。

    参数说明：
        cn_font=CN_FONT, en_font=None  → 纯中文：a:latin charset=-122 + a:cs
        cn_font=None, en_font=EN_FONT  → 纯英文/数字：a:latin（无 charset）
        cn_font=CN_FONT, en_font=EN_FONT → 中英混排：a:latin=Arial + a:ea=华文黑体_易方达

    用法：
        run = para.add_run()
        run.text = "易方达 2024"
        apply_font(run, size=28, bold=False, color=DEEP_BLUE)
    """
    rPr = run._r.get_or_add_rPr()

    if cn_font and not en_font:
        # 纯中文：使用 a:latin charset="-122"（GBK），匹配模板 slide 的实际做法
        latin = rPr.find(qn('a:latin'))
        if latin is None:
            latin = etree.SubElement(rPr, qn('a:latin'))
        latin.set('typeface', cn_font)
        latin.set('charset', '-122')
        cs = rPr.find(qn('a:cs'))
        if cs is None:
            cs = etree.SubElement(rPr, qn('a:cs'))
        cs.set('typeface', cn_font)
        cs.set('charset', '-122')

    elif en_font and not cn_font:
        # 纯英文/数字
        latin = rPr.find(qn('a:latin'))
        if latin is None:
            latin = etree.SubElement(rPr, qn('a:latin'))
        latin.set('typeface', en_font)
        latin.attrib.pop('charset', None)

    elif cn_font and en_font:
        # 中英混排：a:latin=Arial 处理英文数字，a:ea=华文黑体_易方达 处理中文
        for tag, face in [('a:latin', en_font), ('a:ea', cn_font)]:
            el = rPr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rPr, qn(tag))
            el.set('typeface', face)

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
| **[中文封面] 大标题** | 华文黑体_易方达 | — | 28pt | DEEP_BLUE |
| **[中文封面] 副标题** | 华文黑体_易方达 加粗 | — | 22pt | DEEP_BLUE |
| **[中文封面] 姓名/部门/日期** | 华文黑体_易方达 | Arial | 14pt | DEEP_BLUE |
| **[中英文封面] 首行中文时-首行** | 华文黑体_易方达 | — | 28pt | DEEP_BLUE |
| **[中英文封面] 首行中文时-第二行** | — | Arial 加粗 | 21pt | DEEP_BLUE |
| **[中英文封面] 首行英文时-首行** | — | Arial | 28pt | DEEP_BLUE |
| **[中英文封面] 首行英文时-第二行** | 华文黑体_易方达 加粗 | — | 22pt | DEEP_BLUE |
| **[中英文封面] 中文在上姓名行** | 华文黑体_易方达 加粗 | Arial | 14pt / 12pt | DEEP_BLUE |
| **[中英文封面] 英文在上姓名行** | 华文黑体_易方达 | Arial 加粗 | 12pt / 14pt | DEEP_BLUE |
| **[中英文封面] 日期/地点(仅英文)** | — | Arial 加粗 | 12pt | DEEP_BLUE |
| **[英文封面] 大标题** | — | Arial | 28pt（≥21pt） | DEEP_BLUE |
| **[英文封面] 副标题** | — | Arial 加粗 | 21pt | DEEP_BLUE |
| **[英文封面] 姓名/部门/日期** | — | Arial | 14pt | DEEP_BLUE |
| 目录主标题 | 华文黑体_易方达 加粗 | — | 23pt | WHITE |
| 目录序号 | — | Arial | 28pt | BRIGHT_BLUE |
| 目录正文 | 华文黑体_易方达 | — | 18pt | DEEP_BLUE |
| 正文（少量文字） | 华文黑体_易方达 | — | 15pt | DARK_GRAY |
| 正文（大量文字） | 华文黑体_易方达 | — | 12pt | DARK_GRAY |
| 内文一级标题 | 华文黑体_易方达 加粗 | — | 15pt | DEEP_BLUE |
| 内文一级文字 | 华文黑体_易方达 | — | 12pt | DARK_GRAY |
| 内文二级标题 | 华文黑体_易方达 加粗 | — | 12pt | DEEP_BLUE |
| 内文二级文字 | 华文黑体_易方达 | — | 10pt | DARK_GRAY |
| 图表标题 | 华文黑体_易方达 | — | 10pt | DEEP_BLUE |
| 备注/数据来源 | 华文黑体_易方达 | Arial | 7pt | DARK_GRAY |
| 表头文字 | 华文黑体_易方达 加粗 | — | 适当 | WHITE |
| 封底答谢词 | 华文黑体_易方达 加粗 | — | 45pt | DEEP_BLUE |

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
