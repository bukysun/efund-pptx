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

### Layout 3 — 图形内容1（左文右图）
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 0 | TITLE(1) | y=0.23" | 幻灯片标题 |
| 10 | OBJECT(7) | y=1.42", w=4.88" | 左侧文字区（add_body_content） |

> 右侧图表区：x=5.40", y=1.36", w=4.34", h=2.71"，手动 add_chart / add_picture。

### Layout 4 — 图形内容2（双图）
| 区域 | Position | Use |
|------|----------|-----|
| 左图 | x=0.40", y=1.49", w=4.32", h=2.54" | 左侧图表 |
| 右图 | x=4.71", y=1.52", w=4.34", h=2.44" | 右侧图表 |
| 左小标题 | x=0.48", y=1.32" | 10pt DEEP_BLUE |
| 右小标题 | x=5.08", y=1.32" | 10pt DEEP_BLUE |
| 左数据来源 | x=0.40", y=3.86" | 7pt MID_GRAY |
| 右数据来源 | x=5.57", y=3.82" | 7pt MID_GRAY |

> ph idx=10 存在但实际不使用；所有内容通过 `add_layout4_slide()` 手动放置。

### Layout 5 — 结尾页（中文）
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 10 | OBJECT(7) | y=1.30", w=5.31" | 答谢词（45pt 加粗 DEEP_BLUE） |
| 11 | BODY(2) | y=2.74", w=3.00" | 联系信息（12pt DEEP_BLUE 行距170%，可省略） |

### Layout 6 — 结尾页（英文）
| ph idx | Type | Position | Use |
|--------|------|----------|-----|
| 10 | OBJECT(7) | y=1.30", w=5.31" | 答谢词（Arial 45pt 加粗 DEEP_BLUE） |
| 4294967295 | — | y=3.85", w=5.73" | 联系信息（Arial 12pt DEEP_BLUE 行距170%，可省略） |

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
        add_text(tf, "易方达的历史和文化", first=True, size=28, bold=False)
        # 副标题（可选）：华文黑体_易方达 加粗 22pt，DEEP_BLUE
        add_text(tf, "2024年新员工培训", size=22, bold=True)

    elif idx == 10:
        # 姓名行：华文黑体_易方达 14pt / 数字 Arial 14pt（apply_font 同时设两套字体）
        add_text(tf, "汇报人：XXX", first=True, size=14, bold=False)
        add_text(tf, "2024年3月",   size=14, bold=False)


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
                 first=True, cn_font=None, en_font=EN_FONT, size=28, bold=False)
        # 副标题（可选）：Arial 加粗 21pt
        add_text(tf, "Strategic Overview",
                 cn_font=None, en_font=EN_FONT, size=21, bold=True)

    elif idx == 10:
        # 姓名行：Arial 14pt（纯英文，cn_font=None）
        add_text(tf, "Presenter: John Smith", first=True, cn_font=None, en_font=EN_FONT, size=14, bold=False)
        add_text(tf, "March 2024, Shanghai",  cn_font=None, en_font=EN_FONT, size=14, bold=False)


# ── 中英文封面（布局 7，首行中文） ──────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[7])

for shape in slide.shapes:
    if not hasattr(shape, 'placeholder_format'):
        continue
    idx = shape.placeholder_format.idx
    tf  = shape.text_frame

    if idx == 1:
        # 首行中文标题：华文黑体_易方达 28pt
        add_text(tf, "易方达年度报告 2024", first=True, size=28, bold=False)
        # 第二行英文副标题：Arial 加粗 21pt（cn_font=None 只设英文字体）
        add_text(tf, "Annual Report", cn_font=None, en_font=EN_FONT, size=21, bold=True)

    elif idx == 10:
        # 姓名行：华文黑体_易方达 14pt
        add_text(tf, "汇报人：XXX", first=True, size=14, bold=False)
        # 英文在下：Arial 12pt（cn_font=None）
        add_text(tf, "Presenter: XXX", cn_font=None, en_font=EN_FONT, size=12, bold=False)
        # 日期（仅英文）：Arial 12pt
        add_text(tf, "March 2024", cn_font=None, en_font=EN_FONT, size=12, bold=False)
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

    ⚠️ Layout 3/4 的 title placeholder 默认宽度仅 4.3"（左半区），会导致标题提前换行。
    此函数自动将其扩展至 9.17"（蓝色横幅全宽）。
    扩展时必须同时设置 left/top/height/width 四个属性，否则 <a:xfrm> 新建时
    x/y 默认为 0，标题会错位到左上角。
    """
    title_shape = slide.shapes.title
    if title_shape is None:
        slide.shapes._spTree.insert(2, etree.fromstring(_TITLE_SP_XML))
        title_shape = slide.shapes.title
    if title_shape:
        # Layout 3/4 title placeholder 默认只有 4.3"，扩展到蓝色横幅全宽
        # ⚠️ 必须同时设置四个属性，只改 width 会导致 x/y 归零（标题跑到左上角）
        if title_shape.width < Inches(8):
            _l = title_shape.left; _t = title_shape.top; _h = title_shape.height
            title_shape.left = _l; title_shape.top = _t
            title_shape.height = _h; title_shape.width = Inches(9.17)
        tf = title_shape.text_frame
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        add_text(tf, text, first=True,
                 cn_font=CN_FONT, en_font=None,
                 size=size, bold=bold, color=color)


# ── 多级正文 ─────────────────────────────────────────────────
BODY_STYLES = {
    0: dict(size=15, bold=True,  color=DEEP_BLUE, cn_font=CN_FONT, en_font=EN_FONT),  # 一级标题
    1: dict(size=12, bold=False, color=DARK_GRAY,  cn_font=CN_FONT, en_font=EN_FONT),  # 一级文字
    2: dict(size=10, bold=False, color=DARK_GRAY,  cn_font=CN_FONT, en_font=EN_FONT),  # 二级文字
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


# 项目符号悬挂缩进（来自 ppt_template.pptx 实测值：4.8mm = 171450 EMU）
_BULLET_INDENT_EMU = 171450


def _set_para_bullet(para, enabled=True, level=1):
    """
    为段落设置/取消项目符号（黑色小圆点 •，Arial 字体，100% 字高）。

    enabled=True  → 添加 •，按 level 设置悬挂缩进（hanging indent）
    enabled=False → 显式置 buNone（一级标题/无符号段落）

    缩进规则（来自 ppt_template.pptx 实测）：
      level 1: marL=171450 EMU (4.8mm), indent=-171450 EMU（悬挂）
      level 2: marL=342900 EMU (9.5mm), indent=-171450 EMU（悬挂，更深缩进）
    """
    pPr = para._p.get_or_add_pPr()

    # 清除已有 bullet 相关子元素
    for tag in [qn('a:buNone'), qn('a:buChar'), qn('a:buAutoNum'),
                qn('a:buFont'), qn('a:buSzPct')]:
        el = pPr.find(tag)
        if el is not None:
            pPr.remove(el)

    if not enabled:
        etree.SubElement(pPr, qn('a:buNone'))
        pPr.attrib.pop('marL', None)
        pPr.attrib.pop('indent', None)
        return

    # 悬挂缩进：文字左边距 = level * 4.8mm，bullet 向左突出 4.8mm
    pPr.set('marL', str(_BULLET_INDENT_EMU * level))
    pPr.set('indent', str(-_BULLET_INDENT_EMU))

    buFont = etree.SubElement(pPr, qn('a:buFont'))
    buFont.set('typeface', 'Arial')

    buSzPct = etree.SubElement(pPr, qn('a:buSzPct'))
    buSzPct.set('val', '100000')   # 100% 字高

    buChar = etree.SubElement(pPr, qn('a:buChar'))
    buChar.set('char', '•')


# ⚠️ 正文填充函数完整实现已移至 docs/add_body_content.md
# 生成含正文内容的幻灯片前，必须先 Read("docs/add_body_content.md") 获取完整代码。

def add_body_content(tf, items, available_pt=CONTENT_HEIGHT_PT,
                     available_width_pt=None):
    """
    单文本框多级正文。items: list of (text, level)，level 0/1/2。
    ▶ 完整实现见 docs/add_body_content.md，使用前请 Read 该文件。
    """
    raise NotImplementedError("Read docs/add_body_content.md and paste implementation")


_LINE_HT_PT = {0: 15 * 1.40, 1: 12 * 1.55, 2: 10 * 1.55}   # pt per line（含行距）


def add_body_content_blocks(slide, items,
                            left=_PH10_LEFT, top=_PH10_TOP,
                            width=_PH10_WIDTH[2],
                            available_pt=CONTENT_HEIGHT_PT,
                            min_gap_pt=14.0,
                            max_gap_pt=22.0):
    """
    多文本框正文布局，推荐用于含 2+ 个 L0 节的正文页。
    内部在内容足够时自动使用多 textbox，否则退回 add_body_content。
    ▶ 完整实现见 docs/add_body_content.md，使用前请 Read 该文件。
    """
    raise NotImplementedError("Read docs/add_body_content.md and paste implementation")
```

### 表格工具函数

```python
def _rgb_hex(rgb: RGBColor) -> str:
    """RGBColor → 6位十六进制字符串，用于 OOXML srgbClr val 属性。"""
    return f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'


def _set_cell_fill(cell, rgb: RGBColor):
    """设置单元格背景色（solidFill）。"""
    tc = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None:
        tcPr = etree.SubElement(tc, qn('a:tcPr'))
    for tag in [qn('a:solidFill'), qn('a:gradFill'), qn('a:noFill')]:
        el = tcPr.find(tag)
        if el is not None:
            tcPr.remove(el)
    solidFill = etree.SubElement(tcPr, qn('a:solidFill'))
    etree.SubElement(solidFill, qn('a:srgbClr')).set('val', _rgb_hex(rgb))


def _set_cell_valign(cell, anchor='ctr'):
    """设置单元格垂直对齐。anchor: 't'=顶部, 'ctr'=中部, 'b'=底部"""
    tc = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None:
        tcPr = etree.SubElement(tc, qn('a:tcPr'))
    tcPr.set('anchor', anchor)


def style_header_cell(cell, text, size=11):
    """表头单元格：BRIGHT_BLUE 背景 + 白色加粗华文黑体 + 居中对齐。"""
    _set_cell_fill(cell, BRIGHT_BLUE)
    _set_cell_valign(cell, 'ctr')
    add_text(cell.text_frame, text, first=True, align=PP_ALIGN.CENTER,
             cn_font=CN_FONT, en_font=None, size=size, bold=True, color=WHITE)


def style_body_cell(cell, text, size=10, align=PP_ALIGN.LEFT,
                    is_number=False, bold=False, highlight=False, alt_row=False):
    """
    正文单元格样式。

    Args:
        is_number : True → Arial 字体（数字/英文内容）
        bold      : True → 加粗突出
        highlight : True → 强制 PALE_GRAY 背景（手动突出某行/某格）
        alt_row   : True → 交替行 PALE_GRAY 背景，False → 白色背景
    """
    _set_cell_fill(cell, PALE_GRAY if (highlight or alt_row) else WHITE)
    _set_cell_valign(cell, 'ctr')
    add_text(cell.text_frame, text, first=True, align=align,
             cn_font=CN_FONT if not is_number else None,
             en_font=EN_FONT if is_number else None,
             size=size, bold=bold, color=DARK_GRAY)


def add_vi_table(slide, headers, rows_data,
                 left=Inches(0.40), top=Inches(1.42),
                 width=Inches(9.20), height=None,
                 header_size=11, body_size=10,
                 col_widths=None):
    """
    在 slide 上添加符合 VI 规范的表格。

    Args:
        slide       : 幻灯片对象
        headers     : 表头列表，如 ['基金名称', '规模(亿)', '成立日期']
        rows_data   : 二维列表，每行为若干单元格值。
                      单元格可以是：
                        - 字符串：默认样式（中文字体，左对齐，DARK_GRAY）
                        - (text, dict)：dict 支持 is_number / bold / highlight / align
        left/top    : 表格左上角坐标（默认贴内容区边界）
        width       : 表格总宽度（默认填满内容区）
        height      : 表格总高度（None → 按行数自动估算 0.38"/行）
        header_size : 表头字号，默认 11pt
        body_size   : 正文字号，默认 10pt
        col_widths  : 各列宽度权重列表，如 [3, 1, 1]（不传则均分）

    Returns:
        table 对象

    Example:
        add_vi_table(slide,
            headers=['基金公司', 'AUM（亿元）', '成立年份', '近三年收益'],
            rows_data=[
                ['易方达基金', ('21,300', dict(is_number=True)),
                 ('2001', dict(is_number=True)),
                 ('+38.2%', dict(is_number=True, bold=True))],
                ['华夏基金',   ('18,500', dict(is_number=True)),
                 ('1998', dict(is_number=True)),
                 ('+31.6%', dict(is_number=True))],
            ],
            col_widths=[3, 2, 1.5, 2],
        )
    """
    n_cols = len(headers)
    n_rows = len(rows_data) + 1   # +1 表头行
    auto_h = Inches(0.38) * n_rows if height is None else height
    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, auto_h).table

    # 列宽分配
    if col_widths:
        total_w = sum(col_widths)
        for ci, w in enumerate(col_widths):
            tbl.columns[ci].width = int(width * w / total_w)
    else:
        unit_w = width // n_cols
        for ci in range(n_cols):
            tbl.columns[ci].width = unit_w

    # 行高均分
    row_h = auto_h // n_rows
    for ri in range(n_rows):
        tbl.rows[ri].height = row_h

    # 表头行
    for ci, hdr in enumerate(headers):
        style_header_cell(tbl.cell(0, ci), hdr, size=header_size)

    # 正文行（奇数行 PALE_GRAY，偶数行白色）
    for ri, row in enumerate(rows_data):
        alt = (ri % 2 == 1)
        for ci, cell_val in enumerate(row):
            if isinstance(cell_val, tuple):
                text, kwargs = cell_val
            else:
                text, kwargs = str(cell_val), {}
            style_body_cell(tbl.cell(ri + 1, ci), text,
                            size=body_size, alt_row=alt, **kwargs)

    return tbl
```

### 正文页（Layout 2 纯文字）

```python
slide = prs.slides.add_slide(prs.slide_layouts[2])

set_slide_title(slide, '投资理念与策略框架')

# 多节内容：使用 add_body_content_blocks（每节独立 text box，节间绝对空白）
add_body_content_blocks(slide, [
    ('核心投资理念', 0),
    ('客户至上：为客户提供有针对性的产品与专业解决方案', 1),
    ('价值导向：以基本面研究驱动投资决策', 1),
    ('风险管理', 0),
    ('全流程风险管理，覆盖投前、投中、投后各环节', 1),
    ('独立的风控团队，与投资团队形成制衡机制', 1),
])
# 单节/无 L0 内容：仍可用 add_body_content(shape.text_frame, items)
```

### 正文页（Layout 3 左文右图）

```python
# ── 坐标常量 ──────────────────────────────────────────────────
# 间距规则：label(0.27") → 0.10" gap → chart → 0.10" gap → caption(0.22")
# label y=1.20"→1.47", chart y=1.57"→3.83"(h=2.26"), caption y=3.93"
# ⚠️ chart 必须在 label/caption 之前添加（先加的在下层），label/caption 才不会被图片遮盖
_L3_RIGHT_LABEL = (Inches(5.40), Inches(1.20), Inches(4.34), Inches(0.27))
_L3_RIGHT_CHART = (Inches(5.40), Inches(1.57), Inches(4.34), Inches(2.26))
_L3_RIGHT_CAP   = (Inches(5.40), Inches(3.93), Inches(4.34), Inches(0.22))


def add_layout3_slide(prs, title, body_items,
                      right_label='', right_caption='', right_image=None):
    """
    添加 Layout 3（左文右图）幻灯片，自动填写左侧文字，返回右侧图表区坐标。

    Args:
        prs           : Presentation 对象
        title         : 幻灯片标题
        body_items    : 左侧文字内容，格式同 add_body_content 的 items 参数
                        如 [('规模概况', 0), ('突破2万亿元', 1), ...]
        right_label   : 右侧图表小标题（图表上方，10pt DEEP_BLUE，可省略）
        right_caption : 右侧数据来源（图表下方，7pt DARK_GRAY，可省略）

    Returns:
        slide            : 幻灯片对象
        _L3_RIGHT_CHART  : 右图区域 (l, t, w, h)，供 add_chart / add_picture 使用

    Example:
        from pptx.enum.chart import XL_CHART_TYPE
        from pptx.chart.data import CategoryChartData

        slide, chart_area = add_layout3_slide(
            prs, '资产管理规模增长',
            body_items=[
                ('规模概况', 0),
                ('截至2024年底，管理规模突破2万亿元', 1),
                ('公募基金规模行业前三', 1),
                ('主要产品线', 0),
                ('权益类基金：占比约40%', 1),
                ('固收类基金：占比约45%', 1),
            ],
            right_label='AUM增长趋势（亿元）',
            right_caption='数据来源：公司季报，截至2024年12月',
        )

        cd = CategoryChartData()
        cd.categories = ['2021', '2022', '2023', '2024']
        cd.add_series('管理规模', (15800, 17200, 18900, 21300))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, *chart_area, cd).chart
        chart.has_title = False
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = DEEP_BLUE
    """
    slide = prs.slides.add_slide(prs.slide_layouts[3])
    set_slide_title(slide, title)

    # 左侧文字区：多节内容用 add_body_content_blocks（节间绝对空白，更可靠）
    add_body_content_blocks(slide, body_items,
                            left=_PH10_LEFT, top=_PH10_TOP,
                            width=_PH10_WIDTH[3])

    # 图片先加（在下层），label/caption 后加（在上层，不被图片遮挡）
    if right_image and os.path.exists(right_image):
        l, t, w, h = _L3_RIGHT_CHART
        add_picture_fit(slide, right_image, l, t, w, h)
    if right_label:   _add_textbox(slide, _L3_RIGHT_LABEL, right_label,   10, color=DEEP_BLUE)
    if right_caption: _add_textbox(slide, _L3_RIGHT_CAP,   right_caption,  7, color=DARK_GRAY)

    return slide, _L3_RIGHT_CHART
```

> **不要在文本字符串中手动嵌入 • 字符**；`add_body_content` 通过 `_set_para_bullet` 在 XML 层面添加项目符号（Arial •，悬挂缩进 4.8mm），level 0 标题显式设为 `buNone`。

### 正文页（上文下表 / 上文下图）

当一页中**文字在上、表格/图表在下**时，需缩短 ph10 高度为文字腾出空间。

> ⚠️ **必须用 `shrink_ph10()` 缩短 ph10，不可直接赋值 `shape.height`。**
> 直接赋值会让 python-pptx 新建 `<a:xfrm>` 元素，但 `off/@x` 和 `off/@y` 默认为 0，
> 导致文字出现在标题位置（top=0）或文字竖排（width=0）。

```python
# ── 常量（来自模板实测值） ────────────────────────────────────
_PH10_TOP  = Inches(1.42)   # 所有正文布局内容区顶部
_PH10_LEFT = Inches(0.40)   # 所有正文布局内容区左边距

_PH10_WIDTH = {
    2: Inches(8.782),   # Layout 2 全宽
    3: Inches(4.884),   # Layout 3 左半区
}


def shrink_ph10(slide, layout_idx, new_height_inches):
    """
    缩短 ph10（内容占位符）高度，为下方表格/图表腾出空间。

    ⚠️ 必须同时显式设置 left / top / width / height 全部四个属性，
    否则修改任意一个都会创建新的 <a:xfrm>，而其余三个属性默认为 0，
    导致文字移位（top=0 → 出现在标题位置）或竖排（width=0）。

    Args:
        slide             : 幻灯片对象
        layout_idx        : 布局索引（2 或 3）
        new_height_inches : 文字区新高度（英寸），例如 0.9 或 1.2

    Returns:
        找到的 ph10 shape，或 None。
    """
    for shape in slide.shapes:
        try:
            if shape.placeholder_format.idx == 10:
                shape.left   = _PH10_LEFT
                shape.top    = _PH10_TOP
                shape.height = Inches(new_height_inches)
                shape.width  = _PH10_WIDTH[layout_idx]
                return shape
        except:
            pass
    return None


# ── 示例：Layout 2 上文下表 ──────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[2])
set_slide_title(slide, '主要指标对比')

# 1. 缩短 ph10，留出 1.2" 给文字
ph10 = shrink_ph10(slide, layout_idx=2, new_height_inches=1.2)
if ph10:
    add_body_content(ph10.text_frame, [
        ('核心结论', 0),
        ('2024年旗舰主动权益基金平均超额收益 +8.2%', 1),
        ('固收类产品最大回撤控制在 0.5% 以内', 1),
    ], available_pt=1.2 * 72)   # ← 与 new_height_inches 对应

# 2. 表格放在文字区下方
# ph10 占 y=1.42"～(1.42+1.2)=2.62"，表格从 y=2.65" 开始
TABLE_TOP = Inches(1.42 + 1.2 + 0.05)   # 5px 间隙
add_vi_table(slide,
    headers=['指标', '2023年', '2024年', '同比变化'],
    rows_data=[
        ['管理规模(亿)',  ('18,900', dict(is_number=True)),
                         ('21,300', dict(is_number=True)),
                         ('+12.7%', dict(is_number=True, bold=True))],
        ['超额收益',     ('+5.1%',  dict(is_number=True)),
                         ('+8.2%',  dict(is_number=True)),
                         ('+3.1pp', dict(is_number=True, bold=True))],
        ['最大回撤',     ('-0.8%',  dict(is_number=True)),
                         ('-0.5%',  dict(is_number=True)),
                         ('+0.3pp', dict(is_number=True))],
    ],
    top=TABLE_TOP,
    height=Inches(3.86 - 1.42 - 1.2 - 0.05),  # 填满内容区剩余空间
)
```

### 图表样式工具函数

```python
def style_chart(chart,
                series_colors=None,
                gridlines='h',
                show_legend=True,
                legend_size=9):
    """
    为图表应用 VI 规范样式。add_chart 之后立即调用。

    VI 规范强制项（自动应用）：
      - 图表区 / 绘图区：透明背景，无边框（noFill）
      - 无图表标题（chart.has_title = False）

    Args:
        chart         : python-pptx Chart 对象
        series_colors : 各系列颜色列表（按顺序），如 [DEEP_BLUE, LIGHT_GRAY]
                        柱/面积图 → 设置 fill；折线图 → 设置 line.color
                        None → 不修改颜色
        gridlines     : 'h'  → 横网格线（value 轴，默认）
                        'v'  → 纵网格线（category 轴）
                        'none' → 无网格线
                        颜色固定 LIGHT_GRAY(R204,204,204) 实线，宽度 0.5pt
        show_legend   : True → 保留图例（默认）；False → 隐藏图例
        legend_size   : 图例字号，默认 9pt，字体 Arial

    常用系列颜色组合：
        单系列    → [DEEP_BLUE]
        双系列    → [DEEP_BLUE, LIGHT_GRAY]
        三系列    → [DEEP_BLUE, BRIGHT_BLUE, LIGHT_GRAY]
        突出+淡化 → [DEEP_BLUE, LIGHT_GRAY, LIGHT_GRAY, LIGHT_GRAY]

    Example:
        cd = CategoryChartData()
        cd.categories = ['2021', '2022', '2023', '2024']
        cd.add_series('易方达', (12.3, 8.7, 15.2, 18.6))
        cd.add_series('沪深300', (5.1, -21.6, -11.4, 14.7))
        chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE, *chart_area, cd).chart
        style_chart(chart,
                    series_colors=[DEEP_BLUE, LIGHT_GRAY],
                    gridlines='h',
                    show_legend=True)
    """
    # ── 1. 无图表标题 ────────────────────────────────────────
    chart.has_title = False

    # ── 2. 背景透明（XML 操作，chart_area / plot_area 无高层 API）──
    chartSpace = chart._element                          # <c:chartSpace>
    c_chart    = chartSpace.find(qn('c:chart'))
    plotArea   = c_chart.find(qn('c:plotArea'))

    for node in (chartSpace, plotArea):
        spPr = node.find(qn('c:spPr'))
        if spPr is None:
            spPr = etree.SubElement(node, qn('c:spPr'))
        for tag in [qn('a:solidFill'), qn('a:gradFill'),
                    qn('a:noFill'),    qn('a:pattFill')]:
            el = spPr.find(tag)
            if el is not None:
                spPr.remove(el)
        etree.SubElement(spPr, qn('a:noFill'))
        # 无边框
        ln = spPr.find(qn('a:ln'))
        if ln is None:
            ln = etree.SubElement(spPr, qn('a:ln'))
        noFillLn = ln.find(qn('a:noFill'))
        if noFillLn is None:
            ln.clear(); etree.SubElement(ln, qn('a:noFill'))

    # ── 3. 系列颜色 ──────────────────────────────────────────
    if series_colors:
        for i, color in enumerate(series_colors):
            if i >= len(chart.series): break
            s = chart.series[i]
            try:
                # 折线图：设置线条颜色
                s.format.line.color.rgb = color
            except Exception:
                pass
            try:
                # 柱/面积图：设置填充颜色
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = color
            except Exception:
                pass

    # ── 4. 网格线 ────────────────────────────────────────────
    _GRAY_EMU = str(int(0.5 * 12700))   # 0.5pt in EMU

    def _set_gridline_color(axis):
        """为指定轴的主网格线设置 LIGHT_GRAY 实线 0.5pt。"""
        gl = axis.major_gridlines
        ln = gl.format.line
        ln.color.rgb = LIGHT_GRAY
        # 线宽：直接操作 XML
        ln_el = gl.format._element.find(qn('a:ln'))
        if ln_el is not None:
            ln_el.set('w', _GRAY_EMU)

    def _hide_gridlines(axis):
        axis.has_major_gridlines = False

    # 默认先全部关掉，再按参数开一侧
    try: _hide_gridlines(chart.value_axis)
    except Exception: pass
    try: _hide_gridlines(chart.category_axis)
    except Exception: pass

    if gridlines == 'h':
        try:
            chart.value_axis.has_major_gridlines = True
            _set_gridline_color(chart.value_axis)
        except Exception:
            pass
    elif gridlines == 'v':
        try:
            chart.category_axis.has_major_gridlines = True
            _set_gridline_color(chart.category_axis)
        except Exception:
            pass
    # gridlines == 'none'：保持关闭

    # ── 5. 图例 ──────────────────────────────────────────────
    chart.has_legend = show_legend
    if show_legend:
        legend = c_chart.find(qn('c:legend'))
        if legend is not None:
            txPr = legend.find(qn('c:txPr'))
            if txPr is None:
                txPr = etree.SubElement(legend, qn('c:txPr'))
            txPr.clear()
            etree.SubElement(txPr, qn('a:bodyPr'))
            etree.SubElement(txPr, qn('a:lstStyle'))
            p      = etree.SubElement(txPr, qn('a:p'))
            pPr    = etree.SubElement(p, qn('a:pPr'))
            defRPr = etree.SubElement(pPr, qn('a:defRPr'))
            defRPr.set('sz', str(legend_size * 100))   # 单位：百分之一磅
            defRPr.set('b', '0')
            latin  = etree.SubElement(defRPr, qn('a:latin'))
            latin.set('typeface', EN_FONT)
            etree.SubElement(p, qn('a:endParaRPr')).set('lang', 'zh-CN')
```

### 正文页（Layout 4 双图）

Layout 4 不使用 ph idx=10，两个图表区域通过坐标手动放置。

```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ── 坐标常量 ──────────────────────────────────────────────────
# 间距规则：label(0.27") → 0.10" gap → chart → 0.10" gap → caption(0.22")
# label y=1.32"→1.59", chart y=1.69"→3.83"(h=2.14"), caption y=3.93"
# ⚠️ chart 必须在 label/caption 之前添加（先加的在下层），label/caption 才不会被图片遮盖
_L4_LEFT_LABEL  = (Inches(0.48), Inches(1.32), Inches(3.29), Inches(0.27))
_L4_RIGHT_LABEL = (Inches(5.08), Inches(1.32), Inches(3.29), Inches(0.27))
_L4_LEFT_CHART  = (Inches(0.40), Inches(1.69), Inches(4.32), Inches(2.14))
_L4_RIGHT_CHART = (Inches(4.71), Inches(1.69), Inches(4.34), Inches(2.14))
_L4_LEFT_CAP    = (Inches(0.40), Inches(3.93), Inches(4.05), Inches(0.22))
_L4_RIGHT_CAP   = (Inches(5.57), Inches(3.93), Inches(4.05), Inches(0.22))


from PIL import Image as _PILImage

def add_picture_fit(slide, image_path, box_l, box_t, box_w, box_h):
    """
    按原始宽高比缩放图片以适应 box，居中放置，不拉伸。

    ⚠️ 不要直接用 slide.shapes.add_picture(path, l, t, w, h) 同时指定宽高，
    否则图片会被强制拉伸填满 box，导致文字变形。
    改用此函数，以 box 为边界等比缩放并居中。

    Args:
        slide              : 幻灯片对象
        image_path         : 图片路径
        box_l/t/w/h        : 目标区域坐标和尺寸（EMU），通常直接传入 _L3_RIGHT_CHART 等常量
    """
    with _PILImage.open(image_path) as im:
        iw, ih = im.size
    if iw / ih > box_w / box_h:   # 图片更宽，以宽度为基准缩放
        new_w = box_w
        new_h = int(box_w * ih / iw)
    else:                          # 图片更高，以高度为基准缩放
        new_h = box_h
        new_w = int(box_h * iw / ih)
    offset_l = (box_w - new_w) // 2
    offset_t = (box_h - new_h) // 2
    slide.shapes.add_picture(image_path, box_l + offset_l, box_t + offset_t, new_w, new_h)


def _add_textbox(slide, ltwh, text, size, bold=False, color=DARK_GRAY, cn_font=CN_FONT):
    l, t, w, h = ltwh
    txb = slide.shapes.add_textbox(l, t, w, h)
    add_text(txb.text_frame, text, first=True,
             cn_font=cn_font, size=size, bold=bold, color=color)
    return txb


def add_layout4_slide(prs, title,
                      left_label='', right_label='',
                      left_caption='', right_caption='',
                      left_image=None, right_image=None):
    """
    添加 Layout 4（双图）幻灯片，返回 slide 和两个图表区域坐标。

    Args:
        prs           : Presentation 对象
        title         : 幻灯片标题
        left_label    : 左图小标题（图表上方，10pt DEEP_BLUE）
        right_label   : 右图小标题
        left_caption  : 左图数据来源（图表下方，7pt MID_GRAY）
        right_caption : 右图数据来源

    Returns:
        slide         : 幻灯片对象
        _L4_LEFT_CHART  : 左图区域 (l, t, w, h)，供 add_chart / add_picture 使用
        _L4_RIGHT_CHART : 右图区域 (l, t, w, h)

    Example:
        from pptx.enum.chart import XL_CHART_TYPE
        from pptx.chart.data import CategoryChartData

        slide, l_area, r_area = add_layout4_slide(
            prs, '规模与收益对比',
            left_label='AUM增长趋势（亿元）',
            right_label='年化收益率对比（%）',
            left_caption='数据来源：公司季报',
            right_caption='数据来源：Wind',
        )

        cd = CategoryChartData()
        cd.categories = ['2021', '2022', '2023', '2024']
        cd.add_series('规模', (15800, 17200, 18900, 21300))
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, *l_area, cd).chart
        chart.has_title = False
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = DEEP_BLUE
    """
    slide = prs.slides.add_slide(prs.slide_layouts[4])
    set_slide_title(slide, title)

    # 图片先加（在下层），label/caption 后加（在上层，不被图片遮挡）
    if left_image and os.path.exists(left_image):
        l, t, w, h = _L4_LEFT_CHART
        add_picture_fit(slide, left_image, l, t, w, h)
    if right_image and os.path.exists(right_image):
        l, t, w, h = _L4_RIGHT_CHART
        add_picture_fit(slide, right_image, l, t, w, h)
    if left_label:    _add_textbox(slide, _L4_LEFT_LABEL,  left_label,  10, color=DEEP_BLUE)
    if right_label:   _add_textbox(slide, _L4_RIGHT_LABEL, right_label, 10, color=DEEP_BLUE)
    if left_caption:  _add_textbox(slide, _L4_LEFT_CAP,    left_caption,  7, color=DARK_GRAY)
    if right_caption: _add_textbox(slide, _L4_RIGHT_CAP,   right_caption, 7, color=DARK_GRAY)

    return slide, _L4_LEFT_CHART, _L4_RIGHT_CHART
```

### 目录页

> **⚠️ 目录表格不在 layout 中，需从模板幻灯片克隆。**
> 必须在删除所有 slides 之前先提取 `toc_table_xml`（见 Quick Start）。
> 表格结构：7行×2列，col0=序号（Arial 28pt 亮蓝），col1=标题（华文黑体_易方达 18pt）。

```python
from pptx.oxml.ns import qn
from lxml import etree
import copy


def _set_toc_row(tr, number: str, title: str, color_hex: str):
    """更新一行目录：序号文字 + 标题文字 + 序号/标题颜色（操作原始 XML tr 元素）。

    color_hex 同时控制序号和标题的颜色：
      '005096' (DEEP_BLUE)  → 标题深蓝，序号亮蓝（1EB9E1）
      'CCCCCC' (LIGHT_GRAY) → 标题和序号均灰化
    """
    cells = tr.findall(qn('a:tc'))

    def _apply_color(rEl, hex_val):
        """向 run 元素写入 solidFill 颜色（insert(0) 保证在 latin/ea/cs 之前）。"""
        rPr = rEl.find(qn('a:rPr'))
        if rPr is None:
            rPr = etree.SubElement(rEl, qn('a:rPr'))
        old = rPr.find(qn('a:solidFill'))
        if old is not None:
            rPr.remove(old)
        fill = etree.Element(qn('a:solidFill'))
        etree.SubElement(fill, qn('a:srgbClr')).set('val', hex_val)
        rPr.insert(0, fill)

    # 序号颜色：灰化时同步灰化，激活/总览时保持亮蓝（BRIGHT_BLUE = 1EB9E1）
    num_color = color_hex if color_hex == 'CCCCCC' else '1EB9E1'

    # Col 0：序号文字 + 颜色
    for rEl in cells[0].findall('.//' + qn('a:r')):
        t = rEl.find(qn('a:t'))
        if t is not None:
            t.text = number
        _apply_color(rEl, num_color)

    # Col 1：标题文字 + 颜色
    for rEl in cells[1].findall('.//' + qn('a:r')):
        t = rEl.find(qn('a:t'))
        if t is not None:
            t.text = title
        _apply_color(rEl, color_hex)


def fill_toc_table(tbl, chapters: list[str], active_idx):
    """
    填写目录表格内容，自动调整行数，高亮当前章节，灰化其余章节。

    Args:
        tbl        : shape.table 对象
        chapters   : 章节标题列表（任意长度，无上限）
        active_idx : 当前章节的 0-based 索引（显示为 DEEP_BLUE，其余灰化）
                     None → 全部章节显示为 DEEP_BLUE（总览目录页专用）

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
        if active_idx is None:
            color_hex = '005096'   # 全部 DEEP_BLUE（总览页）
        else:
            color_hex = '005096' if ri == active_idx else 'CCCCCC'
        _set_toc_row(tr, f'{ri + 1:02d}.', chapters[ri], color_hex)


def add_toc_slide(prs, toc_table_xml, chapters: list[str], active_idx=None):
    """
    添加目录页：克隆模板表格 XML → 填入内容 → 返回 slide。

    Args:
        prs           : Presentation 对象
        toc_table_xml : Quick Start 阶段提取的 deepcopy XML 元素
        chapters      : 章节标题列表（最多 7 项）
        active_idx    : 当前高亮章节的 0-based 索引；None → 全部章节蓝色（总览页）
    """
    if toc_table_xml is None:
        raise ValueError(
            "toc_table_xml 为 None —— 必须在删除 slides 之前提取，见 Quick Start 提取步骤。"
            "常见错误：把 'while len(prs.slides) > 0' 的清空操作放在了提取之前。"
        )
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

# ① 演示文稿开头：总览目录页，全部章节蓝色
add_toc_slide(prs, toc_table_xml, chapters)           # active_idx=None（默认）

# ② 每章正文前：章节跳转页 + 本章正文，必须穿插，不可批量
# ⚠️ 错误示范：先生成所有跳转页，再生成所有正文 —— 会导致目录页扎堆
# ✅ 正确示范：每章跳转页紧跟本章正文，循环内同时生成
for i, chapter in enumerate(chapters):
    add_toc_slide(prs, toc_table_xml, chapters, active_idx=i)  # 章节跳转页
    # 紧接着生成本章的正文页，不要移到循环外
    add_layout3_slide(prs, title="...", body_items=[...])       # 本章第一页
    # add_layout3_slide(prs, ...)                               # 本章更多页 ...
```

### 封底页

```python
def add_cn_closing_slide(prs, thanks='谢谢', contacts=None):
    """
    添加中文封底（Layout 5）。

    Args:
        prs      : Presentation 对象
        thanks   : 答谢词，华文黑体_易方达 45pt 加粗 DEEP_BLUE
        contacts : 联系信息行列表（可选），华文黑体_易方达 12pt DEEP_BLUE 行距 170%
                   如：['联系我们：', '张三   Tel：+86(20)8510-xxxx', 'Email：xxx@efunds.com.cn']

    Returns:
        slide
    """
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        tf  = ph.text_frame
        tf.word_wrap = True

        if idx == 10:
            para = add_text(tf, thanks, first=True,
                            cn_font=CN_FONT, en_font=None,
                            size=45, bold=True, color=DEEP_BLUE)
            _set_para_spacing(para, line_spc_pct=100)

        elif idx == 11 and contacts:
            for i, line in enumerate(contacts):
                para = add_text(tf, line, first=(i == 0),
                                cn_font=CN_FONT, en_font=None,
                                size=12, bold=False, color=DEEP_BLUE)
                _set_para_spacing(para, line_spc_pct=170)

    return slide


def add_en_closing_slide(prs, thanks='Thank You', contacts=None):
    """
    添加英文封底（Layout 6）。

    Args:
        prs      : Presentation 对象
        thanks   : 答谢词，Arial 45pt 加粗 DEEP_BLUE
        contacts : 联系信息行列表（可选），Arial 12pt DEEP_BLUE 行距 170%

    Returns:
        slide
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        tf  = ph.text_frame
        tf.word_wrap = True

        if idx == 10:
            para = add_text(tf, thanks, first=True,
                            cn_font=None, en_font=EN_FONT,
                            size=45, bold=True, color=DEEP_BLUE)
            _set_para_spacing(para, line_spc_pct=100)

        elif idx == 4294967295 and contacts:
            for i, line in enumerate(contacts):
                para = add_text(tf, line, first=(i == 0),
                                cn_font=None, en_font=EN_FONT,
                                size=12, bold=False, color=DEEP_BLUE)
                _set_para_spacing(para, line_spc_pct=170)

    return slide


# ── 示例调用 ────────────────────────────────────────────────
# 中文封底
add_cn_closing_slide(prs,
    thanks='谢谢',
    contacts=[
        '联系我们：',
        '张三   Tel：+86(20)8510-xxxx',
        'Email：investor@efunds.com.cn',
    ]
)

# 英文封底
add_en_closing_slide(prs,
    thanks='Thank You',
    contacts=[
        'Contact us：',
        'John Smith   Tel：+86(20)8510-xxxx',
        'Email：investor@efunds.com.cn',
    ]
)
```

**封底规范：**
- 答谢词可根据实际需要修改，不得添加任何图片
- 联系信息非必须，可省略 `contacts` 参数
- 文字、颜色、位置不得擅自变动

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
import math
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

CN_FONT = "华文黑体_易方达"   # ⚠️ 公司定制字体，必须使用完整名称
EN_FONT = "Arial"


def apply_font(run, cn_font=CN_FONT, en_font=EN_FONT,
               size=None, bold=None, color=None,
               italic=None, underline=None):
    """
    设置中文/英文/混排字体。

    ⚠️ 关键规则：中文字符必须通过 a:latin charset="-122"（GBK 中文字符集）设置，
    这是模板 slide 中实际使用的方式。a:ea 字段不被 PowerPoint 识别为中文字体覆盖。

    参数说明：
        cn_font=CN_FONT, en_font=None  → 纯中文：a:latin charset=-122 + a:cs
        cn_font=None, en_font=EN_FONT  → 纯英文/数字：a:latin（无 charset）
        cn_font=CN_FONT, en_font=EN_FONT → 中英混排：a:latin=Arial + a:ea=华文黑体_易方达
        italic    : True/False，斜体（文字突出用）
        underline : True/False，下划线（文字突出用）

    用法：
        run = para.add_run()
        run.text = "易方达 2024"
        apply_font(run, size=28, bold=False, color=DEEP_BLUE)

        # 突出显示某段文字
        apply_font(run, bold=True, color=BRIGHT_BLUE)
        apply_font(run, italic=True)
        apply_font(run, underline=True)
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

    if size      is not None: run.font.size      = Pt(size)
    if bold      is not None: run.font.bold      = bold
    if italic    is not None: run.font.italic    = italic
    if underline is not None: run.font.underline = underline
    if color     is not None: run.font.color.rgb = color


def add_text(tf, text, *, first=False, align=PP_ALIGN.LEFT,
             cn_font=CN_FONT, en_font=EN_FONT,
             size=None, bold=None, color=DEEP_BLUE,
             italic=None, underline=None):
    """
    向 text_frame 追加一个段落并应用字体。

    Args:
        tf        : shape.text_frame
        text      : 段落文字
        first     : True → 复用 tf.paragraphs[0]（清空已有内容），
                    False → 新增段落（默认）
        align     : 对齐方式，VI 规范默认靠左
        color     : 默认 DEEP_BLUE（R0,G80,B150），可覆盖
        italic    : True → 斜体（正文突出方式之一）
        underline : True → 下划线（正文突出方式之一）

    文字突出方式（规范允许）：
        bold=True               加粗
        size=原字号+1 或 +2     放大字号
        underline=True          下划线
        italic=True             斜体
        color=BRIGHT_BLUE 等    使用 VI 颜色

    Returns:
        para  : 设置完成的 paragraph 对象
    """
    para = tf.paragraphs[0] if first else tf.add_paragraph()
    if first:
        para.clear()
    run = para.add_run()
    run.text = text
    apply_font(run, cn_font=cn_font, en_font=en_font,
               size=size, bold=bold, color=color,
               italic=italic, underline=underline)
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
| **表头文字** | 华文黑体_易方达 加粗 | — | 11pt（可调） | WHITE |
| **表头背景** | — | — | — | BRIGHT_BLUE |
| **表格正文（中文）** | 华文黑体_易方达 | — | 10pt（可调） | DARK_GRAY |
| **表格正文（数字/英文）** | — | Arial | 10pt（可调） | DARK_GRAY |
| **表格交替行背景** | — | — | — | PALE_GRAY / WHITE |

---

## Key Rules (Summary)

**封面：**
- 非副标题文字必须显式传 `bold=False`，否则会继承 layout placeholder 的默认加粗样式
- 所有文字靠左对齐，只能在设定文本框内编辑
- 免责声明底色、Logo 底色不得修改；封面图片不得替换
- 合作方 Logo 放在标题与日期之间，第一个与标题左侧对齐，高度 ≤ 大标题字高

**目录：**
- 内容必须填入给定表格，不得添加图片，格式不能自行调整
- 当前章节正常显示，其他章节灰化（LIGHT_GRAY）
- **⚠️ 必须包含两种目录页，缺一不可：**
  1. **总览目录页**（演示文稿开头，紧跟封面）：`add_toc_slide(prs, toc_table_xml, chapters)` — `active_idx=None`，全部章节蓝色
  2. **章节跳转页**（每个章节正文前各一张）：`add_toc_slide(prs, toc_table_xml, chapters, active_idx=i)` — 当前章节蓝色，其余灰化
- 常见错误：只生成总览目录页，缺少章节跳转页 → 导致演示时无法指示当前进度

**正文：**
- "上文下图" 或 "左文右图"
- 正文页下方横线、公司名称、页码不得删除
- 不建议出现三级标题

**图表：**
- 纯白底色，无边框，无渐变/阴影/3D 效果
- 网格线二选一（横或纵）
- 模型图连接线 1.5 磅 MID_GRAY，内容框使用直角方框

**表格：**
- 使用 `add_vi_table(slide, headers, rows_data, ...)` 生成，自动处理颜色和字体
- 表头：BRIGHT_BLUE 背景 + 白色加粗华文黑体_易方达，居中对齐
- 正文：奇数行 PALE_GRAY / 偶数行白色交替，中部垂直对齐，内容左对齐
- 数字/英文列传 `is_number=True` 使用 Arial 字体
- 局部突出优先用 `bold=True`；`highlight=True` 会**强制覆盖**为 PALE_GRAY，
  会打破同行其他单元格的白底，通常不建议使用
- 禁止使用规定色系外颜色（如红色）

**免责声明：** 位于封面底部蓝色横幅，固定文本，不得删除、不得修改、底色不得修改。
- 中文/中英文封面 → `disclaimers/disclaimer_cn.txt`
- 英文封面 → `disclaimers/disclaimer_en.txt`

---

## QA Checklist

**假设输出有问题。你的工作是找出问题。**

第一次渲染几乎不会完全正确。把 QA 当作 bug 排查，而不是走形式的确认步骤。如果第一次检查没发现任何问题，说明你看得不够仔细。

### 第一步：内容检查

```bash
# 提取文字，检查内容完整性
python3 -c "
from pptx import Presentation
from markitdown import MarkItDown
md = MarkItDown()
print(md.convert('output.pptx').text_content)
"

# 检查残留占位符
python3 -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|单击此处|click.*edit"
```

若 grep 返回结果，修复后再继续。

### 第二步：转图检查

```bash
python3 scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
ls slide-*.jpg

# 快速缩略图网格（选用）
python3 scripts/thumbnail.py output.pptx
```

### 第三步：视觉检查（⚠️ 使用 Subagent，哪怕只有 2-3 张幻灯片）

**必须使用 subagent**——你已经盯着代码看了太久，会看到你期望的东西而不是实际存在的东西。Subagent 有全新的视角。

向 subagent 发送以下 prompt（替换实际图片路径）：

```
请仔细检查以下幻灯片图片，假设存在问题——找出所有问题。

检查要点（VI 规范专项）：
- 标题文字是否显示在蓝色横幅内（白色文字不应溢出到横幅外，也不应被压缩）
- 目录页：当前章节是否正确蓝色高亮，其余章节是否灰化（R204,G204,B204）
- 正文区文字是否超出内容区下边界（约 y=4.13"）
- 页码、横线、公司名称是否被内容遮挡（位于幻灯片底部，每页必须可见）
- 左右两侧留白是否保留（约 0.40" 边距，内容不得超出参考线）
- 封面免责声明是否存在于底部蓝色横幅中
- 正文多节内容：节标题（蓝色加粗）与正文条目是否视觉层级清晰
- 项目符号缩进是否一致（• 不应与文字重叠或错位）
- 节间空白是否合理（不过密也不过大）
- 文字框是否有截断（文字被框边界切割）
- 元素间是否存在重叠（图片遮盖文字，标题压住内容区等）

请逐张幻灯片列出发现的问题（哪怕是细微问题也要报告）：
1. /path/to/slide-01.jpg（预期内容：封面页）
2. /path/to/slide-02.jpg（预期内容：目录页）
...
```

### 第四步：验证循环

1. 生成幻灯片 → 转图 → 检查
2. **列出发现的问题**（若没发现任何问题，再仔细看一遍）
3. 修复问题
4. **重新检查受影响的幻灯片**——修复一个问题常常引发另一个问题
5. 重复直到全面检查后没有新问题

**在完成至少一次"修复并重新验证"循环之前，不得宣告任务完成。**
