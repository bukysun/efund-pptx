# 正文内容填充函数完整实现

> 本文件由 SKILL.md 引用。生成正文幻灯片时，先 Read 此文件，再将函数粘贴到脚本中。

## 依赖（已在 SKILL.md 主体中定义，无需重复导入）

```
math, etree, qn, add_text, _set_para_bullet, _set_para_spacing
BODY_STYLES, CONTENT_HEIGHT_PT, _BULLET_INDENT_EMU
_PH10_LEFT, _PH10_TOP, _PH10_WIDTH
```

## add_body_content — 单文本框多级正文

```python
def add_body_content(tf, items, available_pt=CONTENT_HEIGHT_PT,
                     available_width_pt=None):
    """
    填写多级正文，段前空间动态分配，自动估算换行、强化 L0/L1 视觉层级。

    分配逻辑：
      1. 用 available_width_pt 估算各条目实际行数（换行会增大 base_target，压缩间距）
      2. 节标题（level 0，非首条）：基础 _MIN_L0_SPC + 额外量（上限 24pt，总上限 30pt）
      3. 正文/二级条目：剩余空间均分（上限 14pt）
         → L0 间距始终约 2-3× L1 间距，保持清晰视觉层级
      4. 内容多时行距自动等比压缩至 100%（单倍）保底

    Args:
        tf                 : shape.text_frame（ph idx=10 的文本框）
        items              : list of (text, level) tuples
                             level 0 → 一级标题：华文黑体_易方达 加粗 15pt DEEP_BLUE
                             level 1 → 一级文字：华文黑体_易方达 12pt DARK_GRAY
                             level 2 → 二级文字：华文黑体_易方达 10pt DARK_GRAY
        available_pt       : 内容区高度（pt），默认 195pt（Layout 2/3 的 ph idx=10 高度）
        available_width_pt : 内容区宽度（pt），用于估算长文本换行数，影响间距分配精度。
                             通常不需手动传入；add_layout3_slide 会自动传入。
                             Layout 2 ≈ 633pt，Layout 3 ≈ 352pt，None → 按单行计算。

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

    # 先确定是否加 bullet（影响缩进宽度，进而影响换行估算）
    n_body_items = sum(1 for _, lv in items if lv != 0)
    use_bullet   = (n_body_items > 1)

    def _est_lines(text, level):
        """估算该段落在 available_width_pt 下的实际显示行数。"""
        if available_width_pt is None:
            return 1
        font_sz    = _FONT_SZ.get(level, 12)
        indent_pt  = (_BULLET_INDENT_EMU / 12700) if (use_bullet and level > 0) else 0
        usable     = max(font_sz, available_width_pt - indent_pt - 18)  # ~9pt 内边距/侧
        chars_line = max(1, usable / font_sz)
        return max(1, math.ceil(len(text) / chars_line))

    n_l0_gaps  = sum(1 for i, (_, lv) in enumerate(items) if lv == 0 and i > 0)
    avail_text = available_pt - n_l0_gaps * _MIN_L0_SPC

    # 考虑换行后的实际高度
    base_target = sum(_FONT_SZ.get(lv,12) * _TARGET_LNS.get(lv,160)/100
                      * _est_lines(t, lv) for t, lv in items)
    base_min    = sum(_FONT_SZ.get(lv,12) * _est_lines(t, lv) for t, lv in items)

    if base_target <= avail_text:
        scale     = 1.0
        remaining = avail_text - base_target
    elif base_min <= avail_text:
        scale     = avail_text / base_target
        remaining = 0.0
    else:
        scale     = 100 / max(_TARGET_LNS.values())
        remaining = 0.0

    actual_lns = {lv: max(100, int(_TARGET_LNS[lv] * scale)) for lv in [0, 1, 2]}

    # 两步间距分配：L0 权重 = 4× L1，保证节标题在视觉上明显分隔
    #   L0 额外间距上限 24pt（总上限 6+24=30pt）；L1 间距上限 14pt
    _MAX_L0_EXTRA = 28.0
    _MAX_OTHER    = 12.0
    n_other_gaps  = sum(1 for idx, (_, lv) in enumerate(items) if idx > 0 and lv != 0)

    if n_l0_gaps > 0 and remaining > 0:
        total_w  = 4 * n_l0_gaps + n_other_gaps
        unit_raw = remaining / total_w
        l0_extra = min(_MAX_L0_EXTRA, 4 * unit_raw)
    else:
        l0_extra = 0.0

    other_budget = remaining - n_l0_gaps * l0_extra
    unit_other   = (other_budget / n_other_gaps) if n_other_gaps > 0 else 0.0
    unit_other   = min(_MAX_OTHER, unit_other)

    for i, (text, level) in enumerate(items):
        para = add_text(tf, text, first=(i == 0),
                        **BODY_STYLES.get(level, BODY_STYLES[1]))
        # 项目符号：一级标题永远无符号；正文/二级文字仅在多段时加 •
        if level == 0 or not use_bullet:
            _set_para_bullet(para, enabled=False)
        else:
            _set_para_bullet(para, enabled=True, level=level)
        if i == 0:
            spc_before = 0.0
        elif level == 0:
            spc_before = _MIN_L0_SPC + l0_extra
        else:
            spc_before = unit_other
        _set_para_spacing(para, spc_before_pt=spc_before, line_spc_pct=actual_lns[level])
```

## add_body_content_blocks — 多文本框正文布局（推荐用于 2+ 个 L0 节）

每个 L0 节独立一个 text box，节间空白是两个 shape 之间的绝对空间，
任何渲染引擎都不会压缩，视觉效果比 spc_before 段落间距更可靠。

```python
_LINE_HT_PT = {0: 15 * 1.40, 1: 12 * 1.55, 2: 10 * 1.55}   # pt per line（含行距）


def add_body_content_blocks(slide, items,
                            left=_PH10_LEFT, top=_PH10_TOP,
                            width=_PH10_WIDTH[2],
                            available_pt=CONTENT_HEIGHT_PT,
                            min_gap_pt=14.0,
                            max_gap_pt=22.0):
    """
    多文本框正文布局：每个 L0 节独立一个 text box，节间留真实空白间距。

    相比 add_body_content（单文本框 + spc_before），节间空隙是两个 shape
    之间的绝对空白，渲染结果更稳定可靠。推荐用于含 2+ 个 L0 节的正文页。

    ⚠️ 调用此函数时会自动移除幻灯片中的 ph idx=10 占位框，防止空框显示。

    Args:
        slide        : 幻灯片对象（需要 slide，不是 tf）
        items        : list of (text, level) tuples，同 add_body_content
        left         : 内容区左边距（EMU），默认 _PH10_LEFT (0.40")
        top          : 内容区顶部（EMU），默认 _PH10_TOP (1.42")
        width        : 内容区宽度（EMU），默认 _PH10_WIDTH[2]（Layout 2 全宽）
        available_pt : 内容区可用高度（pt），默认 195pt
        min_gap_pt   : 节间最小间距（pt），默认 14pt
        max_gap_pt   : 节间最大间距（pt），默认 22pt（防止内容少时间距过大）

    Returns:
        list of text box shapes（每节一个）

    Example:
        add_body_content_blocks(slide, [
            ('核心投资理念', 0),
            ('客户至上：为客户提供有针对性的产品', 1),
            ('价值导向，研究驱动', 1),
            ('风险管理', 0),
            ('全流程风险管理体系', 1),
        ])
    """
    _FONT_SZ       = {0: 15, 1: 12, 2: 10}
    _LNS_TARGET    = {0: 140, 1: 155, 2: 155}   # 目标行距 %
    _ITEM_GAP_BASE = 3.0                          # 同节内相邻条目段前间距，pt

    # 1. 按 L0 分组
    sections = []
    current  = []
    for text, level in items:
        if level == 0 and current:
            sections.append(current)
            current = []
        current.append((text, level))
    if current:
        sections.append(current)
    if not sections:
        return []

    n_body     = sum(1 for _, lv in items if lv != 0)
    use_bullet = (n_body > 1)
    width_pt   = width / 12700   # EMU → pt
    n_gaps     = len(sections) - 1

    def _est_lines(text, level):
        font_sz   = _FONT_SZ.get(level, 12)
        indent_pt = (_BULLET_INDENT_EMU / 12700) if (use_bullet and level > 0) else 0
        usable    = max(font_sz, width_pt - indent_pt - 12)
        return max(1, math.ceil(len(text) / max(1, usable / font_sz)))

    def _section_h(sec_items, lns_pct, item_gap):
        h = 0.0
        for i, (txt, lv) in enumerate(sec_items):
            h += _FONT_SZ.get(lv, 12) * lns_pct.get(lv, 155) / 100 * _est_lines(txt, lv)
            if i > 0:
                h += item_gap
        return h

    lns_pct  = dict(_LNS_TARGET)
    item_gap = _ITEM_GAP_BASE
    heights  = [_section_h(sec, lns_pct, item_gap) for sec in sections]
    total_h  = sum(heights)

    # 2. 动态判断：多文本框 or 单文本框
    #    多文本框需要：内容高度 + n_gaps × min_gap_pt ≤ available_pt
    #    否则内容太密，gap 开销太大 → 退回单文本框（段落间距控制）
    total_with_min_gaps = total_h + (n_gaps * min_gap_pt if n_gaps > 0 else 0)

    if total_with_min_gaps > available_pt:
        # ── 单文本框模式 ──────────────────────────────────────
        # 找到 ph10 直接使用，找不到则新建一个等尺寸 textbox
        ph10_shape = None
        for shape in slide.shapes:
            try:
                if shape.placeholder_format.idx == 10:
                    ph10_shape = shape
                    break
            except Exception:
                pass
        if ph10_shape:
            add_body_content(ph10_shape.text_frame, items,
                             available_pt=available_pt,
                             available_width_pt=width_pt)
        else:
            txb = slide.shapes.add_textbox(left, top, width, int(available_pt * 12700))
            txb.text_frame.word_wrap = True
            add_body_content(txb.text_frame, items,
                             available_pt=available_pt,
                             available_width_pt=width_pt)
        return []

    # ── 多文本框模式 ──────────────────────────────────────────
    # 3. 移除 ph idx=10 占位框，防止空框渲染为"单击此处编辑母版标题样式"
    for shape in list(slide.shapes):
        try:
            if shape.placeholder_format.idx == 10:
                shape._element.getparent().remove(shape._element)
                break
        except Exception:
            pass

    # 4. 节间空隙：剩余空间均分，限制在 [min_gap_pt, max_gap_pt]
    spare  = max(0.0, available_pt - total_h)
    gap_pt = min(max_gap_pt, max(min_gap_pt, spare / n_gaps)) if n_gaps > 0 else 0.0

    # 5. 逐节创建 text box，y 坐标累加
    y_emu  = top
    shapes = []
    for si, sec_items in enumerate(sections):
        sec_h_emu = int(_section_h(sec_items, lns_pct, item_gap) * 12700 * 1.25)  # 25% buffer 防裁剪
        txb = slide.shapes.add_textbox(left, y_emu, width, sec_h_emu)
        txb.text_frame.word_wrap = True
        shapes.append(txb)

        for i, (text, level) in enumerate(sec_items):
            para = add_text(txb.text_frame, text, first=(i == 0),
                            **BODY_STYLES.get(level, BODY_STYLES[1]))
            if level == 0 or not use_bullet:
                _set_para_bullet(para, enabled=False)
            else:
                _set_para_bullet(para, enabled=True, level=level)
            spc = 0.0 if i == 0 else item_gap
            _set_para_spacing(para, spc_before_pt=spc,
                              line_spc_pct=lns_pct.get(level, 155))

        y_emu += sec_h_emu + int(gap_pt * 12700)

    return shapes
```
