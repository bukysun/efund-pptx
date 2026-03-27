# efund-pptx QA & Tooling Improvements Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 强化 QA 流程（subagent 视觉检查、验证循环、VI 专属 checklist），并修复 scripts 引用缺失问题，补全视觉工具链。

**Architecture:** 分四步：① 复制 soffice.py 修复引用错误 → ② 复制 thumbnail.py 补全工具 → ③ 重写 SKILL.md QA 章节 → ④ 提交推送。所有改动在当前工作目录完成（symlink 到 ~/.claude/skills/efund-pptx，改动立即生效）。

**Tech Stack:** python-pptx, LibreOffice (soffice), Poppler (pdftoppm), Pillow, defusedxml

---

## Task 1：创建 scripts/ 目录结构并添加 soffice.py

**文件：**
- 创建: `scripts/__init__.py`（空文件）
- 创建: `scripts/office/__init__.py`（空文件）
- 创建: `scripts/office/soffice.py`（从参考 pptx skill 复制）

**目的：** QA Checklist 中引用了 `python3 scripts/office/soffice.py`，但目录不存在，执行会报 `No such file`。

**Step 1: 创建目录结构**

```bash
mkdir -p scripts/office
touch scripts/__init__.py
touch scripts/office/__init__.py
```

**Step 2: 从参考 skill 复制 soffice.py**

```bash
cp /Users/wuhui/Documents/workspace/skills/anthropic/skills/skills/pptx/scripts/office/soffice.py \
   scripts/office/soffice.py
```

**Step 3: 验证文件存在且可执行**

```bash
python3 scripts/office/soffice.py --version 2>&1 | head -3
# 预期：输出 soffice 版本信息（或 "command not found" 如未安装 LibreOffice，但不应报 Python import 错误）
```

**Step 4: Commit**

```bash
git add scripts/
git commit -m "feat(scripts): add soffice.py helper to fix QA command reference"
```

---

## Task 2：添加 thumbnail.py 快速视觉预览工具

**文件：**
- 创建: `scripts/thumbnail.py`（从参考 pptx skill 复制）

**目的：** 补全快速视觉分析工具，让 AI 在生成幻灯片后可用一条命令生成缩略图网格，不需要走完整 soffice+pdftoppm 流程。

**Step 1: 从参考 skill 复制 thumbnail.py**

```bash
cp /Users/wuhui/Documents/workspace/skills/anthropic/skills/skills/pptx/scripts/thumbnail.py \
   scripts/thumbnail.py
```

**Step 2: 验证导入正常（依赖：defusedxml, Pillow）**

```bash
python3 -c "import scripts.thumbnail; print('OK')" 2>&1
# 若报 ModuleNotFoundError(defusedxml/PIL)：
# pip install defusedxml Pillow
```

**Step 3: 验证命令行可用**

```bash
python3 scripts/thumbnail.py --help
# 预期：打印用法说明，无报错
```

**Step 4: Commit**

```bash
git add scripts/thumbnail.py
git commit -m "feat(scripts): add thumbnail.py for quick slide visual overview"
```

---

## Task 3：重写 SKILL.md 的 QA Checklist 章节

**文件：**
- 修改: `SKILL.md`（仅 QA Checklist 章节）

**目的：** 现有 QA 只有两条命令，缺乏"挑剔找问题"的心态指引、subagent 视觉检查指导和 efund VI 专属 checklist。参考 pptx skill 的 QA 章节作为对标。

**需要替换的内容：** 将现有 `## QA Checklist` 章节整体替换为以下内容：

```markdown
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
```

**Step 1: 定位并替换 QA 章节**

打开 `SKILL.md`，找到 `## QA Checklist` 开始的位置（当前内容只有两段代码块），将其替换为上方完整内容。

**Step 2: 验证替换后 SKILL.md 格式正确**

```bash
python3 -c "
import re
with open('SKILL.md') as f:
    content = f.read()
assert '假设输出有问题' in content, 'QA mindset missing'
assert 'Subagent' in content, 'Subagent guidance missing'
assert '验证循环' in content, 'Verification loop missing'
assert '目录页' in content, 'efund-specific checklist missing'
print('OK - all required sections present')
"
```

**Step 3: Commit**

```bash
git add SKILL.md
git commit -m "feat(qa): rewrite QA checklist with subagent guidance, verification loop, and VI-specific checklist"
```

---

## Task 4：推送到远端

**Step 1: 推送所有提交**

```bash
git push origin master
```

**Step 2: 验证**

```bash
git log --oneline -5
# 应显示 Task 1、2、3 的三条提交
```

---

## 依赖说明

```bash
# QA 工具链依赖（首次使用前安装）
pip install "markitdown[pptx]"   # 文字提取
pip install defusedxml Pillow    # thumbnail.py 依赖
# LibreOffice：brew install libreoffice（macOS）
# Poppler：brew install poppler（macOS）
```
