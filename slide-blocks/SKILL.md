---
name: slide-blocks
description: PPT 智能组装助手。用于从素材库中挑选幻灯片、拼装成一份完整的、风格统一的 PPT。当用户说"帮我做一个PPT"、"组装一份汇报"、"从素材里选几页拼一下"、"我需要一个售前演示"，或提到具体章节主题和来源文件时，使用此 skill。也适用于需要预览素材库内容、查询有哪些可用片段、或对已有组装方案做调整的场景。
---

# SlideBlocks 组装助手

## 项目位置

- **引擎**：`D:/Claude/SlideBlocks/engine/assemble_template.py`
- **局部编辑**：`D:/Claude/SlideBlocks/engine/edit_pptx.py`
- **数据库**：`D:/Claude/SlideBlocks/slide_vault.db`
- **模板**：`D:/Claude/SlideBlocks/模板/` 下所有 `.pptx`（当前有「科技风（深色底）」和「蓝色商务（浅色底）」）
- **素材源**：`D:/Claude/SlideMatrix/素材/`
- **输出目录**：`D:/Claude/SlideBlocks/输出/`
- **任务脚本**：`D:/Claude/SlideBlocks/tasks/`（每次组装任务生成独立脚本，保留历史记录）

## 工作流程

### 第一步：收集需求

向用户确认以下信息（没有明确说明的才询问，已说明的直接跳过）：

1. **封面标题**：PPT 的主标题是什么？（可以帮用户起一个备选）
2. **章节结构**：要哪几个部分？每部分大概讲什么？
3. **素材来源**：指定来源文件，或让系统从库里检索
4. **模板风格**：深色底（默认）还是浅色底？

### 第二步：搜索素材库

在 `D:/Claude/SlideBlocks/` 目录下运行以下 Python 代码来检索可用素材：

```python
import sys
sys.path.insert(0, 'D:/Claude/SlideBlocks')
from slide_vault.search import search_content, search_structural, print_results

# 按内容检索（有标签的页面）
results = search_content(
    scene="售前汇报",          # 或 "行业会议"
    content_type="解决方案",    # 内容类型
    keywords=["AI", "智慧医疗"], # 关键词
    quality_min=4,              # 最低质量分
    source_file="北大",         # 来源文件关键词（可选）
    limit=10
)
print_results(results, mode="content")
```

**search_content 可用参数：**
- `scene`：场景，如 "售前汇报" / "行业会议"
- `content_type`：内容类型，如 "解决方案" / "行业背景" / "产品功能" / "标杆案例" / "实施保障"
- `keywords`：关键词列表，匹配幻灯片标题（优先）和 AI 标签，不查正文；支持同义词自动扩展（如"一体化"自动也搜"一体"）
- `quality_min`：最低质量分（1-5），建议传 4
- `source_file`：来源文件名关键词，如 "哈密" / "北大" / "福建"
- `limit`：返回条数上限

**返回字段**：`src`（文件路径）、`page`（页码）、`title`、`summary`、`quality_score`

将检索结果整理成表格展示给用户，让用户确认或调整选页。表格必须包含**来源文件**列，格式如下：

| 页 | 来源文件 | 内容 | 质量 |
|----|---------|------|------|
| P05 | 北大医疗交流 | 复星健康双SaaS云模式 | ★★★★★ |
| P28 | 哈密规划建议 | 复旦大学附属医院智能病历案例 | ★★★★★ |

来源文件名取文件名中最具辨识度的关键词（如"北大医疗"、"哈密"、"福建省立"），不需要写完整路径。

### 第三步：构建 PLAN

根据用户确认的页面，按以下格式构建 PLAN：

```python
PLAN = [
    # 封面（必须，模板P1）
    {"template_page": 1, "replace_title": "PPT主标题"},

    # 第一章节：开篇不加过渡页，直接上内容
    {"src": "D:/路径/文件名.pptx", "page": 5},
    {"src": "D:/路径/文件名.pptx", "page": 6},

    # 第二章节及之后：每章前加过渡页（模板P2）
    {"template_page": 2, "replace_title": "章节名称"},
    {"src": "D:/路径/文件名.pptx", "page": 12},
    {"src": "D:/路径/文件名.pptx", "page": 13},

    # 封底（必须，模板P5）
    {"template_page": 5},
]
```

**组装规则：**
- 封面后第一部分：直接接内容，**不加**过渡页
- 第二部分及之后：每章前加 `{"template_page": 2, "replace_title": "章节名"}`
- 封底：每份 PPT 最后必须加 `{"template_page": 5}`
- `replace_title` 仅覆盖标题文字，不影响模板页的选择（P3 vs P4 由引擎自动检测）

### 第四步：执行组装

在 `tasks/` 目录下写任务脚本并运行：

```python
# tasks/task_xxx.py（文件名用主题命名，如 task_卫宁AI场景实践.py）
# -*- coding: utf-8 -*-
import sys
from pathlib import Path

_ROOT = Path(__file__).parent.parent          # tasks/ 的上层 = SlideBlocks 根目录
sys.path.insert(0, str(_ROOT / "engine"))     # 找 assemble_template

import assemble_template

# 默认深色底；浅色底取消下面这行注释：
# assemble_template.TEMPLATE_PATH = _ROOT / "模板/蓝色商务（浅色底）.pptx"

PLAN = [
    # ... 在这里填入构建好的 PLAN
]

if __name__ == "__main__":
    assemble_template.assemble(PLAN, "输出文件名")
```

用 Bash 运行：
```bash
cd D:/Claude/SlideBlocks && python tasks/task_xxx.py
```

运行时会看到每页的处理进度，完成后输出文件路径。

### 第五步：告知用户结果

运行完成后，告知用户：
- 输出文件位置：`D:/Claude/SlideBlocks/输出/文件名.pptx`
- 共几页，结构是什么
- 如需小幅调整，使用 edit_pptx.py 直接编辑，无需重新组装（见下方说明）

---

## 局部编辑规范（重要）

**原则**：已组装好的 PPT，以下操作直接操作输出文件，**无需重新组装**：
- 删除某页
- 调换 / 移动页面顺序
- 插入模板页（过渡页 / 封底）
- 将某页替换为另一个源文件的内容

**只有**需要大幅重构章节结构时，才重新跑 task_current.py。

**工具**：`D:/Claude/SlideBlocks/engine/edit_pptx.py`

```python
import sys
sys.path.insert(0, 'D:/Claude/SlideBlocks/engine')
from edit_pptx import edit

output = "D:/Claude/SlideBlocks/输出/xxx.pptx"
edit(output, [
    # 删除页面
    {"op": "delete", "pages": [7, 8]},

    # 移动页面到指定位置之后
    {"op": "move", "pages": [7, 8], "after": 15},

    # 插入模板过渡页
    {"op": "insert_template", "template_page": 2, "after": 5, "title": "新章节"},

    # 将某页替换为另一个源文件的内容（保持模板背景）
    {"op": "replace", "page": 12, "src": "路径/文件.pptx", "src_page": 5},
])
```

写一个临时脚本运行：
```bash
cd D:/Claude/SlideBlocks && python -c "
import sys; sys.path.insert(0, 'engine')
from edit_pptx import edit
edit('输出/xxx.pptx', [
    {'op': 'delete', 'pages': 7},
])
"
```

---

## 模板页说明

| 页码 | 用途 | 何时使用 |
|------|------|---------|
| P1 | 封面 | 每份 PPT 开头，配合 `replace_title` 填标题 |
| P2 | 过渡页 | 第二章节及之后每章前，配合 `replace_title` 填章节名 |
| P3 | 带标题栏内容页 | 引擎自动选择（源页 top<65pt 有文字时） |
| P4 | 无标题栏内容页 | 引擎自动选择（源页 top<65pt 无文字时） |
| P5 | 封底 | 每份 PPT 结尾 |

P3/P4 由引擎自动检测，**不需要在 PLAN 里手动指定**。

---

## 已知素材文件（常用）

```
D:/Claude/SlideMatrix/素材/01-最佳实践/
├── 完整版-售前汇报-（北大医疗交流）卫宁健康智慧医疗系统设计与实践-深色底.250625.pptx
├── 完整版-售前汇报-福州大学附属省立医院一体化提升方案-深色底.250831.pptx
├── 完整版-售前汇报-哈密市医疗健康数智融合发展规划建议-浅色底.260202.pptx
└── ...（更多见数据库）
```

---

## 素材搜集稿（不套模板，原始格式汇总）

**场景**：把多份完整版 PPT 里的特定页（如"所有售前汇报的总体设计图"）汇总成一份参考文档，保留各页原始风格，**不套任何模板**。

这种场景**不使用** `assemble_template.py`，在 `tasks/` 下写任务脚本，用整页复制方式：

```python
# tasks/task_搜集稿_xxx.py
# -*- coding: utf-8 -*-
import sys, time
from pathlib import Path

_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(_ROOT / "engine"))
import assemble_template  # 复用 _get_ppt_app()，处理 WPS 劫持问题

PLAN = [
    {"src": "D:/Claude/SlideMatrix/素材/01-最佳实践/完整版-xxx.pptx", "page": 5, "label": "说明"},
    # ...更多页面
]
OUTPUT = str(_ROOT / "输出/搜集稿-xxx.pptx")

pptApp = assemble_template._get_ppt_app()   # ⚠️ 必须用此函数，不要直接 Dispatch
pptApp.Visible = True
pptApp.DisplayAlerts = 0

new_pres = pptApp.Presentations.Add(WithWindow=True)  # ⚠️ 含一个空白页，见下方陷阱说明

for i, item in enumerate(PLAN, 1):
    pptApp.DisplayAlerts = 0                           # ⚠️ 每次循环重置，Open() 会悄悄重置它
    src_path = str(Path(item["src"]).resolve())
    src_pres = pptApp.Presentations.Open(src_path, ReadOnly=True, Untitled=True, WithWindow=False)
    pptApp.DisplayAlerts = 0
    src_pres.Slides(item["page"]).Copy()
    src_pres.Close()
    count_before = new_pres.Slides.Count
    new_pres.Windows(1).ViewType = 1
    new_pres.Windows(1).Activate()
    pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")
    if i > 1:                                          # 第一次粘贴替换空白页，count 不增
        start = time.time()
        while new_pres.Slides.Count == count_before:
            if time.time() - start > 8: break
            time.sleep(0.1)
    time.sleep(0.3)

# ⚠️ 不要在这里 delete Slides(1)！见下方陷阱说明
new_pres.SaveAs(str(Path(OUTPUT).resolve()), 24)
print(f"[完成] {OUTPUT}")
```

**素材发现流程**：
1. 直接写 SQL 查 `slide_vault.db`（用 .py 文件执行，不要在命令行直接写中文字符串）
2. 先查 title + keywords，发现不够再加 body_text（探索用）
3. 多份 PPT 有相同内容时，取日期最新的那份

---

## 注意事项

- 电脑必须安装 Microsoft Office 或 WPS（二选一即可，两者均支持 `ExecuteMso("PasteSourceFormatting")`）
- WPS 和 Office **共存**时，引擎会通过 `_get_ppt_app()` 优先使用真正的 PowerPoint（因为 WPS 会劫持 COM 注册）；若只装了 WPS，则回退到 `Dispatch` 直接使用 WPS
- 运行期间不要手动操作 PowerPoint 窗口
- 如果 PowerPoint 没有正确关闭，重新运行前先手动关闭所有 PPT 窗口
- 输出文件如果同名会被覆盖，重要版本请提前重命名备份

### ⚠️ win32com 陷阱：Presentations.Add() 的空白页

`Presentations.Add()` 创建的演示文稿自带 1 页空白页。**第一次** `PasteSourceFormatting`（ViewType=1）会把这个空白页**替换**掉（不是追加），所以粘贴完后 Slides.Count 仍为 1。

**结论：不要在最后 delete Slides(1)**。如果错误地删除，会把第一页内容删掉，最终少一页。
