---
name: slide-blocks
description: PPT 智能组装助手。用于从素材库中挑选幻灯片、拼装成一份完整的、风格统一的 PPT。当用户说"帮我做一个PPT"、"组装一份汇报"、"从素材里选几页拼一下"、"我需要一个售前演示"，或提到具体章节主题和来源文件时，使用此 skill。也适用于需要预览素材库内容、查询有哪些可用片段、或对已有组装方案做调整的场景。
---

# SlideBlocks 组装助手

## 初始化：读取 skill 路径和配置

**每次被调用时，首先执行以下代码确定路径**，后续所有操作都基于这些变量：

```python
import sys
from pathlib import Path
import yaml

# skill 根目录（SKILL.md 所在的文件夹）
# 用 Glob 找到 SKILL.md，取其父目录
import subprocess, json

# 读取 config.yaml
_SKILL_DIR = Path("D:/Claude/skills/slide-blocks")  # 初始默认，下面会自动修正

# 自动找到真正的 skill 目录（兼容不同安装位置）
for candidate in [
    Path.home() / ".claude/skills/slide-blocks",
    Path("D:/Claude/skills/slide-blocks"),
    Path("C:/Users") / Path.home().name / ".claude/skills/slide-blocks",
]:
    if (candidate / "SKILL.md").exists():
        _SKILL_DIR = candidate
        break

_CONFIG = yaml.safe_load((_SKILL_DIR / "config.yaml").read_text(encoding="utf-8"))

SKILL_DIR    = _SKILL_DIR
DB_PATH      = SKILL_DIR / _CONFIG.get("db_path", "slide_vault.db")
MATERIALS_DIR = _CONFIG.get("materials_dir", "")
OUTPUT_DIR   = Path(_CONFIG.get("output_dir", str(SKILL_DIR / "输出")))
TEMPLATE_DIR = SKILL_DIR / "模板"

sys.path.insert(0, str(SKILL_DIR))
sys.path.insert(0, str(SKILL_DIR / "engine"))
```

如果 `MATERIALS_DIR` 为空，提醒用户先填写 `config.yaml` 并运行 `setup_paths.py`。

---

## 工作流程

### 第一步：收集需求

向用户确认以下信息（没有明确说明的才询问，已说明的直接跳过）：

1. **封面标题**：PPT 的主标题是什么？（可以帮用户起一个备选）
2. **章节结构**：要哪几个部分？每部分大概讲什么？
3. **素材来源**：指定来源文件，或让系统从库里检索
4. **模板风格**：深色底（默认）还是浅色底？

### 第二步：搜索素材库

```python
from slide_vault.search import search_content, search_structural, print_results

# 按内容检索（有标签的页面）
results = search_content(
    scene="售前汇报",           # 或 "行业会议"
    content_type="解决方案",     # 内容类型
    keywords=["AI", "智慧医疗"], # 关键词
    quality_min=4,
    source_file="北大",          # 来源文件关键词（可选）
    limit=10,
    db_path=str(DB_PATH)
)
print_results(results, mode="content")
```

**search_content 可用参数：**
- `scene`：场景，如 "售前汇报" / "行业会议"
- `content_type`：内容类型，如 "解决方案" / "行业背景" / "产品功能" / "标杆案例" / "实施保障"
- `keywords`：关键词列表，匹配幻灯片标题（优先）和 AI 标签；支持同义词自动扩展
- `quality_min`：最低质量分（1-5），建议传 4
- `source_file`：来源文件名关键词
- `limit`：返回条数上限
- `db_path`：数据库路径（传入上面的 `DB_PATH`）

将检索结果整理成表格展示给用户：

| 页 | 来源文件 | 内容 | 质量 |
|----|---------|------|------|
| P05 | 北大医疗交流 | 复星健康双SaaS云模式 | ★★★★★ |

来源文件名取最具辨识度的关键词，不写完整路径。

### 第三步：构建 PLAN

```python
PLAN = [
    # 封面（必须，模板P1）
    {"template_page": 1, "replace_title": "PPT主标题"},

    # 第一章节：开篇不加过渡页，直接上内容
    {"src": "完整路径/文件名.pptx", "page": 5},
    {"src": "完整路径/文件名.pptx", "page": 6},

    # 第二章节及之后：每章前加过渡页（模板P2）
    {"template_page": 2, "replace_title": "章节名称"},
    {"src": "完整路径/文件名.pptx", "page": 12},

    # 封底（必须，模板P5）
    {"template_page": 5},
]
```

**组装规则：**
- 封面后第一部分：直接接内容，**不加**过渡页
- 第二部分及之后：每章前加 `{"template_page": 2, "replace_title": "章节名"}`
- 封底：每份 PPT 最后必须加 `{"template_page": 5}`

### 第四步：执行组装

在 `tasks/` 目录下写任务脚本并运行（`tasks/` 放在 SKILL_DIR 下）：

```python
# tasks/task_xxx.py
# -*- coding: utf-8 -*-
import sys
from pathlib import Path

_SKILL_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(_SKILL_DIR / "engine"))

import assemble_template
import yaml

_config = yaml.safe_load((_SKILL_DIR / "config.yaml").read_text(encoding="utf-8"))
OUTPUT_DIR = Path(_config.get("output_dir", str(_SKILL_DIR / "输出")))

# 浅色底模板取消下面注释：
# assemble_template.TEMPLATE_PATH = _SKILL_DIR / "模板/蓝色商务（浅色底）.pptx"
assemble_template.TEMPLATE_PATH = _SKILL_DIR / "模板/科技风（深色底）.pptx"
assemble_template.OUTPUT_DIR = OUTPUT_DIR

PLAN = [
    # ... 填入构建好的 PLAN
]

if __name__ == "__main__":
    assemble_template.assemble(PLAN, "输出文件名")
```

```bash
cd <SKILL_DIR> && python tasks/task_xxx.py
```

### 第五步：告知用户结果

完成后告知：
- 输出文件位置
- 共几页，结构是什么
- 如需小幅调整，用 edit_pptx 直接编辑，无需重新组装

---

## 局部编辑

```python
import sys
from pathlib import Path
_SKILL_DIR = Path("...自动确定...")
sys.path.insert(0, str(_SKILL_DIR / "engine"))
from edit_pptx import edit

edit("输出/xxx.pptx", [
    {"op": "delete",          "pages": [7, 8]},
    {"op": "move",            "pages": [7, 8], "after": 15},
    {"op": "insert_template", "template_page": 2, "after": 5, "title": "新章节"},
    {"op": "replace",         "page": 12, "src": "路径/文件.pptx", "src_page": 5},
])
```

---

## 模板页说明

| 页码 | 用途 |
|------|------|
| P1 | 封面 |
| P2 | 过渡页（章节标题页） |
| P3 | 带标题栏内容页（引擎自动选择） |
| P4 | 无标题栏内容页（引擎自动选择） |
| P5 | 封底 |

P3/P4 由引擎自动检测，不需要在 PLAN 里手动指定。

---

## 素材搜集稿（不套模板，原始格式汇总）

把多份 PPT 里的特定页汇总成参考文档，保留原始风格，不套模板：

```python
import sys, time
from pathlib import Path

_SKILL_DIR = Path("...自动确定...")
sys.path.insert(0, str(_SKILL_DIR / "engine"))
import assemble_template

PLAN = [
    {"src": "完整路径/文件名.pptx", "page": 5, "label": "说明"},
]
OUTPUT = str(_SKILL_DIR / "输出/搜集稿-xxx.pptx")

pptApp = assemble_template._get_ppt_app()
pptApp.Visible = True
pptApp.DisplayAlerts = 0

new_pres = pptApp.Presentations.Add(WithWindow=True)

for i, item in enumerate(PLAN, 1):
    pptApp.DisplayAlerts = 0
    src_pres = pptApp.Presentations.Open(str(Path(item["src"]).resolve()), ReadOnly=True, Untitled=True, WithWindow=False)
    pptApp.DisplayAlerts = 0
    src_pres.Slides(item["page"]).Copy()
    src_pres.Close()
    count_before = new_pres.Slides.Count
    new_pres.Windows(1).ViewType = 1
    new_pres.Windows(1).Activate()
    pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")
    if i > 1:
        start = time.time()
        while new_pres.Slides.Count == count_before:
            if time.time() - start > 8: break
            time.sleep(0.1)
    time.sleep(0.3)

new_pres.SaveAs(str(Path(OUTPUT).resolve()), 24)
print(f"[完成] {OUTPUT}")
```

---

## 注意事项

- 电脑必须安装 Microsoft Office 或 WPS
- 运行期间不要手动操作 PowerPoint 窗口
- 深色底素材 + 浅色底模板：在 plan 项加 `"fix_colors": true` 可自动修复白色文字
- 输出同名文件会被覆盖，重要版本提前重命名备份
- `Presentations.Add()` 自带 1 个空白页，第一次粘贴会替换它，**不要在最后手动删 Slides(1)**
