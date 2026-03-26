---
name: slide-blocks
description: PPT 智能组装助手。当用户说"帮我做一个PPT"、"组装一份汇报"、"从素材里选几页拼一下"、"我需要一个售前演示"、"选几张幻灯片"、"从历史PPT里找内容"、"做行业会议/售前汇报"、"换个模板风格"，或提到具体章节主题和来源文件时，必须使用此 skill。也适用于需要预览素材库内容、查询有哪些可用片段、或对已有组装方案做调整的场景。不要等用户说"用 slide-blocks"才触发——只要涉及 PPT 组装、幻灯片检索、模板套用，就应该主动使用。
---

# SlideBlocks 组装助手

## 初始化：确定 skill 路径

**每次被调用时，首先执行以下代码**，后续所有操作都基于这些变量：

```python
import sys, os
from pathlib import Path

# config.py 的 __file__ 固定指向 skill 根目录下的 slide_vault/config.py
# 所以 skill 根目录 = config.py 所在目录的上一级，无需猜测安装位置
sys.path.insert(0, str(Path(__file__).parent))  # 先临时用任意路径找到 config
from slide_vault.config import CONFIG_PATH
SKILL_DIR = CONFIG_PATH.parent                   # config.yaml 所在目录 = skill 根目录

os.chdir(str(SKILL_DIR))                         # 确保相对路径（如 slide_vault.db）正确解析
sys.path.insert(0, str(SKILL_DIR))
sys.path.insert(0, str(SKILL_DIR / "engine"))

import yaml
_config = yaml.safe_load((SKILL_DIR / "config.yaml").read_text(encoding="utf-8"))
MATERIALS_DIR = _config.get("materials_dir", "")
OUTPUT_DIR    = _config.get("output_dir", "")
TEMPLATE_DIR  = SKILL_DIR / "模板"
```

初始化后，根据用户需求判断模式：

- **Mode A（工具模式）**：统一标题栏、深浅色转换、局部编辑 — 不依赖素材库，`MATERIALS_DIR` 为空也可以直接进行
- **Mode B（素材库模式）**：搜索素材、组装 PPT — 需要 `MATERIALS_DIR` 非空且已运行 `setup_paths.py`

如果 `OUTPUT_DIR` 为空，提醒用户先在 `config.yaml` 填写 `output_dir`（建议填 skill 目录之外的路径，如 `D:/PPT输出`，避免 skill 更新时被覆盖）。

---

## 工作流程

### 第一步：收集需求

向用户确认以下信息（已说明的直接跳过，不要重复询问）：

1. **封面标题**：PPT 的主标题（可帮用户起一个备选）
2. **章节结构**：要哪几个部分，每部分讲什么
3. **素材来源**：指定来源文件，或从库里检索
4. **模板风格**：深色底（默认）还是浅色底

### 第二步：搜索素材库

```python
from slide_vault.search import search_content, search_structural, print_results

# 按内容检索（有标签的页面）
results = search_content(
    scene="售前汇报",           # 或 "行业会议"
    content_type="解决方案",
    keywords=["AI", "智慧医疗"],
    quality_min=4,
    source_file="北大",          # 来源文件关键词（可选）
    limit=10
)
print_results(results, mode="content")
```

**search_content 参数说明：**
- `scene`：场景，如 "售前汇报" / "行业会议"
- `content_type`：如 "解决方案" / "行业背景" / "产品功能" / "标杆案例" / "实施保障"
- `keywords`：关键词列表，匹配幻灯片标题（优先）和 AI 标签；支持同义词自动扩展
- `quality_min`：最低质量分（1-5），建议传 4
- `source_file`：来源文件名关键词，如 "哈密" / "北大" / "福建"
- `limit`：返回条数上限

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

在 `tasks/` 目录下写任务脚本并运行：

```python
# tasks/task_xxx.py
# -*- coding: utf-8 -*-
import sys, os
from pathlib import Path

# 先把 skill 根目录加入 path，才能 import slide_vault
_SKILL_DIR_BOOTSTRAP = Path(__file__).parent.parent
sys.path.insert(0, str(_SKILL_DIR_BOOTSTRAP))

from slide_vault.config import CONFIG_PATH
SKILL_DIR = CONFIG_PATH.parent
os.chdir(str(SKILL_DIR))
sys.path.insert(0, str(SKILL_DIR / "engine"))

import assemble_template, yaml

_config = yaml.safe_load((SKILL_DIR / "config.yaml").read_text(encoding="utf-8"))
OUTPUT_DIR = Path(_config["output_dir"])  # 必须在 config.yaml 里填写，不设默认值

# 浅色底模板取消下面注释：
# assemble_template.TEMPLATE_PATH = SKILL_DIR / "模板/蓝色商务（浅色底）.pptx"
assemble_template.TEMPLATE_PATH = SKILL_DIR / "模板/科技风（深色底）.pptx"
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
import sys, os
from pathlib import Path
from slide_vault.config import CONFIG_PATH
SKILL_DIR = CONFIG_PATH.parent
os.chdir(str(SKILL_DIR))
sys.path.insert(0, str(SKILL_DIR / "engine"))
from edit_pptx import edit

edit("输出/xxx.pptx", [
    {"op": "delete",          "pages": [7, 8]},
    {"op": "move",            "pages": [7, 8], "after": 15},
    {"op": "insert_template", "template_page": 2, "after": 5, "title": "新章节"},
    {"op": "replace",         "page": 12, "src": "路径/文件.pptx", "src_page": 5},
])
```

---

## 整份 PPT 色系转换

将一整份 PPT 从深色底转为浅色底（或反向），无需重新组装：

```python
import sys, os
from pathlib import Path

_SKILL_DIR_BOOTSTRAP = Path(__file__).parent.parent
sys.path.insert(0, str(_SKILL_DIR_BOOTSTRAP))

from slide_vault.config import CONFIG_PATH
SKILL_DIR = CONFIG_PATH.parent
os.chdir(str(SKILL_DIR))
sys.path.insert(0, str(SKILL_DIR / "engine"))

from convert_deck import convert

convert("D:/输入文件.pptx", to="light")   # 深色底 → 浅色底，自动选模板
convert("D:/输入文件.pptx", to="dark")    # 浅色底 → 深色底，自动选模板
```

输出文件自动保存至同目录，文件名加 `_浅色底` / `_深色底` 后缀。颜色修复自动触发。

**封面自动判断**：第一页如果是极简封面（无图表/表格，文字 ≤ 3 个且总字数 ≤ 60），自动替换为模板封面；否则视为内容页，只做色系转换，不替换内容。

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

## 注意事项

- 电脑必须安装 Microsoft Office 或 WPS
- 运行期间不要手动操作 PowerPoint 窗口
- **颜色修复**：按需触发，不主动改动
  - 文件名含"深色底"/"浅色底"关键词 → 与模板方向相反时自动触发
  - 文件名没有关键词，但用户明确说"转成浅色底/深色底" → 在每个内容页 plan 项手动加 `"fix_colors": true`（浅色底方向）或 `"fix_colors_dark": true`（深色底方向）
  - 用户没有提转换需求 → 不触发，保留原始颜色
- 输出同名文件会被覆盖，重要版本提前重命名备份

---

## 素材搜集稿（不套模板）

需要把多份 PPT 里的特定页汇总成原始风格的参考文档时，读取 `references/collect_mode.md`。
