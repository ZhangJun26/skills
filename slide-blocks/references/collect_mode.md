# 素材搜集稿模式（不套模板，原始格式汇总）

**场景**：把多份 PPT 里的特定页汇总成一份参考文档，保留各页原始风格，**不套任何模板**。

此模式**不使用** `assemble_template.py`，直接用整页复制方式。

## 任务脚本模板

```python
# tasks/task_搜集稿_xxx.py
# -*- coding: utf-8 -*-
import sys, time, os
from pathlib import Path

# bootstrap：先把 skill 根目录加入 path，才能 import slide_vault
_SKILL_DIR_BOOTSTRAP = Path(__file__).parent.parent
sys.path.insert(0, str(_SKILL_DIR_BOOTSTRAP))

from slide_vault.config import CONFIG_PATH
SKILL_DIR = CONFIG_PATH.parent
os.chdir(str(SKILL_DIR))
sys.path.insert(0, str(SKILL_DIR / "engine"))
import assemble_template  # 复用 _get_ppt_app()，处理 WPS 劫持

PLAN = [
    {"src": "完整路径/文件名.pptx", "page": 5, "label": "说明"},
    # ...更多页面
]
OUTPUT = str(SKILL_DIR / "输出/搜集稿-xxx.pptx")

pptApp = assemble_template._get_ppt_app()
pptApp.Visible = True
pptApp.DisplayAlerts = 0

new_pres = pptApp.Presentations.Add(WithWindow=True)  # 含一个空白页

for i, item in enumerate(PLAN, 1):
    pptApp.DisplayAlerts = 0                           # 每次循环重置！
    src_pres = pptApp.Presentations.Open(
        str(Path(item["src"]).resolve()),
        ReadOnly=True, Untitled=True, WithWindow=False
    )
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

# ⚠️ 不要 delete Slides(1)！第一次粘贴已替换空白页，删掉会少一页
new_pres.SaveAs(str(Path(OUTPUT).resolve()), 24)
print(f"[完成] {OUTPUT}")
```

## 素材发现方法

1. 直接写 SQL 查 `slide_vault.db`（用 .py 文件执行，不要在命令行直接写中文）
2. 先查 title + keywords，不够再加 body_text（探索用）
3. 多份 PPT 有相同内容时，取日期最新的那份
