# -*- coding: utf-8 -*-
"""
edit_pptx.py — 对已组装的 PPT 文件进行局部编辑

规范：以下操作无需重新组装，直接操作输出文件：
  - 删除页面
  - 调换 / 移动页面顺序
  - 插入模板页（过渡页 / 封底等）
  - 将某页替换为另一个源文件的内容

仅当需要大幅重构章节结构时，才需要重新跑 task_current.py 重新组装。
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))  # engine/ 内互相引用

from assemble_template import (
    get_source_title, get_content_indices, set_template_title,
    paste_slide_with_source_format, paste_shapes_with_source_format,
    TEMPLATE_PATH, TEMPLATE_CONTENT_PAGE, TEMPLATE_CONTENT_PAGE_NO_TITLE,
    PP_VIEW_NORMAL,
)

OUTPUT_DIR = Path(__file__).parent.parent / "输出"


def edit(output_path, operations):
    """
    对已组装的输出文件执行一系列编辑操作（保存并覆盖原文件）。

    output_path : str | Path   输出文件路径
    operations  : dict | list  单个操作 or 操作列表

    支持的操作类型：

    1. 删除页面
        {"op": "delete", "pages": 7}
        {"op": "delete", "pages": [7, 8]}

    2. 移动页面（移到 after 页之后）
        {"op": "move", "pages": [7, 8], "after": 15}

    3. 插入模板页（after 页之后）
        {"op": "insert_template", "template_page": 2, "after": 5, "title": "新章节"}

    4. 替换某页（用另一个源文件的内容，保持模板背景）
        {"op": "replace", "page": 12, "src": "路径.pptx", "src_page": 5}
        {"op": "replace", "page": 12, "src": "路径.pptx", "src_page": 5, "title": "覆盖标题"}
    """
    import win32com.client

    output_path = Path(output_path).resolve()
    if not output_path.exists():
        raise FileNotFoundError(f"找不到文件：{output_path}")

    if isinstance(operations, dict):
        operations = [operations]

    pptApp = win32com.client.Dispatch("PowerPoint.Application")
    pptApp.Visible = True
    pptApp.DisplayAlerts = 0

    pres = None
    try:
        pres = pptApp.Presentations.Open(
            str(output_path), ReadOnly=False, Untitled=False, WithWindow=True
        )
        pres.Windows(1).Activate()
        pres.Windows(1).ViewType = PP_VIEW_NORMAL

        for op in operations:
            kind = op["op"]

            # ── 1. 删除页面 ──────────────────────────────────────────────────
            if kind == "delete":
                pages = op["pages"]
                if isinstance(pages, int):
                    pages = [pages]
                for p in sorted(pages, reverse=True):   # 倒序删，防止索引漂移
                    pres.Slides(p).Delete()
                    print(f"  [删除] 第 {p} 页")

            # ── 2. 移动页面 ──────────────────────────────────────────────────
            elif kind == "move":
                pages = op["pages"]
                after = op["after"]
                if isinstance(pages, int):
                    pages = [pages]

                # 用 SlideID 追踪，避免索引在移动过程中漂移
                page_ids  = [pres.Slides(p).SlideID for p in sorted(pages)]
                anchor_id = pres.Slides(after).SlideID   # 插入锚点

                def find_by_id(sid):
                    for j in range(1, pres.Slides.Count + 1):
                        if pres.Slides(j).SlideID == sid:
                            return j
                    raise ValueError(f"SlideID {sid} 未找到")

                # 先把所有目标页移到末尾（保持相对顺序）
                for sid in page_ids:
                    pres.Slides(find_by_id(sid)).MoveTo(pres.Slides.Count)

                # 再逐页插到锚点之后
                current_anchor_id = anchor_id
                for sid in page_ids:
                    anchor_pos = find_by_id(current_anchor_id)
                    slide_pos  = find_by_id(sid)
                    target_pos = anchor_pos + 1
                    if slide_pos != target_pos:
                        pres.Slides(slide_pos).MoveTo(target_pos)
                    current_anchor_id = sid   # 下一页紧跟在这页后面

                print(f"  [移动] {sorted(pages)} 页 → 第 {after} 页之后")

            # ── 3. 插入模板页 ────────────────────────────────────────────────
            elif kind == "insert_template":
                tmpl_page = op["template_page"]
                after     = op.get("after", pres.Slides.Count)
                title     = op.get("title", "")

                tmpl_pres = pptApp.Presentations.Open(
                    str(TEMPLATE_PATH.resolve()), ReadOnly=True, Untitled=True, WithWindow=False
                )
                tmpl_pres.Slides(tmpl_page).Copy()
                tmpl_pres.Close()

                count_before = pres.Slides.Count
                paste_slide_with_source_format(pptApp, pres, count_before)
                new_idx = pres.Slides.Count

                if title:
                    set_template_title(pres.Slides(new_idx), title)

                pres.Slides(new_idx).MoveTo(after + 1)
                print(f"  [插入模板] P{tmpl_page}「{title}」→ 第 {after + 1} 页")

            # ── 4. 替换某页 ──────────────────────────────────────────────────
            elif kind == "replace":
                page           = op["page"]
                src            = Path(op["src"]).resolve()
                src_page       = op["src_page"]
                override_title = op.get("title")

                # Step 1：分析源文件，决定使用 P3 还是 P4 模板
                src_pres  = pptApp.Presentations.Open(
                    str(src), ReadOnly=True, Untitled=True, WithWindow=False
                )
                src_slide = src_pres.Slides(src_page)
                source_title, title_shape_idx = get_source_title(src_slide)
                source_has_title = source_title is not None

                if source_has_title:
                    tmpl_page       = TEMPLATE_CONTENT_PAGE
                    content_indices = get_content_indices(src_slide, exclude_idx=title_shape_idx)
                    title_text      = override_title or source_title
                else:
                    tmpl_page       = TEMPLATE_CONTENT_PAGE_NO_TITLE
                    content_indices = list(range(1, src_slide.Shapes.Count + 1))
                    title_text      = None
                src_pres.Close()

                # Step 2：复制模板页，粘到末尾（建立背景）
                tmpl_pres = pptApp.Presentations.Open(
                    str(TEMPLATE_PATH.resolve()), ReadOnly=True, Untitled=True, WithWindow=False
                )
                tmpl_pres.Slides(tmpl_page).Copy()
                tmpl_pres.Close()

                count_before = pres.Slides.Count
                paste_slide_with_source_format(pptApp, pres, count_before)
                new_idx = pres.Slides.Count

                # Step 3：粘贴源内容形状
                if content_indices:
                    src_pres = pptApp.Presentations.Open(
                        str(src), ReadOnly=True, Untitled=True, WithWindow=False
                    )
                    src_pres.Slides(src_page).Shapes.Range(content_indices).Copy()
                    src_pres.Close()
                    paste_shapes_with_source_format(pptApp, pres, new_idx)

                # Step 4：写标题
                if source_has_title and title_text:
                    set_template_title(pres.Slides(new_idx), title_text)

                # Step 5：移到目标位置，删掉旧页（旧页此时顺移到 target+1）
                pres.Slides(new_idx).MoveTo(page)
                pres.Slides(page + 1).Delete()

                print(f"  [替换] 第 {page} 页 ← {src.name} P{src_page}  标题:「{title_text or '(无)'}」")

            else:
                print(f"  [警告] 未知操作：{kind}，已跳过")

        pptApp.DisplayAlerts = 0
        pres.Save()
        total = pres.Slides.Count
        print(f"\n[完成] 共 {total} 页  {output_path}")

    finally:
        try:
            if pres is not None:
                pres.Saved = True
                pres.Close()
        except Exception:
            pass
        try:
            pptApp.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    print(__doc__)
