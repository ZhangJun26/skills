# -*- coding: utf-8 -*-
"""
assembler_com.py — 基于 win32com PowerPoint COM 自动化的 PPT 组装模块

策略（v4）：
  模拟用户手动操作：复制 → 粘贴 → 选择"保留源格式"。
  Slides.Paste() 等同于"使用目标主题"，会改变颜色。
  CommandBars.ExecuteMso("PasteSourceFormatting") 才是"保留源格式"粘贴。

用法：
  from slide_vault.assembler_com import assemble
  result = assemble(plan, "输出文件名")
"""

import time
from pathlib import Path
from datetime import datetime

OUTPUT_DIR = Path("D:/Claude/SlideBlocks/输出")


def assemble(plan: list, output_name: str = None) -> Path:
    """
    用 win32com PowerPoint 组装 PPT。

    plan: list of dict
        - src (str/Path)       : 源 PPTX 路径
        - page (int)           : 页码（从 1 开始）
        - replace_title (str)  : （可选）替换该页标题文字

    output_name: 输出文件名（不含扩展名）
    返回：输出文件的 Path
    """
    try:
        import win32com.client
    except ImportError:
        raise ImportError("需要 pywin32：pip install pywin32")

    if not output_name:
        output_name = datetime.now().strftime("%Y%m%d_%H%M%S") + "_assembled"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"{output_name}.pptx"

    pptApp = win32com.client.Dispatch("PowerPoint.Application")
    pptApp.Visible = True  # ExecuteMso 需要应用程序可见

    pres = None
    try:
        print(f"[组装] 共 {len(plan)} 页\n")

        # 新建空演示文稿
        pres = pptApp.Presentations.Add(WithWindow=True)
        while pres.Slides.Count > 0:
            pres.Slides(1).Delete()

        # 切换到普通视图，确保幻灯片缩略图面板处于焦点
        pres.Windows(1).Activate()
        pres.Windows(1).ViewType = 1  # ppViewNormal = 1

        # ── 逐页：复制 + 保留源格式粘贴 ────────────────────────────
        for i, item in enumerate(plan, 1):
            src   = Path(item["src"]).resolve()
            page  = item["page"]
            title = item.get("replace_title")

            # 打开源文件，复制指定页，关闭
            src_pres = pptApp.Presentations.Open(
                str(src), ReadOnly=True, Untitled=True, WithWindow=False
            )
            src_pres.Slides(page).Copy()
            src_pres.Close()

            # 确保目标窗口激活，并导航到最后一页（粘贴插在其后）
            pres.Windows(1).Activate()
            if pres.Slides.Count > 0:
                pres.Windows(1).View.GotoSlide(pres.Slides.Count)

            count_before = pres.Slides.Count

            # 保留源格式粘贴（等同于手动点"保留源格式"选项）
            pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")

            # 等待粘贴完成（ExecuteMso 是异步的，需要等 slide 出现）
            deadline = time.time() + 5.0
            while pres.Slides.Count == count_before and time.time() < deadline:
                time.sleep(0.1)

            if pres.Slides.Count == count_before:
                # ExecuteMso 失败，降级为普通粘贴
                print(f"  [警告] P{i:02d} ExecuteMso 失败，降级为普通粘贴")
                pres.Slides.Paste(pres.Slides.Count + 1)

            if title:
                _replace_title(pres.Slides(pres.Slides.Count), title)

            print(f"  P{i:02d}: {src.name}  第{page}页" + (f"  ->  「{title}」" if title else ""))

        # ── 保存 ────────────────────────────────────────────────────
        pres.SaveAs(str(output_path.resolve()))
        print(f"\n[完成] {output_path}")
        return output_path

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


def _replace_title(slide, new_text: str):
    """
    替换幻灯片的标题占位符文字。
    优先找标题/居中标题占位符，找不到则改最左上角的文本框。
    """
    # ppPlaceholderTitle = 1, ppPlaceholderCenterTitle = 3
    for shape in slide.Shapes:
        try:
            ph_type = shape.PlaceholderFormat.Type
            if ph_type in (1, 3) and shape.HasTextFrame:
                shape.TextFrame.TextRange.Text = new_text
                return
        except Exception:
            pass

    # 备用：找左上角最近的有内容文本框
    best_shape = None
    best_score = float("inf")
    for shape in slide.Shapes:
        try:
            if shape.HasTextFrame and shape.TextFrame.TextRange.Text.strip():
                score = shape.Left + shape.Top
                if score < best_score:
                    best_score = score
                    best_shape = shape
        except Exception:
            pass

    if best_shape:
        best_shape.TextFrame.TextRange.Text = new_text
