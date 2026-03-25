# -*- coding: utf-8 -*-
"""
assemble_template.py — 模板优先的 PPT 组装脚本 v2

策略：
  - 纯模板页（封面/过渡页/封底）：直接复制模板指定页，ExecuteMso 保留源格式粘贴
  - 内容页：先复制模板 P3 建立背景+标题栏，再从源文件复制内容形状，
    切换到 ppViewSlide（幻灯片编辑区）后 ExecuteMso 保留源格式粘贴 → 颜色不变

Plan 项格式：
  纯模板页：{"template_page": N, "replace_title": "xxx"}
  内容页：  {"src": "路径", "page": N, "replace_title": "xxx"}
"""

import time
import sys
import subprocess
from pathlib import Path
from datetime import datetime

# ─── PowerPoint 启动路径（WPS 劫持 COM 时的绕过方案） ─────────────────────────
# WPS 会把自己注册为 PowerPoint.Application 的 COM 提供方，导致 Dispatch() 拿到 WPS。
# 绕过方式：直接启动 POWERPNT.EXE，再用 GetActiveObject 附加真正的 PowerPoint 实例。
# 如果机器上只有 WPS，找不到任何路径，系统会回退到 Dispatch("PowerPoint.Application")。
_POWERPNT_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
    r"C:\Program Files\Microsoft Office\root\Office15\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office15\POWERPNT.EXE",
    r"C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE",
]
POWERPNT_EXE = next((p for p in _POWERPNT_CANDIDATES if Path(p).exists()), None)

# ─── 路径配置 ─────────────────────────────────────────────────────────────────

_BASE_DIR = Path(__file__).parent.parent  # engine/ 的上一层 = SlideBlocks 根目录

TEMPLATE_PATH = _BASE_DIR / "模板" / "科技风（深色底）.pptx"
TEMPLATE_CONTENT_PAGE          = 3  # 带标题栏内容页（模板P3）
TEMPLATE_CONTENT_PAGE_NO_TITLE = 4  # 无标题栏内容页（模板P4）

OUTPUT_DIR = _BASE_DIR / "输出"

# 标题区域阈值（pt）：top 小于此值的形状视为标题栏，不复制进输出
# win32com shape.Top 单位是 pt（不是 EMU）
TITLE_THRESHOLD_PT = 65

# PowerPoint ViewType 常量
PP_VIEW_NORMAL = 1   # 普通视图（缩略图面板+编辑区）
PP_VIEW_SLIDE  = 9   # 幻灯片视图（全屏编辑单张幻灯片，编辑区获得焦点）

# ─── 辅助函数 ─────────────────────────────────────────────────────────────────

def get_source_title(src_slide):
    """从源幻灯片提取标题文字，同时返回该形状的索引（用于精准排除）。
    返回：(title_text, shape_index) 或 (None, None)
    """
    # 优先：标题占位符（ppPlaceholderTitle = 1）
    for j in range(1, src_slide.Shapes.Count + 1):
        shape = src_slide.Shapes(j)
        try:
            if shape.PlaceholderFormat.Type == 1 and shape.HasTextFrame:
                t = shape.TextFrame.TextRange.Text.strip()
                if t:
                    return t, j
        except Exception:
            pass
    # 备选：标题区域（top < TITLE_THRESHOLD_PT）内有文字、top 最小的形状
    candidates = []
    for j in range(1, src_slide.Shapes.Count + 1):
        shape = src_slide.Shapes(j)
        if shape.Top < TITLE_THRESHOLD_PT:
            try:
                if shape.HasTextFrame:
                    t = shape.TextFrame.TextRange.Text.strip()
                    if t:
                        candidates.append((shape.Top, j, t))
            except Exception:
                pass
    if candidates:
        candidates.sort()
        _, j, t = candidates[0]
        return t, j
    return None, None


def get_content_indices(src_slide, exclude_idx=None):
    """返回要复制的形状索引列表。
    排除规则：
    - 标题占位符（ppPlaceholderTitle=1）
    - exclude_idx 指定的标题文字形状
    - 标题区（top < TITLE_THRESHOLD_PT）内的图片/图形形状（type=13，msoPicture），
      即 logo、图标等装饰图片；文字形状保留（如标题区右侧的矩形文字框）
    """
    MSO_PICTURE = 13  # msoPicture
    indices = []
    for j in range(1, src_slide.Shapes.Count + 1):
        shape = src_slide.Shapes(j)
        # 跳过标题占位符
        try:
            if shape.PlaceholderFormat.Type == 1:
                continue
        except Exception:
            pass
        # 跳过检测到的标题形状
        if j == exclude_idx:
            continue
        # 标题区内的图片形状（logo/图标）一律排除
        if shape.Top < TITLE_THRESHOLD_PT and shape.Type == MSO_PICTURE:
            continue
        indices.append(j)
    return indices


def set_template_title(target_slide, text):
    """将文字写入模板标题文本框（top 最小、且当前有非空文字的形状）。
    必须过滤掉空文字的背景矩形，否则会写错目标。
    """
    best, best_top = None, float("inf")
    for j in range(1, target_slide.Shapes.Count + 1):
        shape = target_slide.Shapes(j)
        try:
            if shape.HasTextFrame:
                existing = shape.TextFrame.TextRange.Text.strip()
                if existing and shape.Top < best_top:  # 必须有非空文字
                    best_top = shape.Top
                    best = shape
        except Exception:
            pass
    if best:
        try:
            best.TextFrame.TextRange.Text = text
        except Exception as e:
            print(f"    [警告] 写入标题失败: {e}")


def paste_slide_with_source_format(pptApp, pres, count_before):
    """在幻灯片缩略图面板（PP_VIEW_NORMAL）上下文里，ExecuteMso 粘贴整页（保留源格式）。"""
    pres.Windows(1).Activate()
    pres.Windows(1).ViewType = PP_VIEW_NORMAL
    time.sleep(0.3)
    if pres.Slides.Count > 0:
        pres.Windows(1).View.GotoSlide(pres.Slides.Count)
    pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")
    deadline = time.time() + 5.0
    while pres.Slides.Count == count_before and time.time() < deadline:
        time.sleep(0.1)
    if pres.Slides.Count == count_before:
        print(f"    [警告] ExecuteMso 超时，降级普通粘贴")
        pres.Slides.Paste(pres.Slides.Count + 1)


def paste_shapes_with_source_format(pptApp, pres, current_idx):
    """在幻灯片编辑区（PP_VIEW_SLIDE）上下文里，ExecuteMso 粘贴形状（保留源格式、颜色不变）。"""
    pres.Windows(1).Activate()
    pres.Windows(1).View.GotoSlide(current_idx)
    pres.Windows(1).ViewType = PP_VIEW_SLIDE   # 切到全屏编辑视图，编辑区获得焦点
    time.sleep(0.3)
    pptApp.CommandBars.ExecuteMso("PasteSourceFormatting")
    time.sleep(0.4)
    pres.Windows(1).ViewType = PP_VIEW_NORMAL  # 切回普通视图
    time.sleep(0.2)


# ─── 颜色修复（深色底来源 → 浅色底模板适配） ─────────────────────────────────────

_DARK_HEX = '262626'  # 深灰，替换浅色文字

# 浅色预设颜色名（prstClr val="white" 等），全部替换为深色
_LIGHT_PRESET_COLORS = {
    'white', 'ltGray', 'silver', 'gainsboro', 'ghostWhite',
    'snow', 'ivory', 'floralWhite', 'honeydew', 'azure', 'aliceBlue',
    'lavenderBlush', 'mistyRose', 'seashell', 'linen', 'oldLace',
    'mintCream', 'antiqueWhite', 'cornsilk', 'lemonChiffon',
}

def _is_light_hex(hex_val):
    """判断 6 位十六进制颜色是否为浅色（RGB 各分量 >= 200）"""
    try:
        r, g, b = int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16)
        return r >= 200 and g >= 200 and b >= 200
    except Exception:
        return False


def _hex_luminance(hex_val):
    """计算 6 位十六进制颜色的感知亮度（0-255）"""
    r = int(hex_val[0:2], 16)
    g = int(hex_val[2:4], 16)
    b = int(hex_val[4:6], 16)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _shape_has_dark_fill(shape, threshold=140):
    """
    判断形状是否有深色背景填充（深色 → 白字应保留）。
    threshold=140：亮度低于此值视为深色（0-255 范围）。
    支持 SOLID（纯色）和 GRADIENT（渐变，取所有色标平均亮度）。
    返回 False 时表示浅色/透明/不确定 → 白字应改深色。
    """
    try:
        from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
        fill = shape.fill
        ft = fill.type
        if ft is None or ft == MSO_FILL.BACKGROUND:
            return False  # 透明，底色是模板浅色背景

        if ft == MSO_FILL.SOLID:
            fc = fill.fore_color
            if fc.type == MSO_COLOR_TYPE.RGB:
                return _hex_luminance(str(fc.rgb)) < threshold

        if ft == MSO_FILL.GRADIENT:
            # python-pptx 渐变 API 不完整，直接读 XML 色标
            NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            grad = shape._element.find(f'.//{{{NS_A}}}gradFill')
            if grad is not None:
                stops = grad.findall(f'.//{{{NS_A}}}gs')
                lums = []
                for gs in stops:
                    clr = gs.find(f'{{{NS_A}}}srgbClr')
                    if clr is not None:
                        val = clr.get('val', '')
                        if len(val) == 6:
                            lums.append(_hex_luminance(val))
                if lums:
                    return (sum(lums) / len(lums)) < threshold

        # 图案/主题色 → 保守处理（返回 False）
        return False
    except Exception:
        return False


# 深色底模板里用作"浅色文字"的主题色槽位，粘到浅色底模板后会变成白色（不可见）
_LIGHT_SCHEME_COLORS = {'bg1', 'lt1', 'bg2', 'lt2'}


def _fix_solidfill_el(sf_el, dark_hex):
    """修复一个 <a:solidFill> 元素：浅色/空 → 深色"""
    from lxml import etree
    NS_A      = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    PRSTCLR   = f'{{{NS_A}}}prstClr'
    SRGBCLR   = f'{{{NS_A}}}srgbClr'
    SCHEMECLR = f'{{{NS_A}}}schemeClr'

    children = list(sf_el)
    if not children:
        # 空的 solidFill → 设为深色
        etree.SubElement(sf_el, SRGBCLR).set('val', dark_hex)
        return
    child = children[0]
    if child.tag == PRSTCLR and child.get('val', '') in _LIGHT_PRESET_COLORS:
        sf_el.remove(child)
        etree.SubElement(sf_el, SRGBCLR).set('val', dark_hex)
    elif child.tag == SRGBCLR and _is_light_hex(child.get('val', '')):
        child.set('val', dark_hex)
    elif child.tag == SCHEMECLR and child.get('val', '') in _LIGHT_SCHEME_COLORS:
        # bg1/lt1/bg2/lt2 在浅色底模板里 = 白色，必须换成显式深色
        sf_el.remove(child)
        etree.SubElement(sf_el, SRGBCLR).set('val', dark_hex)


def _fix_gradfill_el(gf_el, dark_hex):
    """
    修复文字 <a:gradFill>：若所有色标均为浅色，整个渐变替换为深色 solidFill。
    有深色色标（刻意设计的彩色渐变）则保留不改。
    """
    from lxml import etree
    NS_A    = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    SRGBCLR = f'{{{NS_A}}}srgbClr'
    SCHEMECLR = f'{{{NS_A}}}schemeClr'
    parent  = gf_el.getparent()
    if parent is None:
        return

    # 收集所有色标的亮度
    lums = []
    for gs in gf_el.iter(f'{{{NS_A}}}gs'):
        clr = gs.find(SRGBCLR)
        if clr is not None:
            val = clr.get('val', '')
            if len(val) == 6:
                lums.append(_hex_luminance(val))
        sc = gs.find(SCHEMECLR)
        if sc is not None and sc.get('val', '') in _LIGHT_SCHEME_COLORS:
            lums.append(255)  # 视为浅色

    # 全部色标都是浅色（平均亮度 > 200）→ 替换为深色 solidFill
    if lums and (sum(lums) / len(lums)) > 200:
        idx = list(parent).index(gf_el)
        parent.remove(gf_el)
        sf = etree.Element(f'{{{NS_A}}}solidFill')
        sc_el = etree.SubElement(sf, SRGBCLR)
        sc_el.set('val', dark_hex)
        parent.insert(idx, sf)


def _fix_text_colors_xml(slide_element, dark_hex=_DARK_HEX):
    """
    直接操作 XML，修复文本 rPr / defRPr 里的浅色颜色。
    处理：srgbClr 近白色、prstClr white、schemeClr bg1/lt1、gradFill 全浅色。
    只改文字属性（rPr/defRPr），不动形状背景填充。
    """
    NS_A      = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    SOLIDFILL = f'{{{NS_A}}}solidFill'
    GRADFILL  = f'{{{NS_A}}}gradFill'
    RPR       = f'{{{NS_A}}}rPr'
    DEFRPR    = f'{{{NS_A}}}defRPr'

    for tag in (RPR, DEFRPR):
        for el in slide_element.iter(tag):
            for sf in el.findall(SOLIDFILL):
                try:
                    _fix_solidfill_el(sf, dark_hex)
                except Exception:
                    pass
            for gf in el.findall(GRADFILL):
                try:
                    _fix_gradfill_el(gf, dark_hex)
                except Exception:
                    pass


def _fix_shape_text_colors_smart(shape, dark_hex=_DARK_HEX):
    """
    智能修复单个形状的文字颜色：
    - 形状有深色背景填充 → 跳过（白字保留，适合深色框）
    - 形状背景浅色/透明  → 修复白字为深色
    递归处理组合形状。
    """
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                _fix_shape_text_colors_smart(child, dark_hex)
            return
    except Exception:
        pass

    if _shape_has_dark_fill(shape):
        return  # 深色背景，白字保留

    try:
        _fix_text_colors_xml(shape._element, dark_hex)
    except Exception:
        pass


def _fix_shape_fills(shape):
    """修复形状背景填充：显式浅色实心填充 → 透明（递归处理组合形状）"""
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        st = shape.shape_type
        if st in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE):
            return
        if st == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                _fix_shape_fills(child)
            return
    except Exception:
        pass
    try:
        from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
        fill = shape.fill
        if fill.type == MSO_FILL.SOLID:
            fc = fill.fore_color
            if fc.type == MSO_COLOR_TYPE.RGB and _is_light_hex(str(fc.rgb)):
                fill.background()
    except Exception:
        pass


def fix_colors_for_light_template(pptx_path, slide_indices):
    """
    对指定页（1-based）修复深色来源的颜色：
      - 文字颜色（XML 级）：prstClr white / srgbClr 近白色 → 深灰 #262626
      - 形状填充（API 级）：显式浅色实心填充 → 透明
    在 win32com 保存完成后、文件关闭后调用。
    """
    if not slide_indices:
        return
    try:
        from pptx import Presentation
    except ImportError:
        print("  [警告] 未安装 python-pptx，跳过颜色修复（pip install python-pptx）")
        return

    prs = Presentation(str(pptx_path))
    for idx in slide_indices:
        slide = prs.slides[idx - 1]
        for shape in slide.shapes:             # 智能文字修复（深色背景形状跳过）
            _fix_shape_text_colors_smart(shape)
        for shape in slide.shapes:             # API 级填充修复
            _fix_shape_fills(shape)
    prs.save(str(pptx_path))
    print(f"  [颜色修复] 已修复第 {slide_indices} 页（深色底来源 → 浅色底适配）")


def _fix_text_to_white_xml(element):
    """
    将元素内深色文字改为白色（XML 级，用于浅色底来源 → 深色底模板适配）。
    只处理 rPr / defRPr 中的 solidFill。
    """
    from lxml import etree
    NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    SF     = f'{{{NS_A}}}solidFill'
    SRGB   = f'{{{NS_A}}}srgbClr'
    PRST   = f'{{{NS_A}}}prstClr'
    RPR    = f'{{{NS_A}}}rPr'
    DEFRPR = f'{{{NS_A}}}defRPr'
    _DARK_PRESET = {'black', 'darkGray'}

    for tag in (RPR, DEFRPR):
        for el in element.iter(tag):
            for sf in el.findall(SF):
                children = list(sf)
                if not children:
                    # 空 solidFill（继承默认色）→ 深色底下继承色为黑，改白
                    etree.SubElement(sf, SRGB).set('val', 'FFFFFF')
                    continue
                child = children[0]
                if child.tag == PRST and child.get('val', '') in _DARK_PRESET:
                    sf.remove(child)
                    etree.SubElement(sf, SRGB).set('val', 'FFFFFF')
                elif child.tag == SRGB:
                    val = child.get('val', '')
                    if len(val) == 6 and _hex_luminance(val) < 80:
                        child.set('val', 'FFFFFF')


def _fix_shape_text_to_dark_template(shape):
    """
    智能处理单个形状（浅色底来源 → 深色底模板方向）：
    - 透明/无填充形状：深色文字 → 白色（背景将变为深色模板底）
    - 有显式填充形状：保持原样（形状自身提供视觉背景）
    递归处理组合形状。
    """
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                _fix_shape_text_to_dark_template(child)
            return
    except Exception:
        pass

    try:
        from pptx.enum.dml import MSO_FILL
        ft = shape.fill.type
        is_transparent = (ft is None or ft == MSO_FILL.BACKGROUND)
    except Exception:
        is_transparent = True

    if not is_transparent:
        return  # 有填充 → 保持原文字颜色

    try:
        _fix_text_to_white_xml(shape._element)
    except Exception:
        pass


def fix_colors_for_dark_template(pptx_path, slide_indices):
    """
    对指定页（1-based）修复浅色来源的颜色（适配深色底模板）：
      - 透明/无填充形状中的深色文字 → 白色
      - 有显式填充形状不改（形状自身提供背景色和对比度）
    在 win32com 保存完成后、文件关闭后调用。
    """
    if not slide_indices:
        return
    try:
        from pptx import Presentation
    except ImportError:
        print("  [警告] 未安装 python-pptx，跳过颜色修复（pip install python-pptx）")
        return

    prs = Presentation(str(pptx_path))
    for idx in slide_indices:
        slide = prs.slides[idx - 1]
        for shape in slide.shapes:
            _fix_shape_text_to_dark_template(shape)
    prs.save(str(pptx_path))
    print(f"  [颜色修复] 已修复第 {slide_indices} 页（浅色底来源 → 深色底适配）")


# ─── 主流程 ───────────────────────────────────────────────────────────────────

def _get_ppt_app():
    """获取真正的 PowerPoint COM 实例（绕过 WPS COM 劫持）。"""
    import win32com.client

    exe = Path(POWERPNT_EXE) if POWERPNT_EXE else None

    if exe and exe.exists():
        # 直接启动真正的 POWERPNT.EXE，再附加
        subprocess.Popen([str(exe)])
        # WPS 可能已在 ROT 中注册为 PowerPoint.Application，
        # 需等待真正的 PowerPoint 注册后覆盖。用路径判断而非版本号。
        for _ in range(40):          # 最多等 20 秒
            time.sleep(0.5)
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
                app_path = str(app.Path).lower()
                if "kingsoft" not in app_path and "wps" not in app_path:
                    return app
            except Exception:
                pass
        raise RuntimeError("启动 PowerPoint 超时，请检查 POWERPNT_EXE 路径")

    # 回退：直接 Dispatch（仅 WPS 可用时）
    return win32com.client.Dispatch("PowerPoint.Application")


def assemble(plan, output_name=None):
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        raise ImportError("需要 pywin32：pip install pywin32")

    if not output_name:
        output_name = datetime.now().strftime("%Y%m%d_%H%M%S") + "_assembled"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"{output_name}.pptx"

    pptApp = _get_ppt_app()
    pptApp.Visible = True
    pptApp.DisplayAlerts = 0  # ppAlertsNone，屏蔽保存时的"压缩图片"等弹窗

    # 自动检测颜色修复方向
    template_is_light = "浅色底" in str(TEMPLATE_PATH)
    template_is_dark  = "深色底" in str(TEMPLATE_PATH)
    fix_light_indices = []  # 深色底来源 → 浅色底模板，需修复白字为深色
    fix_dark_indices  = []  # 浅色底来源 → 深色底模板，需修复深字为白色

    pres = None
    try:
        print(f"[组装] 共 {len(plan)} 页\n")

        pres = pptApp.Presentations.Add(WithWindow=True)
        while pres.Slides.Count > 0:
            pres.Slides(1).Delete()
        pres.Windows(1).Activate()
        pres.Windows(1).ViewType = PP_VIEW_NORMAL

        for i, item in enumerate(plan, 1):
            pptApp.DisplayAlerts = 0  # 每次迭代重置，防止 Open/Close 过程中被重置
            replace_title = item.get("replace_title")

            # ══ 纯模板页（封面 / 过渡页 / 封底） ════════════════════════════
            if "template_page" in item:
                tmpl_page = item["template_page"]
                print(f"  P{i:02d}: [模板 P{tmpl_page}]" + (f"  →  「{replace_title}」" if replace_title else ""))

                tmpl_pres = pptApp.Presentations.Open(
                    str(TEMPLATE_PATH.resolve()), ReadOnly=True, Untitled=True, WithWindow=False
                )
                pptApp.DisplayAlerts = 0
                tmpl_pres.Slides(tmpl_page).Copy()
                tmpl_pres.Close()

                count_before = pres.Slides.Count
                paste_slide_with_source_format(pptApp, pres, count_before)

                if replace_title:
                    set_template_title(pres.Slides(pres.Slides.Count), replace_title)
                continue

            # ══ 整页复制（保留源格式，不套模板框架，用于保留封面/封底） ══════
            if "copy_slide" in item:
                src  = Path(item["copy_slide"]).resolve()
                page = item["page"]
                print(f"  P{i:02d}: [整页保留] {src.name}  第{page}页")

                src_pres = pptApp.Presentations.Open(
                    str(src), ReadOnly=True, Untitled=True, WithWindow=False
                )
                pptApp.DisplayAlerts = 0
                src_pres.Slides(page).Copy()
                src_pres.Close()

                count_before = pres.Slides.Count
                paste_slide_with_source_format(pptApp, pres, count_before)
                continue

            # ══ 内容页（模板背景 + 源内容形状） ══════════════════════════════
            src  = Path(item["src"]).resolve()
            page = item["page"]
            print(f"  P{i:02d}: {src.name}  第{page}页")

            # Step 1：打开源文件，检测标题 + 确定内容形状索引（不复制到剪贴板）
            src_pres = pptApp.Presentations.Open(
                str(src), ReadOnly=True, Untitled=True, WithWindow=False
            )
            pptApp.DisplayAlerts = 0
            src_slide = src_pres.Slides(page)

            # 根据源文件是否有标题栏决定使用哪个模板页（与 replace_title 无关）
            source_title, title_shape_idx = get_source_title(src_slide)
            source_has_title = source_title is not None

            if source_has_title:
                tmpl_page       = TEMPLATE_CONTENT_PAGE           # P3：带标题栏
                content_indices = get_content_indices(src_slide, exclude_idx=title_shape_idx)
                title_text      = replace_title or source_title    # 允许 replace_title 覆盖
            else:
                tmpl_page       = TEMPLATE_CONTENT_PAGE_NO_TITLE  # P4：无标题栏
                content_indices = list(range(1, src_slide.Shapes.Count + 1))  # 复制全部形状
                title_text      = None  # 无标题栏，不写标题

            src_pres.Close()  # 先关闭，不复制（避免覆盖后面的模板页剪贴板）

            # Step 2：复制对应模板页 → 粘贴（建立背景）
            tmpl_pres = pptApp.Presentations.Open(
                str(TEMPLATE_PATH.resolve()), ReadOnly=True, Untitled=True, WithWindow=False
            )
            pptApp.DisplayAlerts = 0
            tmpl_pres.Slides(tmpl_page).Copy()
            tmpl_pres.Close()

            count_before = pres.Slides.Count
            paste_slide_with_source_format(pptApp, pres, count_before)
            current_idx = pres.Slides.Count

            # Step 3：重新打开源文件，复制内容形状 → 粘贴（保留源格式/颜色）
            if content_indices:
                src_pres = pptApp.Presentations.Open(
                    str(src), ReadOnly=True, Untitled=True, WithWindow=False
                )
                pptApp.DisplayAlerts = 0
                src_pres.Slides(page).Shapes.Range(content_indices).Copy()
                src_pres.Close()
                paste_shapes_with_source_format(pptApp, pres, current_idx)
            else:
                print(f"    [信息] 未找到内容形状")

            # Step 4：写入标题（仅带标题栏的页面）
            if source_has_title:
                set_template_title(pres.Slides(current_idx), title_text)
                print(f"        [P{tmpl_page} 带标题栏]  标题: 「{title_text}」")
            else:
                print(f"        [P{tmpl_page} 无标题栏]")

            # Step 5：标记需要颜色修复的页（自动检测 + 支持显式指定）
            src_is_dark  = "深色底" in str(item.get("src", ""))
            src_is_light = "浅色底" in str(item.get("src", ""))
            if item.get("fix_colors",      template_is_light and src_is_dark):
                fix_light_indices.append(current_idx)
            if item.get("fix_colors_dark", template_is_dark  and src_is_light):
                fix_dark_indices.append(current_idx)

        # 保存（24 = ppSaveAsOpenXMLPresentation，明确格式可屏蔽压缩图片弹窗）
        pptApp.DisplayAlerts = 0
        pres.SaveAs(str(output_path.resolve()), 24)
        print(f"\n[完成] {output_path}")

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

    # win32com 关闭后，用 python-pptx 修复颜色（文件已解锁）
    if fix_light_indices:
        fix_colors_for_light_template(output_path, fix_light_indices)
    if fix_dark_indices:
        fix_colors_for_dark_template(output_path, fix_dark_indices)

    return output_path


if __name__ == "__main__":
    print("assemble_template.py 是纯引擎，请通过任务脚本或 Skill 调用 assemble(plan, output_name)")
    print("示例：")
    print("  from assemble_template import assemble")
    print("  assemble(plan, '输出文件名')")
