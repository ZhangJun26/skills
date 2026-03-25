# -*- coding: utf-8 -*-
"""
assembler.py — PPT 组装模块（ZIP 级操作，完整保留原始格式）

核心原则：
  直接操作 .pptx 的 zip 结构，保证颜色/字体/图片/图表完全不变。

组装策略（v3）：
  对每一页，用 extract_pages 提取成独立单页临时 PPTX，
  然后把临时文件里的所有内容（slide/layout/master/theme/media）
  整体搬入 dest，只做重编号以避免路径冲突，不做任何内容去重。
  PowerPoint 完全支持一份文件内存在多个 master/layout，故此方案可靠。
"""

import zipfile
import os
import re
import shutil
import tempfile
import hashlib
import gc
from lxml import etree
from pathlib import Path
from datetime import datetime

OUTPUT_DIR = Path("D:/Claude/SlideBlocks/输出")

# XML 命名空间
NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT  = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_SLIDE  = NS_R + "/slide"
REL_LAYOUT = NS_R + "/slideLayout"
REL_MASTER = NS_R + "/slideMaster"
REL_THEME  = NS_R + "/theme"
REL_NOTES  = NS_R + "/notesSlide"


# ─── Presentation 级别读取 ─────────────────────────────────────────

def _parse_pres_rels(z: zipfile.ZipFile) -> dict:
    """返回 {rId: slide_path}，只含 slide 类型"""
    data = z.read("ppt/_rels/presentation.xml.rels")
    tree = etree.fromstring(data)
    result = {}
    for rel in tree.findall(f"{{{NS_REL}}}Relationship"):
        rtype = rel.get("Type", "")
        if rtype == REL_SLIDE:
            rid    = rel.get("Id")
            target = rel.get("Target", "")
            if not target.startswith("ppt/"):
                target = "ppt/" + target
            result[rid] = target
    return result


def _get_ordered_slides(z: zipfile.ZipFile) -> list:
    rid_to_path = _parse_pres_rels(z)
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    sldIdLst = pres.find(f".//{{{NS_P}}}sldIdLst")
    ordered = []
    if sldIdLst is not None:
        for child in sldIdLst:
            rid = child.get(f"{{{NS_R}}}id")
            if rid in rid_to_path:
                ordered.append(rid_to_path[rid])
    return ordered


def _max_sld_id(z: zipfile.ZipFile) -> int:
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    sldIdLst = pres.find(f".//{{{NS_P}}}sldIdLst")
    ids = [int(c.get("id", 0)) for c in (sldIdLst if sldIdLst is not None else [])]
    return max(ids, default=255)


def _max_rid_num(z: zipfile.ZipFile, rels_path: str) -> int:
    if rels_path not in z.namelist():
        return 0
    tree = etree.fromstring(z.read(rels_path))
    nums = [int(m.group()) for rel in tree.findall(f"{{{NS_REL}}}Relationship")
            if (m := re.search(r"\d+", rel.get("Id", "")))]
    return max(nums, default=0)


def _max_num(names: set, pattern: str) -> int:
    return max(
        (int(m.group(1)) for f in names if (m := re.fullmatch(pattern, f))),
        default=0
    )


# ─── extract_pages（从同一 PPTX 提取多页）────────────────────────

def extract_pages(src_path: Path, pages: list, output_path: Path):
    """
    从 src_path 提取指定页（1-based），保持页序输出到 output_path。
    同步清理 presentation.xml.rels、[Content_Types].xml、
    以及每个 slide 内部跳转引用，保证输出文件无孤儿引用。
    """
    shutil.copy2(src_path, output_path)

    with zipfile.ZipFile(src_path, "r") as z:
        ordered = _get_ordered_slides(z)

    keep_paths = {ordered[p - 1] for p in pages if 1 <= p <= len(ordered)}

    tmp = str(output_path) + ".tmp"
    with zipfile.ZipFile(str(output_path), "r") as src_z, \
         zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as dst_z:

        for item in src_z.infolist():
            fn = item.filename

            if re.fullmatch(r"ppt/slides/slide\d+\.xml", fn) and fn not in keep_paths:
                continue
            if re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", fn):
                slide_fn = fn.replace("/_rels/", "/").replace(".rels", "")
                if slide_fn not in keep_paths:
                    continue

            data = src_z.read(fn)

            if fn == "ppt/presentation.xml":
                data = _reorder_pres_xml(data, src_z, ordered, pages)
            if fn == "ppt/_rels/presentation.xml.rels":
                data = _strip_pres_rels(data, keep_paths)
            if fn == "[Content_Types].xml":
                data = _strip_ct_xml(data, keep_paths)
            if re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", fn):
                slide_fn = fn.replace("/_rels/", "/").replace(".rels", "")
                if slide_fn in keep_paths:
                    data = _strip_slide_internal_refs(data, slide_fn, keep_paths)

            dst_z.writestr(item, data)

    gc.collect()
    if os.path.exists(str(output_path)):
        os.remove(str(output_path))
    os.rename(tmp, str(output_path))


def _reorder_pres_xml(pres_data, z, ordered, pages):
    rid_to_path = _parse_pres_rels(z)
    path_to_rid = {v: k for k, v in rid_to_path.items()}
    tree = etree.fromstring(pres_data)
    sldIdLst = tree.find(f".//{{{NS_P}}}sldIdLst")
    if sldIdLst is None:
        return pres_data
    path_to_elem = {}
    for child in list(sldIdLst):
        rid = child.get(f"{{{NS_R}}}id")
        if rid in rid_to_path:
            path_to_elem[rid_to_path[rid]] = child
    for child in list(sldIdLst):
        sldIdLst.remove(child)
    for p in pages:
        if 1 <= p <= len(ordered):
            sp = ordered[p - 1]
            if sp in path_to_elem:
                sldIdLst.append(path_to_elem[sp])
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _strip_pres_rels(rels_data, keep_paths):
    tree = etree.fromstring(rels_data)
    for rel in list(tree.findall(f"{{{NS_REL}}}Relationship")):
        if rel.get("Type", "") != REL_SLIDE:
            continue
        target = rel.get("Target", "")
        abs_target = target if target.startswith("ppt/") else "ppt/" + target
        if abs_target not in keep_paths:
            tree.remove(rel)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _strip_ct_xml(ct_data, keep_paths):
    tree = etree.fromstring(ct_data)
    for ov in list(tree.findall(f"{{{NS_CT}}}Override")):
        pn = ov.get("PartName", "").lstrip("/")
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", pn) and pn not in keep_paths:
            tree.remove(ov)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _strip_slide_internal_refs(rels_data, slide_path, keep_paths):
    tree = etree.fromstring(rels_data)
    slide_dir = os.path.dirname(slide_path)
    for rel in list(tree.findall(f"{{{NS_REL}}}Relationship")):
        if rel.get("Type", "") != REL_SLIDE:
            continue
        target = rel.get("Target", "")
        abs_target = os.path.normpath(os.path.join(slide_dir, target)).replace("\\", "/")
        if abs_target not in keep_paths:
            tree.remove(rel)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


# ─── 标题文字替换 ──────────────────────────────────────────────────

def _replace_title_in_xml(slide_xml: bytes, new_text: str) -> bytes:
    tree = etree.fromstring(slide_xml)
    all_sp = tree.findall(f".//{{{NS_P}}}sp")
    for sp in all_sp:
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None and ph.get("idx", "0") in ("0", ""):
            if ph.get("type") in (None, "title", "ctrTitle"):
                _set_sp_text(sp, new_text)
                return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
    best_sp, best_score = None, float("inf")
    for sp in all_sp:
        off = sp.find(f".//{{{NS_A}}}off")
        if off is not None and _sp_has_text(sp):
            score = int(off.get("x", 0)) + int(off.get("y", 0))
            if score < best_score:
                best_score = score
                best_sp = sp
    if best_sp is not None:
        _set_sp_text(best_sp, new_text)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _sp_has_text(sp) -> bool:
    return any(t.text and t.text.strip() for t in sp.findall(f".//{{{NS_A}}}t"))


def _set_sp_text(sp, new_text: str):
    all_t = sp.findall(f".//{{{NS_A}}}t")
    if not all_t:
        return
    all_t[0].text = new_text
    for t in all_t[1:]:
        t.text = ""


# ─── 路径重命名映射 ────────────────────────────────────────────────

# 编号型路径的重命名规则（pattern, format_string, offset_key）
_NUMBERED_PATTERNS = [
    (r"ppt/slides/slide(\d+)\.xml",
     "ppt/slides/slide{n}.xml"),
    (r"ppt/slides/_rels/slide(\d+)\.xml\.rels",
     "ppt/slides/_rels/slide{n}.xml.rels"),
    (r"ppt/slideLayouts/slideLayout(\d+)\.xml",
     "ppt/slideLayouts/slideLayout{n}.xml"),
    (r"ppt/slideLayouts/_rels/slideLayout(\d+)\.xml\.rels",
     "ppt/slideLayouts/_rels/slideLayout{n}.xml.rels"),
    (r"ppt/slideMasters/slideMaster(\d+)\.xml",
     "ppt/slideMasters/slideMaster{n}.xml"),
    (r"ppt/slideMasters/_rels/slideMaster(\d+)\.xml\.rels",
     "ppt/slideMasters/_rels/slideMaster{n}.xml.rels"),
    (r"ppt/theme/theme(\d+)\.xml",
     "ppt/theme/theme{n}.xml"),
]

# 每种编号型路径对应的 offset key
_OFFSET_KEY = {
    "ppt/slides/slide":                   "slide",
    "ppt/slides/_rels/slide":             "slide",
    "ppt/slideLayouts/slideLayout":       "layout",
    "ppt/slideLayouts/_rels/slideLayout": "layout",
    "ppt/slideMasters/slideMaster":       "master",
    "ppt/slideMasters/_rels/slideMaster": "master",
    "ppt/theme/theme":                    "theme",
}


def _build_path_map(src_names: set, dest_names: set) -> dict:
    """
    为 src 里的每个文件计算在 dest 里的新路径（重编号后）。
    编号型：直接加 offset；媒体/tags：若名称冲突则加数字后缀。
    """
    offsets = {
        "slide":  _max_num(dest_names, r"ppt/slides/slide(\d+)\.xml"),
        "layout": _max_num(dest_names, r"ppt/slideLayouts/slideLayout(\d+)\.xml"),
        "master": _max_num(dest_names, r"ppt/slideMasters/slideMaster(\d+)\.xml"),
        "theme":  _max_num(dest_names, r"ppt/theme/theme(\d+)\.xml"),
    }

    path_map = {}
    dest_all = set(dest_names)

    for name in src_names:
        if name.endswith("/"):
            continue

        matched = False
        for pattern, fmt in _NUMBERED_PATTERNS:
            m = re.fullmatch(pattern, name)
            if m:
                key = None
                for prefix, k in _OFFSET_KEY.items():
                    if name.startswith(prefix):
                        key = k
                        break
                n = int(m.group(1)) + offsets.get(key, 0)
                path_map[name] = fmt.format(n=n)
                matched = True
                break

        if not matched:
            # 媒体 / tags：尝试保留原名，冲突则加 _N 后缀
            if name.startswith("ppt/media/") or name.startswith("ppt/tags/"):
                if name not in dest_all:
                    path_map[name] = name
                else:
                    stem = Path(name).stem
                    ext  = Path(name).suffix
                    pref = name.rsplit("/", 1)[0] + "/"
                    i = 2
                    while True:
                        candidate = f"{pref}{stem}_{i}{ext}"
                        if candidate not in dest_all and candidate not in path_map.values():
                            path_map[name] = candidate
                            break
                        i += 1
                dest_all.add(path_map[name])

    return path_map


# ─── Rels 文件重写 ────────────────────────────────────────────────

def _owner_of_rels(rels_path: str) -> str:
    """ppt/slides/_rels/slide1.xml.rels  →  ppt/slides/slide1.xml"""
    p = rels_path.replace("/_rels/", "/")
    if p.endswith(".rels"):
        p = p[:-5]
    return p


def _rewrite_rels(rels_data: bytes, rels_path: str, path_map: dict) -> bytes:
    """
    把 rels 文件里所有 Target 按 path_map 重定向。
    同时丢弃 notes 和 slide 内部跳转引用。
    """
    owner     = _owner_of_rels(rels_path)
    owner_dir = os.path.dirname(owner)

    new_owner     = path_map.get(owner, owner)
    new_owner_dir = os.path.dirname(new_owner)

    tree      = etree.fromstring(rels_data)
    to_remove = []

    for rel in tree.findall(f"{{{NS_REL}}}Relationship"):
        rtype  = rel.get("Type", "")
        target = rel.get("Target", "")
        mode   = rel.get("TargetMode", "")

        # 丢弃备注页
        if rtype == REL_NOTES:
            to_remove.append(rel)
            continue
        # 丢弃 slide→slide 内部跳转（但保留 layout / master 引用）
        if rtype == REL_SLIDE:
            to_remove.append(rel)
            continue
        # 外部链接保持原样
        if mode == "External":
            continue

        # 解析绝对路径 → 重命名 → 转回相对路径
        abs_src = os.path.normpath(os.path.join(owner_dir, target)).replace("\\", "/")
        abs_dst = path_map.get(abs_src, abs_src)
        new_rel = os.path.relpath(abs_dst, new_owner_dir).replace("\\", "/")
        rel.set("Target", new_rel)

    for rel in to_remove:
        tree.remove(rel)

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


# ─── Presentation 级别修改辅助 ────────────────────────────────────

def _append_slide_to_pres_xml(pres_data: bytes, new_rid: str, new_id: int) -> bytes:
    tree = etree.fromstring(pres_data)
    sldIdLst = tree.find(f".//{{{NS_P}}}sldIdLst")
    if sldIdLst is None:
        return pres_data
    el = etree.SubElement(sldIdLst, f"{{{NS_P}}}sldId")
    el.set("id", str(new_id))
    el.set(f"{{{NS_R}}}id", new_rid)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _append_master_to_pres_xml(pres_data: bytes, master_rid: str) -> bytes:
    tree = etree.fromstring(pres_data)
    lst  = tree.find(f".//{{{NS_P}}}sldMasterIdLst")
    if lst is None:
        return pres_data
    existing = [int(c.get("id", 0)) for c in lst]
    new_id   = max(existing, default=2147483647) + 1
    el = etree.SubElement(lst, f"{{{NS_P}}}sldMasterId")
    el.set("id", str(new_id))
    el.set(f"{{{NS_R}}}id", master_rid)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_pres_slide_rel(rels_data: bytes, rid: str, slide_path: str) -> bytes:
    tree = etree.fromstring(rels_data)
    el   = etree.SubElement(tree, f"{{{NS_REL}}}Relationship")
    el.set("Id",     rid)
    el.set("Type",   REL_SLIDE)
    # Target 相对于 ppt/
    rel_target = slide_path[4:] if slide_path.startswith("ppt/") else slide_path
    el.set("Target", rel_target)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_pres_master_rel(rels_data: bytes, rid: str, master_path: str) -> bytes:
    tree = etree.fromstring(rels_data)
    el   = etree.SubElement(tree, f"{{{NS_REL}}}Relationship")
    el.set("Id",     rid)
    el.set("Type",   REL_MASTER)
    rel_target = master_path[4:] if master_path.startswith("ppt/") else master_path
    el.set("Target", rel_target)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_ct_overrides(ct_data: bytes, entries: list) -> bytes:
    """entries: list of (PartName_with_leading_slash, ContentType)"""
    CONTENT_TYPES = {
        "slide":   "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        "layout":  "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
        "master":  "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
        "theme":   "application/vnd.openxmlformats-officedocument.drawingml.theme+xml",
    }
    tree = etree.fromstring(ct_data)
    for part_name, ct_key in entries:
        ct = CONTENT_TYPES.get(ct_key, ct_key)
        el = etree.SubElement(tree, f"{{{NS_CT}}}Override")
        el.set("PartName",    part_name)
        el.set("ContentType", ct)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


# ─── 核心：追加单页（原封不动，重编号） ──────────────────────────

def _append_slide_from_source(
    dest_path: Path,
    src_path: Path,
    src_page_1based: int,
    replace_title: str = None,
):
    """
    从 src_path 取第 src_page_1based 页，追加到 dest_path 末尾。

    策略（v3 简化版）：
      1. 用 extract_pages 把该页提取为独立临时 PPTX
      2. 把临时 PPTX 里所有内容搬入 dest，仅重编号以避免冲突
      3. 更新 presentation.xml / rels / [Content_Types].xml
    """
    tmp = Path(tempfile.mktemp(suffix=".pptx"))
    try:
        # ── 步骤 1：提取单页 ─────────────────────────────────────────
        extract_pages(src_path, [src_page_1based], tmp)
        if replace_title:
            _apply_title_to_file(tmp, replace_title)

        # ── 步骤 2：读取 dest 现状 ───────────────────────────────────
        with zipfile.ZipFile(str(dest_path), "r") as dz:
            dest_names   = set(dz.namelist())
            max_pres_rid = _max_rid_num(dz, "ppt/_rels/presentation.xml.rels")
            new_sld_id   = _max_sld_id(dz) + 1

        # ── 步骤 3：构建路径重映射 ───────────────────────────────────
        with zipfile.ZipFile(str(tmp), "r") as sz:
            src_names = set(sz.namelist())

        path_map = _build_path_map(src_names, dest_names)

        # 找出新 slide 路径、新 master 路径列表
        new_slide_path   = next(v for k, v in path_map.items()
                                if re.fullmatch(r"ppt/slides/slide\d+\.xml", k))
        new_master_paths = sorted(v for k, v in path_map.items()
                                  if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", k))

        # 分配 rId：slide 用 +1，每个 master 再依次 +1
        slide_rid  = f"rId{max_pres_rid + 1}"
        master_rids = [f"rId{max_pres_rid + 2 + i}" for i in range(len(new_master_paths))]

        # ── 步骤 4：读取 src 内容并重写 rels ─────────────────────────
        files_to_add: dict[str, bytes] = {}
        with zipfile.ZipFile(str(tmp), "r") as sz:
            for name in sz.namelist():
                if name.endswith("/"):
                    continue
                # 跳过 notes slides（保留 layout 里的 notes layout 不跳过，因为 master 引用它）
                if re.fullmatch(r"ppt/notesSlides/.*", name):
                    continue
                # 跳过 presentation 级别的 meta 文件（单独处理）
                if name in ("ppt/presentation.xml",
                            "ppt/_rels/presentation.xml.rels",
                            "[Content_Types].xml",
                            "_rels/.rels"):
                    continue

                new_name = path_map.get(name)
                if new_name is None:
                    continue  # 不在 path_map 里的文件（如 docProps）跳过

                data = sz.read(name)
                if name.endswith(".rels"):
                    data = _rewrite_rels(data, name, path_map)

                files_to_add[new_name] = data

        # ── 步骤 5：写入新 dest ──────────────────────────────────────
        ct_entries = []
        ct_entries.append((f"/{new_slide_path}", "slide"))
        for mp in new_master_paths:
            ct_entries.append((f"/{mp}", "master"))
        for k, v in path_map.items():
            if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", k):
                ct_entries.append((f"/{v}", "layout"))
            if re.fullmatch(r"ppt/theme/theme\d+\.xml", k):
                ct_entries.append((f"/{v}", "theme"))

        dest_tmp = str(dest_path) + ".tmp"
        with zipfile.ZipFile(str(dest_path), "r") as src_z, \
             zipfile.ZipFile(dest_tmp, "w", compression=zipfile.ZIP_DEFLATED) as dst_z:

            for item in src_z.infolist():
                fn   = item.filename
                data = src_z.read(fn)

                if fn == "ppt/presentation.xml":
                    data = _append_slide_to_pres_xml(data, slide_rid, new_sld_id)
                    for mr in master_rids:
                        data = _append_master_to_pres_xml(data, mr)

                elif fn == "ppt/_rels/presentation.xml.rels":
                    data = _add_pres_slide_rel(data, slide_rid, new_slide_path)
                    for mr, mp in zip(master_rids, new_master_paths):
                        data = _add_pres_master_rel(data, mr, mp)

                elif fn == "[Content_Types].xml":
                    data = _add_ct_overrides(data, ct_entries)

                dst_z.writestr(item, data)

            for path, content in files_to_add.items():
                dst_z.writestr(path, content)

        gc.collect()
        os.remove(str(dest_path))
        os.rename(dest_tmp, str(dest_path))

    finally:
        if tmp.exists():
            tmp.unlink()


# ─── 标题替换（对已有文件的第一页）──────────────────────────────

def _apply_title_to_file(pptx_path: Path, new_title: str):
    with zipfile.ZipFile(str(pptx_path), "r") as z:
        ordered   = _get_ordered_slides(z)
        if not ordered:
            return
        slide_path = ordered[0]
        slide_xml  = z.read(slide_path)

    new_xml = _replace_title_in_xml(slide_xml, new_title)

    tmp = str(pptx_path) + ".tmp"
    with zipfile.ZipFile(str(pptx_path), "r") as src_z, \
         zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as dst_z:
        for item in src_z.infolist():
            data = new_xml if item.filename == slide_path else src_z.read(item.filename)
            dst_z.writestr(item, data)

    gc.collect()
    os.remove(str(pptx_path))
    os.rename(tmp, str(pptx_path))


# ─── 主入口 ───────────────────────────────────────────────────────

def assemble(plan: list, output_name: str = None) -> Path:
    """
    组装 PPT。

    plan: list of dict
        - src (str/Path)       : 源 PPTX 路径
        - page (int)           : 页码（从 1 开始）
        - replace_title (str)  : （可选）替换该页标题文字

    output_name: 输出文件名（不含扩展名）
    """
    if not output_name:
        output_name = datetime.now().strftime("%Y%m%d_%H%M%S") + "_assembled"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"{output_name}.pptx"

    print(f"[组装] 共 {len(plan)} 页\n")

    # 第一页：直接 extract 作为基础文件
    first = plan[0]
    src0  = Path(first["src"])
    pg0   = first["page"]
    t0    = first.get("replace_title")
    extract_pages(src0, [pg0], output_path)
    if t0:
        _apply_title_to_file(output_path, t0)
    print(f"  P01: {src0.name}  第{pg0}页" + (f"  ->  「{t0}」" if t0 else ""))

    # 后续页：逐一追加
    for i, item in enumerate(plan[1:], 2):
        src   = Path(item["src"])
        page  = item["page"]
        title = item.get("replace_title")
        print(f"  P{i:02d}: {src.name}  第{page}页" + (f"  ->  「{title}」" if title else ""))
        _append_slide_from_source(output_path, src, page, title)

    print(f"\n[完成] {output_path}")
    return output_path


# ─── 保留旧接口（兼容性）─────────────────────────────────────────

def insert_external_slide(dest_path: Path, src_path: Path,
                           src_page_1based: int, insert_after: int,
                           replace_title: str = None):
    """已废弃，内部转发到 _append_slide_from_source。"""
    _append_slide_from_source(dest_path, src_path, src_page_1based, replace_title)
