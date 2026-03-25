"""
search.py - 素材检索模块

两种检索模式：
  search_content()    - 语义检索（需要打过标签的内容页）
  search_structural() - 结构检索（按文件名识别 layout / 背景色）

返回格式统一兼容 assembler.py 的 plan：
  {"src": file_path, "page": slide_index, "title": ..., ...}
"""

import sqlite3
import json
from pathlib import Path

from .config import get_db_path


# ─── 同义词扩展表 ────────────────────────────────────────────────────
# 搜索任意一个词时，自动也搜索其同义词
_SYNONYMS: dict[str, list[str]] = {
    "一体化": ["一体"],
    "一体":   ["一体化"],
    "AI":     ["人工智能"],
    "人工智能": ["AI"],
}

def _expand_keywords(keywords: list[str]) -> list[str]:
    """将关键词列表展开，加入同义词（去重）"""
    expanded = list(keywords)
    for kw in keywords:
        for syn in _SYNONYMS.get(kw, []):
            if syn not in expanded:
                expanded.append(syn)
    return expanded


# ─── 内容检索（语义标签）────────────────────────────────────────────

def search_content(
    scene: str = None,
    content_type: str = None,
    keywords: list[str] = None,
    quality_min: int = None,
    source_file: str = None,
    limit: int = 10,
) -> list[dict]:
    """
    按语义标签检索内容页。

    参数：
        scene        - 场景，如 "售前汇报" / "行业会议"
        content_type - 内容类型，如 "解决方案" / "行业背景" / "产品功能"
        keywords     - 关键词列表，如 ["AI", "诊断"]，模糊匹配
        quality_min  - 最低质量分（1-5），建议传 4
        source_file  - 来源文件名关键词，如 "哈密" / "卫宁"
        limit        - 返回条数上限

    返回：
        list of dict，每项含 src / page / title / summary / scene /
        content_type / keywords / quality_score
    """
    db = get_db_path()
    conn = sqlite3.connect(db)
    conn.row_factory = sqlite3.Row

    conditions = ["1=1"]
    params = []

    if scene:
        conditions.append("t.scene = ?")
        params.append(scene)

    if content_type:
        conditions.append("t.content_type = ?")
        params.append(content_type)

    if quality_min is not None:
        conditions.append("t.quality_score >= ?")
        params.append(quality_min)

    if source_file:
        conditions.append("s.file_name LIKE ?")
        params.append(f"%{source_file}%")

    # keywords：同义词展开后，标题（s.title）和AI标签（t.keywords）分别匹配
    # 标题命中优先级高于标签命中
    title_hit_expr = "0"  # 用于排序：标题命中得1分
    if keywords:
        expanded = _expand_keywords(keywords)
        kw_parts = []
        for kw in expanded:
            kw_parts.append(
                "(t.keywords LIKE ? OR s.title LIKE ?)"
            )
            params += [f"%{kw}%", f"%{kw}%"]
        conditions.append(f"({' OR '.join(kw_parts)})")
        # 标题命中打分表达式（任一展开词命中标题则得1）
        title_cases = " OR ".join([f"s.title LIKE ?" for kw in expanded])
        title_hit_expr = f"CASE WHEN ({title_cases}) THEN 1 ELSE 0 END"
        params_order = [f"%{kw}%" for kw in expanded]
    else:
        params_order = []

    sql = f"""
        SELECT
            s.id, s.file_path, s.file_name, s.slide_index,
            s.title, s.has_image, s.has_chart,
            t.scene, t.content_type, t.industries, t.keywords,
            t.quality_score, t.summary
        FROM slides s
        JOIN tags t ON s.id = t.slide_id
        WHERE {' AND '.join(conditions)}
        ORDER BY {title_hit_expr} DESC, t.quality_score DESC
        LIMIT ?
    """
    params += params_order
    params.append(limit)

    rows = conn.execute(sql, params).fetchall()
    conn.close()

    return [_format_content_row(r) for r in rows]


# ─── 结构检索（按文件名）────────────────────────────────────────────

# 文件名中可识别的 layout 关键词
_LAYOUT_KEYWORDS = ["封面页", "过渡页", "目录页", "结尾页", "二分类", "三分类", "四分类"]

def search_structural(
    layout: str = None,
    background: str = None,
    limit: int = 10,
) -> list[dict]:
    """
    按文件名识别结构型素材（封面/过渡页/目录页等）。

    参数：
        layout     - 版式关键词，如 "过渡页" / "封面页" / "三分类"
        background - 背景色，"浅色底" 或 "深色底"
        limit      - 返回条数上限

    返回：
        list of dict，每项含 src / page / title / layout / background
    """
    db = get_db_path()
    conn = sqlite3.connect(db)
    conn.row_factory = sqlite3.Row

    conditions = ["1=1"]
    params = []

    if layout:
        conditions.append("s.file_name LIKE ?")
        params.append(f"%{layout}%")

    if background:
        conditions.append("s.file_name LIKE ?")
        params.append(f"%{background}%")

    # 只返回文件名中含有已知 layout 关键词的（排除完整版内容页）
    layout_filter = " OR ".join(
        [f"s.file_name LIKE '%{kw}%'" for kw in _LAYOUT_KEYWORDS]
    )
    conditions.append(f"({layout_filter})")

    sql = f"""
        SELECT s.id, s.file_path, s.file_name, s.slide_index, s.title
        FROM slides s
        WHERE {' AND '.join(conditions)}
        ORDER BY s.file_name, s.slide_index
        LIMIT ?
    """
    params.append(limit)

    rows = conn.execute(sql, params).fetchall()
    conn.close()

    return [_format_structural_row(r) for r in rows]


# ─── 格式化工具 ─────────────────────────────────────────────────────

def _format_content_row(row) -> dict:
    kw_raw = row["keywords"]
    try:
        kw = json.loads(kw_raw) if kw_raw else []
    except Exception:
        kw = [kw_raw] if kw_raw else []

    ind_raw = row["industries"]
    try:
        industries = json.loads(ind_raw) if ind_raw else []
    except Exception:
        industries = [ind_raw] if ind_raw else []

    return {
        "id": row["id"],
        "src": row["file_path"],
        "page": row["slide_index"],
        "title": row["title"],
        "file_name": row["file_name"],
        "has_image": bool(row["has_image"]),
        "has_chart": bool(row["has_chart"]),
        "scene": row["scene"],
        "content_type": row["content_type"],
        "industries": industries,
        "keywords": kw,
        "quality_score": row["quality_score"],
        "summary": row["summary"],
    }


def _format_structural_row(row) -> dict:
    fname = row["file_name"]
    layout = next((kw for kw in _LAYOUT_KEYWORDS if kw in fname), None)
    background = "深色底" if "深色底" in fname else ("浅色底" if "浅色底" in fname else None)

    return {
        "id": row["id"],
        "src": row["file_path"],
        "page": row["slide_index"],
        "title": row["title"],
        "file_name": fname,
        "layout": layout,
        "background": background,
    }


# ─── 便捷打印（调试用）──────────────────────────────────────────────

def print_results(results: list[dict], mode: str = "content"):
    if not results:
        print("  （无结果）")
        return

    for i, r in enumerate(results, 1):
        print(f"\n[{i}] {r['file_name']}  第 {r['page']} 页")
        print(f"     标题：{r.get('title') or '（无）'}")

        if mode == "content":
            print(f"     场景：{r.get('scene')}  类型：{r.get('content_type')}  质量：{r.get('quality_score')}")
            kw = r.get("keywords", [])
            if kw:
                print(f"     关键词：{', '.join(kw)}")
            summary = r.get("summary", "")
            if summary:
                print(f"     摘要：{summary[:80]}{'...' if len(summary) > 80 else ''}")
        else:
            print(f"     版式：{r.get('layout')}  背景：{r.get('background')}")


# ─── 快速测试入口 ────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=== 内容检索测试 ===")
    results = search_content(scene="售前汇报", quality_min=4, limit=5)
    print_results(results, mode="content")

    print("\n=== 结构检索测试 ===")
    results = search_structural(layout="过渡页", background="浅色底", limit=5)
    print_results(results, mode="structural")
