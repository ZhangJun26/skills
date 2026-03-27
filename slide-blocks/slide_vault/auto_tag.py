# -*- coding: utf-8 -*-
"""
auto_tag.py - 批量为素材库打标签（规则匹配，无需 API）

用法：
    cd <skill 根目录>
    python tools/auto_tag.py

前置条件：
    1. 已配置 config.yaml（db_path 指向你的 slide_vault.db）
    2. 已运行 scanner.py 完成文字提取
"""
import json
import re
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from slide_vault.config import get_db_path

DB_PATH = get_db_path()


# ── 场景判断 ──────────────────────────────────────────────────────────────────
def get_scene(fn, title, body):
    if "售前汇报" in fn:
        return "售前汇报"
    if "行业会议" in fn:
        return "行业会议"
    if "公司简介" in fn:
        return "公司简介"
    text = title + body
    if any(k in text for k in ["应用场景实践", "AI场景", "标杆案例", "解决方案"]):
        return "售前汇报"
    if any(k in text for k in ["行业会议", "HIT", "AI重塑"]):
        return "行业会议"
    return "其他"


# ── 内容类型判断 ──────────────────────────────────────────────────────────────
def get_content_type(fn, title, body, has_chart):
    cat = extract_category(fn)
    text = (title + " " + body).lower()

    structural = {
        "封面页": "封面/扉页",
        "目录":   "目录",
        "过渡页": "过渡页",
        "结尾页": "结束页",
    }
    for kw, ct in structural.items():
        if kw in cat:
            return ct

    if cat.startswith("ai场景") or "应用场景实践" in text or "ai赋能" in text[:50]:
        return "AI场景"
    if "标杆案例" in cat or any(k in text for k in ["标杆", "落地", "三级甲等", "已在全国", "客户案例"]):
        return "客户案例"
    if has_chart or any(k in text for k in ["同比", "市场规模", "占比", "增长", "亿元", "top10"]):
        return "数据洞察"
    if any(k in text for k in ["政策", "卫健委", "医改", "行业背景", "市场格局", "信息化发展"]):
        return "行业背景"
    if any(k in text for k in ["规划", "战略", "路径", "阶段", "里程碑", "愿景"]):
        return "战略规划"
    if any(k in text for k in ["公司成立", "发展历程", "研发基地", "覆盖全国"]):
        return "公司介绍"
    if any(k in text for k in ["功能", "平台架构", "产品架构", "集成平台", "saas"]):
        return "产品功能"
    if any(k in text for k in ["解决方案", "整体方案", "建设方案", "一体化", "数字化转型"]):
        return "解决方案"
    if has_chart:
        return "数据洞察"
    return "其他"


# ── 行业标签 ──────────────────────────────────────────────────────────────────
def get_industries(fn, title, body):
    text = fn + title + body
    tags = []
    if any(k in text for k in ["医院", "临床", "病历", "患者", "诊疗", "护理", "检验", "影像"]):
        tags.append("医疗健康")
    if any(k in text for k in ["HIS", "EMR", "PACS", "信息化", "数字化医院", "智慧医院"]):
        tags.append("医院信息化")
    if any(k in text for k in ["大模型", "AI", "人工智能", "大语言模型", "CDSS"]):
        tags.append("医疗AI")
    if any(k in text for k in ["基层", "社区卫生", "公共卫生", "卫健委"]):
        tags.append("公共卫生")
    if not tags:
        tags.append("医疗健康")
    return tags[:3]


# ── 关键词提取 ────────────────────────────────────────────────────────────────
def get_keywords(title, body):
    text = title + " " + body
    cn_words = re.findall(r'[\u4e00-\u9fff]{2,8}', text)
    en_words = re.findall(r'[A-Za-z]{4,}', text)
    stopwords = {"的", "了", "和", "是", "在", "等", "与", "以", "对", "为", "将", "从",
                 "通过", "实现", "提供", "支持", "能够", "进行", "建设", "系统", "数据",
                 "内容", "标题"}
    seen, result = set(), []
    for w in cn_words + en_words:
        if w not in stopwords and w not in seen:
            seen.add(w)
            result.append(w)
        if len(result) >= 6:
            break
    return result if result else ["医疗信息化"]


# ── 质量评分 ──────────────────────────────────────────────────────────────────
def get_quality(content_type, title, body, fn):
    if content_type in ("封面/扉页", "结束页"):
        return 1
    if content_type in ("目录", "过渡页"):
        return 2
    text_len = len(title) + len(body)
    if title.strip() in ("XXX", "标题", "", "XX") and len(body) < 10:
        return 2
    if content_type in ("AI场景", "客户案例"):
        return 4 if text_len > 100 else 3
    if text_len > 200:
        return 4
    if text_len > 80:
        return 3
    return 2


# ── 摘要生成 ──────────────────────────────────────────────────────────────────
def get_summary(title, body, content_type):
    if title and title.strip() not in ("XXX", "标题", "", "XX"):
        return title.strip().replace("\n", " ")[:20]
    if body and body.strip() not in ("XXX", "XX"):
        return body.strip().replace("\n", " ")[:20]
    return content_type


# ── 文件名分类提取 ────────────────────────────────────────────────────────────
def extract_category(fn):
    m = re.match(r'单页-(.+?)-[浅深]色底', fn)
    if m:
        return m.group(1).lower()
    return fn.lower()


# ── 主打标签逻辑 ──────────────────────────────────────────────────────────────
def tag_slide(slide):
    fn    = slide["file_name"]
    title = slide["title"] or ""
    body  = slide["body"] or ""
    has_chart = slide["has_chart"]

    scene        = get_scene(fn, title, body)
    content_type = get_content_type(fn, title, body, has_chart)
    industries   = get_industries(fn, title, body)
    keywords     = get_keywords(title, body)
    quality      = get_quality(content_type, title, body, fn)
    summary      = get_summary(title, body, content_type)

    return {
        "scene": scene, "content_type": content_type,
        "industries": industries, "keywords": keywords,
        "quality_score": quality, "summary": summary,
    }


def tag_complete_slide(slide):
    tag = tag_slide(slide)
    if slide["page"] == 1:
        tag["content_type"] = "封面/扉页"
        tag["quality_score"] = 1
    title, body = slide["title"] or "", slide["body"] or ""
    if "目录" in title or ("目录" in body and len(body) < 100):
        tag["content_type"] = "目录"
        tag["quality_score"] = 2
    fn = slide["file_name"]
    if "售前汇报" in fn:
        tag["scene"] = "售前汇报"
    elif "行业会议" in fn:
        tag["scene"] = "行业会议"
    elif "公司简介" in fn:
        tag["scene"] = "公司简介"
    return tag


# ── 写入数据库 ────────────────────────────────────────────────────────────────
def insert_tag(conn, slide_id, tag):
    now = datetime.now().isoformat()
    conn.execute("""
        INSERT OR REPLACE INTO tags
        (slide_id, scene, content_type, industries, keywords, quality_score, summary, tagged_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        slide_id, tag["scene"], tag["content_type"],
        json.dumps(tag["industries"], ensure_ascii=False),
        json.dumps(tag["keywords"], ensure_ascii=False),
        tag["quality_score"], tag["summary"], now,
    ))


# ── 入口 ──────────────────────────────────────────────────────────────────────
def main():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    # 查询未打标签的幻灯片
    rows = conn.execute("""
        SELECT s.id, s.file_name, s.page, s.title, s.body_text as body, s.has_chart
        FROM slides s
        LEFT JOIN tags t ON t.slide_id = s.id
        WHERE t.slide_id IS NULL
    """).fetchall()

    if not rows:
        print("所有幻灯片已打标签，无需处理。")
        conn.close()
        return

    print(f"发现 {len(rows)} 页未打标签，开始处理...")
    ok = 0
    for row in rows:
        slide = dict(row)
        fn = slide["file_name"]
        tag = tag_complete_slide(slide) if ("完整版" in fn or "公司简介" in fn) else tag_slide(slide)
        insert_tag(conn, slide["id"], tag)
        ok += 1
        if ok % 50 == 0:
            conn.commit()
            print(f"  已写入 {ok} 条...")

    conn.commit()
    conn.close()
    print(f"\n完成！共写入 {ok} 条标签。")


if __name__ == "__main__":
    main()
