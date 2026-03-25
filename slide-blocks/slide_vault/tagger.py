"""
tagger.py - AI 打标签模块（使用 Claude API）
"""

import os
import sys
import json
import sqlite3
import time
from datetime import datetime
from pathlib import Path

import anthropic

DB_PATH = Path("D:/Claude/SlideBlocks/slide_vault.db")

PROMPT_TEMPLATE = """你是一位PPT内容分析专家，请分析以下PPT单页内容，返回JSON格式标签。

页面信息：
- 标题：{title}
- 正文：{body_text}
- 含图片：{has_image}
- 含图表：{has_chart}

请返回如下JSON（只返回JSON，不要其他文字）：
{{
  "scene": "售前汇报|行业会议|内部会议|产品介绍|方案报价|培训课件|其他",
  "content_type": "封面/扉页|目录|问题/背景|解决方案|产品功能|数据/成果|案例/客户|团队/公司介绍|行动计划|结束页|过渡页|其他",
  "industries": ["最多3个行业标签，从以下选择：医疗卫生、医院管理、公共卫生、医疗IT、政府/卫健委、保险、通用"],
  "keywords": ["3到8个关键词"],
  "quality_score": 1到5的整数,
  "summary": "一句话描述这页的核心内容，不超过30字"
}}"""


def get_untagged_slides(conn: sqlite3.Connection) -> list[dict]:
    rows = conn.execute("""
        SELECT s.id, s.slide_index, s.title, s.body_text, s.has_image, s.has_chart
        FROM slides s
        LEFT JOIN tags t ON s.id = t.slide_id
        WHERE t.slide_id IS NULL
        ORDER BY s.slide_index
    """).fetchall()
    return [
        {
            "id": r[0],
            "slide_index": r[1],
            "title": r[2] or "",
            "body_text": r[3] or "",
            "has_image": bool(r[4]),
            "has_chart": bool(r[5]),
        }
        for r in rows
    ]


def call_claude(slide: dict, client: anthropic.Anthropic) -> dict | None:
    prompt = PROMPT_TEMPLATE.format(
        title=slide["title"][:300],
        body_text=slide["body_text"][:1000],
        has_image="是" if slide["has_image"] else "否",
        has_chart="是" if slide["has_chart"] else "否",
    )

    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = message.content[0].text.strip()
        # 有时模型会返回 ```json ... ``` 格式，去掉代码块
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"    [!] JSON解析失败：{e}\n    原文：{raw[:200]}")
        return None
    except Exception as e:
        print(f"    [!] API调用失败：{e}")
        return None


def save_tag(conn: sqlite3.Connection, slide_id: int, tag: dict, title: str = ""):
    now = datetime.now().isoformat()
    # 始终将标题加入 keywords，确保按标题文字检索时不漏
    keywords = tag.get("keywords", [])
    if title and title not in keywords:
        keywords = [title] + keywords
    conn.execute("""
        INSERT OR REPLACE INTO tags
        (slide_id, scene, content_type, industries, keywords, quality_score, summary, tagged_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        slide_id,
        tag.get("scene", "其他"),
        tag.get("content_type", "其他"),
        json.dumps(tag.get("industries", []), ensure_ascii=False),
        json.dumps(keywords, ensure_ascii=False),
        tag.get("quality_score", 3),
        tag.get("summary", ""),
        now,
    ))
    conn.commit()


def tag_all_slides(delay: float = 0.3):
    """对数据库中所有未标签的 slides 批量打标签"""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("[错误] 未找到 ANTHROPIC_API_KEY 环境变量")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)
    conn = sqlite3.connect(DB_PATH)

    slides = get_untagged_slides(conn)
    print(f"[打标签] 待处理：{len(slides)} 页\n")

    success = 0
    failed = 0

    for i, slide in enumerate(slides, 1):
        print(f"[{i:02d}/{len(slides)}] 第{slide['slide_index']}页：{slide['title'][:40] or '（无标题）'}")
        tag = call_claude(slide, client)

        if tag:
            save_tag(conn, slide["id"], tag, title=slide["title"])
            print(f"    ✓ {tag.get('content_type')} | {tag.get('scene')} | 评分:{tag.get('quality_score')} | {tag.get('summary','')[:30]}")
            success += 1
        else:
            failed += 1

        if i < len(slides):
            time.sleep(delay)

    conn.close()
    print(f"\n[完成] 成功:{success}  失败:{failed}")


def preview_tags(limit: int = 39):
    """预览打标签结果"""
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("""
        SELECT s.slide_index, s.title, t.scene, t.content_type,
               t.industries, t.keywords, t.quality_score, t.summary
        FROM slides s
        JOIN tags t ON s.id = t.slide_id
        ORDER BY s.slide_index
        LIMIT ?
    """, (limit,)).fetchall()
    conn.close()

    print(f"\n{'='*70}")
    print("标签结果预览")
    print(f"{'='*70}")
    for row in rows:
        idx, title, scene, ctype, industries, keywords, score, summary = row
        ind_list = json.loads(industries or "[]")
        kw_list = json.loads(keywords or "[]")
        print(f"\n【第{idx:02d}页】{title[:35] or '（无标题）'}")
        print(f"  场景:{scene}  类型:{ctype}  评分:{'★'*score}")
        print(f"  行业:{' / '.join(ind_list)}")
        print(f"  关键词:{' · '.join(kw_list)}")
        print(f"  摘要:{summary}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "preview":
        # python tagger.py preview
        preview_tags()
    else:
        tag_all_slides()
        preview_tags()
