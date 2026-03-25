# -*- coding: utf-8 -*-
"""
setup_paths.py — 第一次使用前运行一次

作用：把数据库里存的旧路径前缀，替换成你在 config.yaml 里填写的素材路径。
运行方式：python setup_paths.py
"""

import sqlite3
from pathlib import Path

# ── 读取 config.yaml ──────────────────────────────────────────────────────────
try:
    import yaml
except ImportError:
    print("[错误] 请先安装依赖：pip install pyyaml")
    raise

_HERE = Path(__file__).parent
config = yaml.safe_load((_HERE / "config.yaml").read_text(encoding="utf-8"))

materials_dir = config.get("materials_dir", "").strip()
db_path = _HERE / config.get("db_path", "slide_vault.db")

if not materials_dir:
    print("[错误] 请先在 config.yaml 里填写 materials_dir，再运行此脚本。")
    exit(1)

# ── 数据库里原始的路径前缀（素材库维护者机器上的路径）────────────────────────
OLD_PREFIX = "D:/Claude/SlideMatrix/素材"

# ── 替换路径 ──────────────────────────────────────────────────────────────────
new_base = materials_dir.rstrip("/\\").replace("\\", "/")

conn = sqlite3.connect(str(db_path))
rows = conn.execute("SELECT id, file_path FROM slides").fetchall()

updated = 0
for row_id, old_path in rows:
    normalized = old_path.replace("\\", "/")
    if normalized.startswith(OLD_PREFIX):
        new_path = new_base + normalized[len(OLD_PREFIX):]
        conn.execute("UPDATE slides SET file_path=? WHERE id=?", (new_path, row_id))
        updated += 1

conn.commit()
conn.close()

print(f"[完成] 已更新 {updated} 条路径")
print(f"  旧前缀：{OLD_PREFIX}")
print(f"  新前缀：{new_base}")
print()
print("现在可以正常使用 SlideBlocks 了。")
