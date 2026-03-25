"""
config.py - 读取 config.yaml，提供统一的路径配置
"""

from pathlib import Path

# config.yaml 在项目根目录
CONFIG_PATH = Path(__file__).parent.parent / "config.yaml"

_cache = None


def load_config() -> dict:
    global _cache
    if _cache is not None:
        return _cache

    try:
        import yaml
        with open(CONFIG_PATH, encoding="utf-8") as f:
            _cache = yaml.safe_load(f)
    except ImportError:
        # 没装 pyyaml 时，手动解析简单 key: "value" 格式
        _cache = {}
        with open(CONFIG_PATH, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if ":" in line:
                    k, _, v = line.partition(":")
                    v = v.strip().strip('"').strip("'")
                    if v:
                        _cache[k.strip()] = v

    return _cache


def get_db_path() -> Path:
    return Path(load_config()["db_path"])


def get_materials_dir() -> Path:
    return Path(load_config()["materials_dir"])


def get_output_dir() -> Path:
    return Path(load_config()["output_dir"])
