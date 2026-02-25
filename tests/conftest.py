"""テスト共通設定"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from utils.helpers import load_config, find_file_flexible

CONFIG = load_config(str(PROJECT_ROOT / "config.yaml"))
INPUT_BASE = PROJECT_ROOT / CONFIG["input"]["base_dir"]


def resolve_facility_master(facility_name: str) -> Path:
    """拠点の請求マスターファイルパスを柔軟に解決する"""
    fpath = CONFIG["input"]["facilities"][facility_name]
    facility_dir = INPUT_BASE / fpath["dir"]
    return find_file_flexible(facility_dir, fpath["master"])


def resolve_facility_meal(facility_name: str) -> Path:
    """拠点の食費管理表ファイルパスを柔軟に解決する"""
    fpath = CONFIG["input"]["facilities"][facility_name]
    facility_dir = INPUT_BASE / fpath["dir"]
    return find_file_flexible(facility_dir, fpath["meal"])


def resolve_nick_file() -> Path:
    """ニック請求ファイルパスを柔軟に解決する"""
    common_dir = INPUT_BASE / CONFIG["input"]["common_dir"]
    return find_file_flexible(common_dir, CONFIG["input"]["nick_file"])
