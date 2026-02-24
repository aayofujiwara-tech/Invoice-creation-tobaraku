"""共通ヘルパー関数"""

import math
import re
from pathlib import Path

import yaml


def load_config(config_path: str = "config.yaml") -> dict:
    """config.yamlを読み込む"""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def parse_month_arg(month_str: str) -> tuple[int, int]:
    """'2026-01' 形式の文字列を (year, month) タプルに変換"""
    parts = month_str.split("-")
    return int(parts[0]), int(parts[1])


def to_reiwa_label(year: int, month: int) -> str:
    """西暦年月を和暦ラベル(RX.X)に変換。例: (2026, 1) -> 'R8.1'"""
    reiwa_year = year - 2018
    return f"R{reiwa_year}.{month}"


def parse_reiwa_label(label: str) -> tuple[int, int]:
    """和暦ラベル(RX.X)を西暦(year, month)に変換。例: 'R8.1' -> (2026, 1)"""
    m = re.match(r"R(\d+)\.(\d+)", label)
    if not m:
        raise ValueError(f"Invalid Reiwa label: {label}")
    reiwa_year = int(m.group(1))
    month = int(m.group(2))
    return 2018 + reiwa_year, month


def nick_sheet_name(year: int, month: int) -> str:
    """ニック請求のシート名を生成。例: (2026, 1) -> '20261'"""
    return f"{year}{month}"


def normalize_room(room_val) -> str:
    """居室番号を正規化する。
    - 数値(608.0) -> '608'
    - 文字列('2A') -> '2A'
    - None -> ''
    """
    if room_val is None:
        return ""
    if isinstance(room_val, float):
        return str(int(room_val))
    if isinstance(room_val, int):
        return str(room_val)
    return str(room_val).strip()


def col_letter_to_index(letter: str) -> int:
    """列文字をインデックス(1始まり)に変換。例: 'A'->1, 'E'->5, 'AI'->35"""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def calc_nick_billing_price(set_type: str, config: dict) -> int:
    """ニックのセット種別から請求単価（マークアップ後）を計算する。

    Aセットは908円固定。それ以外は ceil(基本単価 × 1.21)。
    """
    overrides = config.get("nick_price_overrides", {})
    if set_type in overrides:
        return overrides[set_type]

    base_prices = config.get("nick_base_prices", {})
    base = base_prices.get(set_type)
    if base is None:
        raise ValueError(f"Unknown nick set type: {set_type}")

    rate = config.get("nick_markup_rate", 1.21)
    return math.ceil(base * rate)
