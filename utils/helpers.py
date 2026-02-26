"""共通ヘルパー関数"""

import math
import re
import unicodedata
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


def get_sheet_name_candidates(year: int, month: int) -> list[str]:
    """対象月のシート名候補を返す（和暦・西暦両対応）。

    Returns:
        候補リスト。例: ['R8.1', '202601', '20261']
    """
    reiwa = to_reiwa_label(year, month)       # R8.1
    yyyymm = f"{year}{month:02d}"             # 202601
    yyyym = f"{year}{month}"                  # 20261
    candidates = [reiwa, yyyymm]
    if yyyym != yyyymm:
        candidates.append(yyyym)              # 20261 (ゼロなし)
    return candidates


def find_matching_sheet(sheetnames: list[str], year: int, month: int) -> str | None:
    """シート名リストから対象月に一致するシートを探す。

    和暦(R8.1)・西暦(202601)・ゼロなし(20261)の順で検索。

    Returns:
        見つかったシート名。なければNone。
    """
    candidates = get_sheet_name_candidates(year, month)
    for candidate in candidates:
        if candidate in sheetnames:
            return candidate
    return None


def parse_sheet_year_month(sheet_name: str) -> tuple[int, int] | None:
    """シート名から(西暦年, 月)を抽出する。和暦・西暦両対応。

    Returns:
        (year, month) or None（解析不能の場合）
    """
    # 和暦パターン: R8.1, R7.12
    m = re.match(r"R(\d+)\.(\d+)$", sheet_name)
    if m:
        return 2018 + int(m.group(1)), int(m.group(2))

    # 西暦6桁パターン: 202601
    m = re.match(r"^(20\d{2})(\d{2})$", sheet_name)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return y, mo

    # 西暦5桁パターン: 20261 (ゼロなし1〜9月)
    m = re.match(r"^(20\d{2})(\d)$", sheet_name)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 9:
            return y, mo

    return None


def normalize_room(room_val) -> str:
    """居室番号を正規化する。
    - 数値(608.0) -> '608'
    - 文字列('2A') -> '2A'
    - 全角数字('５B') -> '5B' (NFKC正規化)
    - None -> ''
    """
    if room_val is None:
        return ""
    if isinstance(room_val, float):
        return str(int(room_val))
    if isinstance(room_val, int):
        return str(room_val)
    # NFKC正規化で全角英数字を半角に変換
    return unicodedata.normalize("NFKC", str(room_val)).strip()


def _normalize_filename(name: str) -> str:
    """ファイル名を正規化（比較用）。閉じ括弧と拡張子の間の_やスペースを除去。"""
    # ')_.xlsx' / ') .xlsx' / ')  .xlsx' 等 → ').xlsx' に統一
    return re.sub(r"\)[_\s]+\.", ").", name)


def find_file_flexible(directory: Path, config_filename: str) -> Path:
    """ファイル名のバリエーション（_あり/なし/スペース）を許容して検索する。

    検索順:
      1. config記載のファイル名で完全一致
      2. ディレクトリ内を走査し、正規化後の名前が一致するファイル
    """
    exact = directory / config_filename
    if exact.exists():
        return exact

    target_normalized = _normalize_filename(config_filename)

    for f in directory.iterdir():
        if f.is_file() and _normalize_filename(f.name) == target_normalized:
            return f

    raise FileNotFoundError(
        f"ファイルが見つかりません: {config_filename}（または類似名）in {directory}"
    )


def col_letter_to_index(letter: str) -> int:
    """列文字をインデックス(1始まり)に変換。例: 'A'->1, 'E'->5, 'AI'->35"""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def calc_nick_billing_price(set_type: str, config: dict) -> int:
    """ニックのセット種別から請求単価（マークアップ後）を計算する。

    マスターデータとの突合により、端数処理は四捨五入（round）が正しいと判明。
    例: 用セット base=120, 120×1.21=145.2 → round=145（マスター一致）
        福セット base=60,  60×1.21=72.6  → round=73（マスター一致）
    Aセットはnick_price_overridesで908円固定（マークアップ率が異なるため個別指定）。

    Args:
        set_type: セット種別（全角文字: Ａ,Ｂ,Ｃ,Ｄ,Ｅ,Ｆ,福,ふ,用 等）
        config: config.yaml の内容

    Returns:
        請求単価（円/日、税込）

    Raises:
        ValueError: 未知のセット種別の場合
    """
    overrides = config.get("nick_price_overrides", {})
    if set_type in overrides:
        return overrides[set_type]

    base_prices = config.get("nick_base_prices", {})
    base = base_prices.get(set_type)
    if base is None:
        raise ValueError(
            f"未知のニックセット種別: '{set_type}'\n"
            f"  確認事項: config.yaml の nick_base_prices に '{set_type}' の単価を追加してください。\n"
            f"  登録済みセット: {list(base_prices.keys())}"
        )

    rate = config.get("nick_markup_rate", 1.21)
    # 四捨五入（日本式: 0.5は切り上げ）— Python標準のround()は銀行丸め(偶数丸め)のため
    # math.floorで明示的に実装
    return math.floor(base * rate + 0.5)
