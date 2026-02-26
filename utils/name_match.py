"""名寄せユーティリティ — 利用者名のファジーマッチ

利用者名がファイル間で表記ブレするため（スペース有無、異体字、漢字違い等）、
居室番号をプライマリキーとしつつ、名前は補助的にファジーマッチで紐付ける。

マッチ優先順位:
  1. 正規化後の完全一致（スペース除去+NFKC）
  2. 姓一致＋名が1文字違い（例: 浩之 vs 浩行）
  3. 部分一致（姓のみ一致等）
"""

import re
import unicodedata


def normalize_name(name: str | None) -> str:
    """名前を正規化する。

    - None/空文字 -> ''
    - 全角/半角スペースを除去
    - NFKC正規化（異体字・全角英数→半角）
    """
    if not name:
        return ""
    # NFKC正規化で異体字・全角文字を統一
    normalized = unicodedata.normalize("NFKC", str(name))
    # 全角/半角スペースを除去
    normalized = re.sub(r"[\s\u3000]+", "", normalized)
    return normalized


def names_match(name1: str | None, name2: str | None) -> bool:
    """2つの名前が同一人物かどうかを判定する。

    居室番号がプライマリキーなので、これは補助的なチェック。
    正規化一致 or 1文字違いまで許容。
    """
    n1 = normalize_name(name1)
    n2 = normalize_name(name2)
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    # 1文字違いまで許容（例: 宮本浩之 vs 宮本浩行）
    return _is_one_char_diff(n1, n2)


def _is_one_char_diff(s1: str, s2: str) -> bool:
    """2つの文字列が1文字だけ異なるかどうか。

    同じ長さの文字列で、1箇所だけ文字が違う場合にTrueを返す。
    名前の漢字違い（例: 浩之 vs 浩行）を検出するための補助関数。
    """
    if len(s1) != len(s2):
        return False
    if len(s1) < 2:
        return False
    diffs = sum(1 for a, b in zip(s1, s2) if a != b)
    return diffs == 1


def find_best_match(target_name: str, candidates: list[str]) -> str | None:
    """候補リストから最も一致する名前を返す。

    マッチ優先順位:
      1. 正規化後の完全一致
      2. 1文字違い（漢字ミス・異体字対応）
      3. 部分一致（姓のみ一致等）

    Args:
        target_name: 検索対象の名前
        candidates: マッチ候補の名前リスト

    Returns:
        最も一致する候補名。見つからなければNone。
    """
    if not target_name:
        return None

    target_norm = normalize_name(target_name)
    if not target_norm:
        return None

    # 正規化後の完全一致
    for c in candidates:
        if normalize_name(c) == target_norm:
            return c

    # 1文字違い（漢字ミス対応: 浩之 vs 浩行 等）
    for c in candidates:
        c_norm = normalize_name(c)
        if c_norm and _is_one_char_diff(target_norm, c_norm):
            return c

    # 部分一致（姓だけ一致等）
    for c in candidates:
        c_norm = normalize_name(c)
        if c_norm and (target_norm in c_norm or c_norm in target_norm):
            return c

    return None
