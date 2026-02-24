"""名寄せユーティリティ — 利用者名のファジーマッチ"""

import re
import unicodedata


def normalize_name(name: str | None) -> str:
    """名前を正規化する。
    - None/空文字 -> ''
    - 全角/半角スペースを除去
    - NFKC正規化（異体字・全角英数→半角）
    - 空白除去
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
    完全一致ではなく、正規化後の一致で判定。
    """
    n1 = normalize_name(name1)
    n2 = normalize_name(name2)
    if not n1 or not n2:
        return False
    return n1 == n2


def find_best_match(target_name: str, candidates: list[str]) -> str | None:
    """候補リストから最も一致する名前を返す。

    完全一致 → 正規化一致 → 部分一致 の順で探す。
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

    # 部分一致（姓だけ一致等）
    for c in candidates:
        c_norm = normalize_name(c)
        if c_norm and (target_norm in c_norm or c_norm in target_norm):
            return c

    return None
