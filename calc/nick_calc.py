"""ニック請求計算モジュール

ニック請求から読み取ったNickRecordを基に、入居者ごとのオムツ・日用品月額を計算する。
請求単価 = ceil(基本単価 × 1.21)（Aセットは908円固定）

ニック請求には居室番号がないため、名寄せで請求マスターと紐付けが必要。
"""

from readers.nick_reader import NickRecord, DIAPER_SETS, SUPPLY_SETS
from utils.helpers import calc_nick_billing_price
from utils.name_match import normalize_name


def calc_nick_billing(record: NickRecord, config: dict) -> tuple[int, int]:
    """1利用者のニック請求月額を計算する。

    Args:
        record: NickRecord
        config: config.yaml の内容

    Returns:
        (オムツ合計, 日用品合計) のタプル
    """
    diaper_total = 0
    supply_total = 0

    for s in record.sets:
        price = calc_nick_billing_price(s.set_type, config)
        amount = s.day_count * price

        if s.is_diaper:
            diaper_total += amount
        elif s.is_supply:
            supply_total += amount

    return diaper_total, supply_total


def calc_all_nick(
    nick_records: list[NickRecord],
    room_name_map: dict[str, str],
    config: dict,
) -> dict[str, tuple[int, int]]:
    """全利用者のニック請求を計算し、居室番号→(オムツ, 日用品)のマッピングを返す。

    ニック請求には居室番号がないため、room_name_map（居室→名前）を使って
    名前ベースで逆引きする。

    Args:
        nick_records: NickRecordのリスト
        room_name_map: {居室番号: 利用者名} のマッピング（全拠点分）
        config: config.yaml の内容

    Returns:
        {居室番号: (オムツ合計, 日用品合計)} の辞書
    """
    # 名前→居室番号の逆引きマップを作成
    name_to_room: dict[str, str] = {}
    for room, name in room_name_map.items():
        norm = normalize_name(name)
        if norm:
            name_to_room[norm] = room

    result = {}
    for nr in nick_records:
        diaper, supply = calc_nick_billing(nr, config)
        if diaper == 0 and supply == 0:
            continue

        # 名前で居室番号を逆引き
        nick_name_norm = normalize_name(nr.name)
        room = name_to_room.get(nick_name_norm)

        if room is None:
            # 部分一致で探す
            for norm_name, r in name_to_room.items():
                if nick_name_norm and norm_name and (
                    nick_name_norm in norm_name or norm_name in nick_name_norm
                ):
                    room = r
                    break

        if room is not None:
            # 同一居室に複数エントリがある場合は加算
            if room in result:
                existing = result[room]
                result[room] = (existing[0] + diaper, existing[1] + supply)
            else:
                result[room] = (diaper, supply)

    return result
