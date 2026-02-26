"""包括的診断テスト — 全拠点・全入居者の計算結果をマスターデータと突合

Step 1: 異常データの特定とパターン分析
  - 食費計算値 vs マスターI列
  - ニック計算値 vs マスターK,L列
  - 調整額計算値 vs マスターJ列
  - 小計(W), 合計(Z) の一致性
  - 福祉上限パース結果の検証
  - 名寄せ失敗パターンの検出
  - 数値型変換ミスの検出
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from utils.helpers import load_config, find_file_flexible, calc_nick_billing_price
from utils.name_match import normalize_name
from readers.meal_reader import read_meal_file
from readers.nick_reader import read_nick_file
from readers.master_reader import read_master_file, RxRow, _safe_int, _is_retired_room
from calc.meal_calc import calc_all_meals
from calc.nick_calc import calc_all_nick
from calc.billing import build_all_billings, build_billing, merge_residents_into_rx, _val

CONFIG = load_config(str(PROJECT_ROOT / "config.yaml"))
INPUT_BASE = PROJECT_ROOT / CONFIG["input"]["base_dir"]
YEAR, MONTH = 2026, 1


def load_all_facilities():
    """全拠点のデータを読み込む"""
    results = {}
    for fname in CONFIG["facilities"]:
        fpath_conf = CONFIG["input"]["facilities"][fname]
        fconf = CONFIG["facilities"][fname]
        facility_dir = INPUT_BASE / fpath_conf["dir"]

        master_path = find_file_flexible(facility_dir, fpath_conf["master"])
        meal_path = find_file_flexible(facility_dir, fpath_conf["meal"])

        residents, rx_rows, is_new = read_master_file(master_path, fconf, YEAR, MONTH)
        try:
            meal_records = read_meal_file(meal_path, YEAR, MONTH)
        except ValueError:
            meal_records = []

        results[fname] = {
            "residents": residents,
            "rx_rows": rx_rows,
            "is_new": is_new,
            "meal_records": meal_records,
        }
    return results


def load_nick():
    """ニック請求データを読み込む"""
    common_dir = INPUT_BASE / CONFIG["input"]["common_dir"]
    nick_path = find_file_flexible(common_dir, CONFIG["input"]["nick_file"])
    return read_nick_file(nick_path, YEAR, MONTH)


def run_full_diagnostic():
    """全拠点の全入居者を診断し、不整合パターンを分類出力する"""
    all_data = load_all_facilities()
    nick_records = load_nick()

    print("=" * 80)
    print("包括的診断レポート: 全拠点×全入居者 計算結果 vs マスターデータ突合")
    print("=" * 80)
    print()

    total_residents = 0
    total_discrepancies = 0
    bug_categories = {
        "welfare_positive_adjustment": [],  # 福祉調整額が正になるバグ
        "meal_mismatch": [],                # 食費不一致
        "nick_diaper_mismatch": [],         # ニックオムツ不一致
        "nick_supply_mismatch": [],         # ニック日用品不一致
        "nick_unmatched": [],               # ニック名寄せ失敗
        "adjustment_mismatch": [],          # 調整額不一致
        "subtotal_mismatch": [],            # 小計不一致
        "total_mismatch": [],               # 合計不一致
        "welfare_limit_parse_error": [],    # 福祉上限パースエラー
        "numeric_type_issue": [],           # 数値型変換ミス
        "name_match_ambiguity": [],         # 名前一致の曖昧性
        "vacant_detection_inconsistency": [], # 空室検知の不整合
        "installment_calc_error": [],       # 分割計算エラー
    }

    for fname, data in all_data.items():
        rx_rows = data["rx_rows"]
        meal_records = data["meal_records"]

        print(f"\n{'─' * 60}")
        print(f"拠点: {fname}")
        print(f"{'─' * 60}")

        # 食費計算
        meal_by_room = calc_all_meals(meal_records, CONFIG)

        # ニック計算（拠点ごとのroom_name_map）
        rx_merged = merge_residents_into_rx(data["residents"], rx_rows, CONFIG)
        room_name_map = {rx.room: rx.name for rx in rx_merged if rx.room and rx.name}
        nick_by_room = calc_all_nick(nick_records, room_name_map, CONFIG)

        # 請求集約
        billings = build_all_billings(
            rx_merged, meal_by_room, nick_by_room, CONFIG,
            is_new_month=data["is_new"],
        )
        billing_by_room = {b.room: b for b in billings}

        # マスターデータのRX.X行を辞書化
        master_by_room = {rx.room: rx for rx in rx_rows if rx.room}

        for rx in rx_rows:
            if not rx.room or not rx.name:
                continue
            total_residents += 1

            b = billing_by_room.get(rx.room)
            if not b:
                print(f"  ⚠ {rx.room} {rx.name}: BillingResult見つからず")
                total_discrepancies += 1
                continue

            discrepancies = []

            # --- 1. 食費突合 ---
            calc_meal = meal_by_room.get(rx.room)
            master_meal = rx.meal
            bill_meal = b.meal
            if calc_meal is not None and master_meal is not None:
                if calc_meal != master_meal:
                    discrepancies.append(
                        f"食費: 計算={calc_meal}, マスター={master_meal}, 差={calc_meal - master_meal}"
                    )
                    bug_categories["meal_mismatch"].append(
                        (fname, rx.room, rx.name, calc_meal, master_meal)
                    )

            # --- 2. ニックオムツ突合 ---
            nick_data = nick_by_room.get(rx.room)
            calc_diaper = nick_data[0] if nick_data else None
            master_diaper = rx.diaper
            if calc_diaper is not None and master_diaper is not None and master_diaper > 0:
                if calc_diaper != master_diaper:
                    discrepancies.append(
                        f"オムツ: 計算={calc_diaper}, マスター={master_diaper}, 差={calc_diaper - master_diaper}"
                    )
                    bug_categories["nick_diaper_mismatch"].append(
                        (fname, rx.room, rx.name, calc_diaper, master_diaper)
                    )

            # --- 3. ニック日用品突合 ---
            calc_supply = nick_data[1] if nick_data else None
            master_supply = rx.daily_supplies
            if calc_supply is not None and master_supply is not None and master_supply > 0:
                if calc_supply != master_supply:
                    discrepancies.append(
                        f"日用品: 計算={calc_supply}, マスター={master_supply}, 差={calc_supply - master_supply}"
                    )
                    bug_categories["nick_supply_mismatch"].append(
                        (fname, rx.room, rx.name, calc_supply, master_supply)
                    )

            # --- 4. 調整額突合 ---
            master_adj = rx.adjustment or 0
            calc_adj = b.adjustment
            if master_adj != calc_adj:
                discrepancies.append(
                    f"調整額: 計算={calc_adj}, マスター={master_adj}, 差={calc_adj - master_adj}"
                )
                bug_categories["adjustment_mismatch"].append(
                    (fname, rx.room, rx.name, calc_adj, master_adj, b.welfare_limit)
                )
                # 正の調整額はバグパターン
                if calc_adj > 0:
                    bug_categories["welfare_positive_adjustment"].append(
                        (fname, rx.room, rx.name, calc_adj, b.welfare_limit, b.subtotal)
                    )

            # --- 5. 小計(W)突合 ---
            master_subtotal = rx.subtotal or 0
            if master_subtotal != 0 and b.subtotal != master_subtotal:
                discrepancies.append(
                    f"小計(W): 計算={b.subtotal}, マスター={master_subtotal}, 差={b.subtotal - master_subtotal}"
                )
                bug_categories["subtotal_mismatch"].append(
                    (fname, rx.room, rx.name, b.subtotal, master_subtotal)
                )

            # --- 6. 合計(Z)突合 ---
            master_total = rx.total or 0
            if master_total != 0 and b.total != master_total:
                discrepancies.append(
                    f"合計(Z): 計算={b.total}, マスター={master_total}, 差={b.total - master_total}"
                )
                bug_categories["total_mismatch"].append(
                    (fname, rx.room, rx.name, b.total, master_total)
                )

            # --- 7. 福祉上限パース検証 ---
            if rx.notes:
                parsed_limit = rx.welfare_limit
                if parsed_limit is not None:
                    # 上限がマスター合計と矛盾していないかチェック
                    if master_total > 0 and parsed_limit > 0:
                        if master_total > parsed_limit and master_adj is not None and master_adj >= 0:
                            # 合計 > 上限 なのに調整額が非負 → パースミスの可能性
                            bug_categories["welfare_limit_parse_error"].append(
                                (fname, rx.room, rx.name, rx.notes, parsed_limit, master_total)
                            )

            # --- 8. is_vacant判定の不整合 ---
            rx_vacant = rx.is_vacant
            bill_vacant = b.is_vacant
            if rx_vacant != bill_vacant:
                bug_categories["vacant_detection_inconsistency"].append(
                    (fname, rx.room, rx.name, rx_vacant, bill_vacant,
                     f"rx.total={rx.total}, b.name='{b.name}'")
                )

            # --- 9. 分割計算チェック ---
            if rx.installment_balance and rx.installment_balance > 0:
                expected_remaining = rx.installment_balance - _val(rx.installment)
                actual_remaining = rx.remaining_balance
                if actual_remaining is not None and actual_remaining != expected_remaining:
                    bug_categories["installment_calc_error"].append(
                        (fname, rx.room, rx.name, rx.installment_balance,
                         rx.installment, expected_remaining, actual_remaining)
                    )

            if discrepancies:
                total_discrepancies += 1
                print(f"\n  ✗ {rx.room} {rx.name} — {len(discrepancies)}件の不整合:")
                for d in discrepancies:
                    print(f"      {d}")
            else:
                print(f"  ✓ {rx.room} {rx.name}")

    # --- ニック名寄せ失敗パターンの検出 ---
    print(f"\n{'─' * 60}")
    print("ニック名寄せ分析")
    print(f"{'─' * 60}")

    # 全拠点のroom_name_mapを統合
    all_room_names = {}
    for fname, data in all_data.items():
        for rx in data["rx_rows"]:
            if rx.room and rx.name:
                all_room_names[f"{fname}:{rx.room}"] = rx.name

    for nr in nick_records:
        matched_facilities = []
        for fname, data in all_data.items():
            room_map = {rx.room: rx.name for rx in data["rx_rows"] if rx.room and rx.name}
            # 名寄せ試行
            nick_norm = normalize_name(nr.name)
            for room, name in room_map.items():
                master_norm = normalize_name(name)
                if nick_norm == master_norm or (nick_norm and master_norm and (
                    nick_norm in master_norm or master_norm in nick_norm)):
                    matched_facilities.append((fname, room, name))

        if len(matched_facilities) == 0:
            print(f"  ✗ 名寄せ失敗: {nr.name} (ニック) → マッチなし")
            bug_categories["nick_unmatched"].append((nr.name, nr.station, nr.user_id))
        elif len(matched_facilities) > 1:
            print(f"  ⚠ 多重マッチ: {nr.name} (ニック) → {matched_facilities}")
            bug_categories["name_match_ambiguity"].append(
                (nr.name, matched_facilities)
            )
        else:
            f, r, n = matched_facilities[0]
            print(f"  ✓ {nr.name} → {f}:{r} ({n})")

    # --- サマリー ---
    print(f"\n{'=' * 80}")
    print("診断サマリー")
    print(f"{'=' * 80}")
    print(f"対象入居者数: {total_residents}")
    print(f"不整合検出数: {total_discrepancies}")
    print(f"正常率: {(total_residents - total_discrepancies) / total_residents * 100:.1f}%")
    print()

    for category, items in bug_categories.items():
        if items:
            print(f"\n【{category}】{len(items)}件")
            for item in items:
                print(f"  - {item}")

    return bug_categories, total_residents, total_discrepancies


if __name__ == "__main__":
    bugs, total, discrepancies = run_full_diagnostic()
    print(f"\n\n{'=' * 80}")
    if discrepancies > 0:
        print(f"結果: {discrepancies}/{total}件に不整合 ({discrepancies/total*100:.1f}%)")
        sys.exit(1)
    else:
        print("結果: 全件一致 — 不整合なし")
        sys.exit(0)
