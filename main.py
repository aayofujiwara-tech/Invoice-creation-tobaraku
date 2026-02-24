"""AA Group ええすまい請求自動化 メインスクリプト

食費管理表 ＋ ニック請求カレンダー → 請求集約シート(RX.X) → 請求書・領収書・合算明細書

Usage:
    python main.py --month 2026-01 --input-dir ./参考 --output-dir ./output
    python main.py --month 2026-01  # デフォルト: input=./参考, output=./output
"""

import argparse
import shutil
import sys
from datetime import date
from pathlib import Path

from utils.helpers import load_config, parse_month_arg, to_reiwa_label
from readers.meal_reader import read_all_facilities_meals
from readers.nick_reader import read_nick_file
from readers.master_reader import read_all_facilities_masters
from calc.meal_calc import calc_all_meals
from calc.nick_calc import calc_all_nick
from calc.billing import build_all_billings
from writers.summary_writer import write_summary_to_file
from writers.invoice_writer import write_all_invoices
from writers.receipt_writer import write_all_receipts
from writers.combined_writer import write_all_combined


def main():
    parser = argparse.ArgumentParser(
        description="ええすまい請求自動化システム",
    )
    parser.add_argument(
        "--month", required=True,
        help="対象年月 (YYYY-MM形式、例: 2026-01)",
    )
    parser.add_argument(
        "--input-dir", default="./参考",
        help="入力データディレクトリ (default: ./参考)",
    )
    parser.add_argument(
        "--output-dir", default="./output",
        help="出力ディレクトリ (default: ./output)",
    )
    parser.add_argument(
        "--issue-date", default=None,
        help="発行日 (YYYY-MM-DD形式、省略時は翌月3日)",
    )
    parser.add_argument(
        "--config", default="config.yaml",
        help="設定ファイルパス (default: config.yaml)",
    )
    parser.add_argument(
        "--skip-invoices", action="store_true",
        help="請求書・領収書・合算明細書の生成をスキップ",
    )

    args = parser.parse_args()

    # --- パラメータ解析 ---
    year, month = parse_month_arg(args.month)
    reiwa_label = to_reiwa_label(year, month)
    input_dir = Path(args.input_dir)
    output_base = Path(args.output_dir) / f"{year}-{month:02d}"
    config = load_config(args.config)

    # 発行日: 指定がなければ翌月3日
    if args.issue_date:
        issue_date = date.fromisoformat(args.issue_date)
    else:
        next_month = month + 1 if month < 12 else 1
        next_year = year if month < 12 else year + 1
        issue_date = date(next_year, next_month, 3)

    print(f"=== ええすまい請求自動化 ===")
    print(f"対象月: {reiwa_label} ({year}年{month}月)")
    print(f"入力: {input_dir}")
    print(f"出力: {output_base}")
    print(f"発行日: {issue_date}")
    print()

    # --- Phase 1: 読取 ---
    print("--- Phase 1: データ読取 ---")

    # 食費管理表
    print("食費管理表を読み取り中...")
    try:
        meals_by_facility = read_all_facilities_meals(input_dir, config, year, month)
        for fname, records in meals_by_facility.items():
            active = [r for r in records if not r.is_empty]
            print(f"  {fname}: {len(records)}スロット, {len(active)}名の食事データ")
    except Exception as e:
        print(f"  食費管理表の読取エラー: {e}")
        meals_by_facility = {}

    # ニック請求
    print("ニック請求を読み取り中...")
    nick_file = input_dir / config["nick_file"]
    try:
        nick_records = read_nick_file(nick_file, year, month)
        print(f"  {len(nick_records)}名のニックデータ")
    except Exception as e:
        print(f"  ニック請求の読取エラー: {e}")
        nick_records = []

    # 請求マスター
    print("請求マスターを読み取り中...")
    masters_by_facility = read_all_facilities_masters(input_dir, config, year, month)
    for fname, (residents, rx_rows) in masters_by_facility.items():
        active = [r for r in rx_rows if not r.is_vacant]
        print(f"  {fname}: {len(residents)}入居者マスタ, {len(active)}名のRXデータ")

    print()

    # --- Phase 2: 計算 ---
    print("--- Phase 2: 計算 ---")

    # 拠点ごとに食費計算・ニック計算→請求集約
    # ※居室番号が拠点間で重複するため（例: 703がセレーネ/パシフィック両方に存在）、
    #   ニック計算は拠点ごとに別のroom_name_mapを使って実行する
    all_billings: dict[str, list] = {}

    for fname, (residents, rx_rows) in masters_by_facility.items():
        # 食費計算
        meal_records = meals_by_facility.get(fname, [])
        meal_by_room = calc_all_meals(meal_records, config)
        print(f"  {fname}: 食費計算完了 ({len(meal_by_room)}名)")

        # ニック計算（拠点固有のroom_name_mapで名寄せ）
        room_name_map = {rx.room: rx.name for rx in rx_rows if rx.room and rx.name}
        nick_by_room = calc_all_nick(nick_records, room_name_map, config)
        print(f"  {fname}: ニック計算完了 ({len(nick_by_room)}名)")

        # 請求集約
        billings = build_all_billings(rx_rows, meal_by_room, nick_by_room, config)
        all_billings[fname] = billings

        # 結果サマリー
        active = [b for b in billings if not b.is_vacant]
        total_amount = sum(b.total for b in active)
        adjusted = [b for b in active if b.adjustment != 0]
        print(f"  {fname}: 請求集約完了 ({len(active)}名, 合計{total_amount:,}円, 調整{len(adjusted)}名)")

    print()

    # --- Phase 3: 出力 ---
    print("--- Phase 3: 出力生成 ---")

    for fname, billings in all_billings.items():
        fconf = config["facilities"][fname]
        display_name = fconf["display_name"]
        facility_output = output_base / fname

        # RX.Xシート更新
        prefix = fconf["file_prefix"]
        master_candidates = list(input_dir.glob(f"{prefix}*"))
        if master_candidates:
            master_src = master_candidates[0]
            master_dst = facility_output / master_src.name
            master_dst.parent.mkdir(parents=True, exist_ok=True)

            # 元ファイルをコピーしてからRX.Xシートを追加/更新
            shutil.copy2(str(master_src), str(master_dst))
            write_summary_to_file(master_dst, billings, year, month, master_dst)
            print(f"  {fname}: RX.Xシート更新 → {master_dst}")

        if not args.skip_invoices:
            # 請求書生成
            invoice_files = write_all_invoices(
                billings, fname, display_name, year, month, issue_date, facility_output
            )
            print(f"  {fname}: 請求書{len(invoice_files)}件生成")

            # 領収書生成
            receipt_files = write_all_receipts(
                billings, year, month, issue_date, facility_output
            )
            print(f"  {fname}: 領収書{len(receipt_files)}件生成")

            # 合算明細書生成
            combined_files = write_all_combined(
                billings, display_name, year, month, issue_date, facility_output
            )
            print(f"  {fname}: 合算明細書{len(combined_files)}件生成")

    print()
    print(f"=== 完了: {output_base} ===")


if __name__ == "__main__":
    main()
