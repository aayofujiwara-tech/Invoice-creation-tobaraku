"""食費計算モジュール

食費管理表から読み取ったMealRecordを基に、入居者ごとの食費月額を計算する。
単価: 朝食330円、昼食550円、夕食550円（config.yamlで設定可能）
"""

from readers.meal_reader import MealRecord


def calc_meal_billing(record: MealRecord, config: dict) -> int:
    """1入居者の食費月額を計算する。

    Args:
        record: MealRecord（食費管理表から読み取ったデータ）
        config: config.yaml の内容

    Returns:
        食費月額（円）
    """
    prices = config.get("meal_prices", {})
    breakfast_price = prices.get("breakfast", 330)
    lunch_price = prices.get("lunch", 550)
    dinner_price = prices.get("dinner", 550)
    return record.calc_billing(breakfast_price, lunch_price, dinner_price)


def calc_all_meals(
    records: list[MealRecord], config: dict
) -> dict[str, int]:
    """全入居者の食費を計算し、居室番号→食費月額のマッピングを返す。

    Args:
        records: MealRecordのリスト
        config: config.yaml の内容

    Returns:
        {居室番号: 食費月額} の辞書（空室・食事なしは含まない）
    """
    result = {}
    for r in records:
        if r.is_empty or not r.room:
            continue
        result[r.room] = calc_meal_billing(r, config)
    return result
