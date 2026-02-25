"""請求集約・調整額計算モジュール

食費・ニック計算結果と請求マスターのRX.Xデータを統合し、
調整額を計算して最終的な請求データを生成する。

調整額ロジック:
  - 備考欄(AA列)に上限額が記載された福祉受給者の場合、
    合計(Z列)が上限額に収まるよう調整額(J列)で逆算調整する。
  - 調整額 = 必要な小計(W) - 調整前小計
  - 必要な小計(W) = 上限額(Z) - 介護負担(X) - 看護負担(Y)
"""

from dataclasses import dataclass, field
from readers.master_reader import ResidentMaster, RxRow


@dataclass
class BillingResult:
    """1入居者の最終請求データ"""
    room: str
    name: str
    # 固定費（来月分前払い）
    rent: int = 0             # D: 家賃
    management: int = 0       # E: 管理費
    common: int = 0           # F: 共益費
    water: int = 0            # G: 水道料金
    utility: int = 0          # H: 定額光熱費
    # 当月実費（自動計算）
    meal: int = 0             # I: 食事
    adjustment: int = 0       # J: 調整額
    diaper: int = 0           # K: オムツ
    daily_supplies: int = 0   # L: 日用品
    # 分割
    installment_balance: int = 0  # C: 分納残（前月から引継ぎ）
    installment: int = 0      # M: 分割支払い
    # 手入力項目（マスターから引き継ぎ）
    office_fee: int = 0       # N: 事務手数料
    day_service: int = 0      # O: デイサービス
    welfare_equip: int = 0    # P: 福祉用具
    pharmacy: int = 0         # Q: しろくま薬局
    doctor: int = 0           # R: 往診診療費
    support: int = 0          # S: サポート費
    other: int = 0            # T: その他
    # 合計
    subtotal: int = 0         # W: 計
    care_burden: int = 0      # X: 一部負担(介護)
    nurse_burden: int = 0     # Y: 一部負担(看護)
    total: int = 0            # Z: 合計
    # メタ
    notes: str = ""           # AA: 備考
    welfare_limit: int | None = None  # 福祉上限額
    remaining_balance: int = 0  # AC: 分割残高

    @property
    def is_vacant(self) -> bool:
        return not self.name

    @property
    def fixed_total(self) -> int:
        return self.rent + self.management + self.common + self.water + self.utility


def _val(v: int | None) -> int:
    return v if v is not None else 0


def build_billing(
    rx_row: RxRow,
    meal_amount: int | None,
    diaper_amount: int | None,
    supply_amount: int | None,
    config: dict,
    is_new_month: bool = False,
) -> BillingResult:
    """1入居者の請求データを構築する。

    食費・ニック計算結果で自動列を上書きし、
    手入力列はマスターの値をそのまま引き継ぐ。
    調整額は福祉上限がある場合に自動計算する。

    Args:
        rx_row: 請求マスターのRX.X行データ
        meal_amount: 食費月額（食費管理表から計算、Noneなら既存値を使用）
        diaper_amount: オムツ月額（ニック請求から計算、Noneなら既存値を使用）
        supply_amount: 日用品月額（ニック請求から計算、Noneなら既存値を使用）
        config: config.yaml の内容
        is_new_month: Trueなら前月データからの新規作成（手入力列はクリア）

    Returns:
        BillingResult
    """
    # 新規月の場合: 分納残を前月の分割残高（C - M）に更新
    if is_new_month:
        new_installment_balance = _val(rx_row.installment_balance) - _val(rx_row.installment)
        if new_installment_balance < 0:
            new_installment_balance = 0
    else:
        new_installment_balance = _val(rx_row.installment_balance)

    # 新規月の場合: 分割支払いは分納残が残っていれば継続
    if is_new_month:
        installment_monthly = config.get("installment_monthly", 10000)
        new_installment = installment_monthly if new_installment_balance > 0 else 0
    else:
        new_installment = _val(rx_row.installment)

    result = BillingResult(
        room=rx_row.room,
        name=rx_row.name,
        # 固定費はマスターから引き継ぎ
        rent=_val(rx_row.rent),
        management=_val(rx_row.management),
        common=_val(rx_row.common),
        water=_val(rx_row.water),
        utility=_val(rx_row.utility),
        # 自動計算列: 計算結果があれば上書き、なければマスターの値を使用
        meal=meal_amount if meal_amount is not None else _val(rx_row.meal),
        diaper=diaper_amount if diaper_amount is not None else _val(rx_row.diaper),
        daily_supplies=supply_amount if supply_amount is not None else _val(rx_row.daily_supplies),
        # 分割
        installment_balance=new_installment_balance,
        installment=new_installment,
        # 手入力項目: 新規月はクリア、既存月はそのまま引き継ぎ
        office_fee=0 if is_new_month else _val(rx_row.office_fee),
        day_service=0 if is_new_month else _val(rx_row.day_service),
        welfare_equip=0 if is_new_month else _val(rx_row.welfare_equip),
        pharmacy=0 if is_new_month else _val(rx_row.pharmacy),
        doctor=0 if is_new_month else _val(rx_row.doctor),
        support=0 if is_new_month else _val(rx_row.support),
        other=0 if is_new_month else _val(rx_row.other),
        # 介護・看護負担はそのまま引き継ぎ
        care_burden=_val(rx_row.care_burden),
        nurse_burden=_val(rx_row.nurse_burden),
        # 備考
        notes=rx_row.notes or "",
    )

    # 福祉上限額の取得
    result.welfare_limit = rx_row.welfare_limit
    if result.welfare_limit is None:
        # configのデフォルト上限は使わない（上限がない人に適用してしまうため）
        pass

    # 調整額の計算
    if result.welfare_limit is not None and result.welfare_limit > 0:
        # 福祉受給者: 合計を上限に収める
        target_total = result.welfare_limit
        needed_subtotal = target_total - result.care_burden - result.nurse_burden

        # 調整前の小計（J列を除くD〜T列の合計）
        pre_adjustment_subtotal = (
            result.fixed_total
            + result.meal
            + result.diaper
            + result.daily_supplies
            + result.installment
            + result.office_fee
            + result.day_service
            + result.welfare_equip
            + result.pharmacy
            + result.doctor
            + result.support
            + result.other
        )

        result.adjustment = needed_subtotal - pre_adjustment_subtotal
        result.subtotal = needed_subtotal
        result.total = target_total
    else:
        # 一般入居者: 調整額なし
        result.adjustment = 0
        result.subtotal = (
            result.fixed_total
            + result.meal
            + result.adjustment
            + result.diaper
            + result.daily_supplies
            + result.installment
            + result.office_fee
            + result.day_service
            + result.welfare_equip
            + result.pharmacy
            + result.doctor
            + result.support
            + result.other
        )
        result.total = result.subtotal + result.care_burden + result.nurse_burden

    # 分割残高
    result.remaining_balance = result.installment_balance - result.installment

    return result


def merge_residents_into_rx(
    residents: list[ResidentMaster],
    rx_rows: list[RxRow],
    config: dict,
) -> list[RxRow]:
    """入居者マスタとRX.Xデータをマージする。

    入居者マスタに存在するがRX.Xに存在しない居室（新規入居者）を
    標準固定費で追加する。✕付き居室（退居済み）は追加しない。

    また、RX.Xの既存行で名前が空だが入居者マスタに名前がある場合は
    名前を反映する。

    Args:
        residents: 入居者マスタリスト
        rx_rows: RX.Xシートのデータ（前月 or 当月）
        config: config.yaml の内容

    Returns:
        マージ済みのRxRowリスト
    """
    # 合計行を除外してからマージ
    filtered_rows = [rx for rx in rx_rows if rx.room != "合計"]

    rx_by_room: dict[str, RxRow] = {}
    for rx in filtered_rows:
        if rx.room:
            rx_by_room[rx.room] = rx

    # 入居者マスタから名前マップを構築（✕なし居室のみ）
    resident_map: dict[str, str] = {}
    for res in residents:
        if not res.room or not res.name:
            continue
        # ✕付き居室は退居済み — 新規追加しない
        if "✕" in res.room or "×" in res.room:
            continue
        resident_map[res.room] = res.name

    fixed = config.get("fixed_charges", {})
    merged = list(filtered_rows)

    for room, name in resident_map.items():
        if room in rx_by_room:
            existing = rx_by_room[room]
            was_empty = not existing.name
            # 名前が空なら入居者マスタから反映
            if was_empty and name:
                existing.name = name
            # 固定費が未設定（None or 0）→ 標準固定費をセット
            # ・名前を新たにセットした場合（新規入居者の空行）
            # ・既に名前があるが固定費がNone（未セットアップ）
            needs_charges = (was_empty and existing.name) or (
                existing.name and existing.rent is None
            )
            if needs_charges:
                existing.rent = fixed.get("rent")
                existing.management = fixed.get("management")
                existing.common = fixed.get("common")
                existing.water = fixed.get("water")
                existing.utility = fixed.get("utility")
        else:
            # 新規入居者: 標準固定費でRxRowを作成
            new_rx = RxRow(
                room=room,
                name=name,
                rent=fixed.get("rent"),
                management=fixed.get("management"),
                common=fixed.get("common"),
                water=fixed.get("water"),
                utility=fixed.get("utility"),
            )
            merged.append(new_rx)

    # 居室番号でソート（数字は数値順、英数混合は文字列順）
    def _room_sort_key(rx: RxRow) -> tuple:
        room = rx.room or ""
        # ✕付きは末尾に配置
        if "✕" in room or "×" in room:
            return (1, room)
        try:
            return (0, int(room))
        except ValueError:
            return (0, room)

    merged.sort(key=_room_sort_key)
    return merged


def build_all_billings(
    rx_rows: list[RxRow],
    meal_by_room: dict[str, int],
    nick_by_room: dict[str, tuple[int, int]],
    config: dict,
    is_new_month: bool = False,
) -> list[BillingResult]:
    """全入居者の請求データを構築する。

    Args:
        rx_rows: RX.Xシートのデータ
        meal_by_room: {居室番号: 食費月額}
        nick_by_room: {居室番号: (オムツ, 日用品)}
        config: config.yaml
        is_new_month: Trueなら前月データからの新規作成

    Returns:
        BillingResultのリスト
    """
    results = []
    for rx in rx_rows:
        meal_amount = meal_by_room.get(rx.room)
        nick_data = nick_by_room.get(rx.room)
        diaper_amount = nick_data[0] if nick_data else None
        supply_amount = nick_data[1] if nick_data else None

        billing = build_billing(
            rx, meal_amount, diaper_amount, supply_amount, config,
            is_new_month=is_new_month,
        )
        results.append(billing)

    return results
