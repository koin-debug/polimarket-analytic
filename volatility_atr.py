import os
import csv
import time
from datetime import datetime, timezone
from collections import defaultdict
import requests
from openpyxl import Workbook
from openpyxl.styles import Font


def sync_and_load_tf(symbol="BTCUSDT", interval="15m", interval_ms=900_000, start_str="2021-01-01 00:00:00"):
    coin_prefix = symbol.replace("USDT", "").lower()
    tf_label = "m15" if interval == "15m" else "m5"
    cache_path = f"{coin_prefix}_{tf_label}_full_cache.csv"

    candles = []
    if os.path.exists(cache_path):
        print(f"Чтение локального кэша '{cache_path}'...", flush=True)
        t0 = time.time()
        with open(cache_path, mode="r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if row:
                    candles.append((
                        int(row[0]), float(row[1]), float(row[2]),
                        float(row[3]), float(row[4]), float(row[5])
                    ))
        print(f"Загружено {len(candles)} свечей ({interval}) за {time.time() - t0:.2f} сек.", flush=True)

    next_start_ts = candles[-1][0] + interval_ms if candles else int(
        datetime.strptime(start_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc).timestamp() * 1000
    )
    now_ms = int(time.time() * 1000)
    current_open_ts = (now_ms // interval_ms) * interval_ms

    if next_start_ts >= current_open_ts:
        return candles

    url = "https://data-api.binance.vision/api/v3/klines"
    new_candles = []
    current_ts = next_start_ts
    batch = 1
    total_est = max(1, (current_open_ts - next_start_ts) // (1000 * interval_ms) + 1)

    print(f"Докачиваем {symbol} ({interval}) с Binance Vision...", flush=True)
    while current_ts < current_open_ts:
        params = {
            "symbol": symbol,
            "interval": interval,
            "startTime": current_ts,
            "endTime": current_open_ts - 1,
            "limit": 1000,
        }
        try:
            resp = requests.get(url, params=params, timeout=10)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            print(f"Ошибка сети ({e}), повтор через 2 сек...", flush=True)
            time.sleep(2)
            continue

        if not data or not isinstance(data, list):
            break

        for c in data:
            ot = int(c[0])
            if ot < current_open_ts:
                new_candles.append((ot, float(c[1]), float(c[2]), float(c[3]), float(c[4]), float(c[5])))

        last_dt = datetime.fromtimestamp(data[-1][0] / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M")
        print(f" [{batch:03d}/{total_est}] {symbol} {interval}: +{len(new_candles)} свечей (до {last_dt} UTC)", flush=True)
        current_ts = int(data[-1][0]) + interval_ms
        batch += 1
        time.sleep(0.04)

    if new_candles:
        file_exists = os.path.exists(cache_path)
        with open(cache_path, mode="a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["open_time", "open", "high", "low", "close", "volume"])
            writer.writerows(new_candles)
        candles.extend(new_candles)

    return candles


def compute_true_ranges(candles):
    """Считает True Range (TR) для каждой свечи: max(H-L, |H-PrevClose|, |L-PrevClose|)."""
    trs = []
    prev_close = None
    for ot, o, h, l, cl, vol in candles:
        if prev_close is None:
            tr = h - l
        else:
            tr = max(h - l, abs(h - prev_close), abs(l - prev_close))
        tr_pct = (tr / o) * 100 if o > 0 else 0.0
        trs.append((ot, o, h, l, cl, vol, tr, tr_pct))
        prev_close = cl
    return trs


def build_daily_volatility_report(symbol, m15_candles, m5_candles=None, atr_period=14):
    coin_name = symbol.replace("USDT", "")
    print("\nРасчет дневной и внутридневной волатильности (ATR)...", flush=True)

    # 1. Группируем М15 по дням (YYYY-MM-DD) и считаем TR каждой М15 свечи
    m15_with_tr = compute_true_ranges(m15_candles)
    daily_m15 = defaultdict(list)
    for item in m15_with_tr:
        day_str = datetime.fromtimestamp(item[0] / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
        daily_m15[day_str].append(item)

    # 2. Если есть М5 — группируем М5 по дням
    daily_m5 = defaultdict(list)
    if m5_candles:
        m5_with_tr = compute_true_ranges(m5_candles)
        for item in m5_with_tr:
            day_str = datetime.fromtimestamp(item[0] / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
            daily_m5[day_str].append(item)

    # 3. Собираем дневные свечи D1 из М15 и считаем классический ATR(14) по Уайлдеру (RMA)
    sorted_days = sorted(daily_m15.keys())
    daily_rows = []
    prev_day_close = None
    d1_atr = None
    d1_tr_history = []

    for day_str in sorted_days:
        items15 = daily_m15[day_str]
        d_open = items15[0][1]
        d_high = max(x[2] for x in items15)
        d_low = min(x[3] for x in items15)
        d_close = items15[-1][4]
        d_vol = sum(x[5] for x in items15)

        # True Range самого дня (D1)
        if prev_day_close is None:
            d_tr = d_high - d_low
        else:
            d_tr = max(d_high - d_low, abs(d_high - prev_day_close), abs(d_low - prev_day_close))
        prev_day_close = d_close
        d_tr_pct = (d_tr / d_open) * 100

        # Сглаженный ATR(14) D1 (Wilder's Smoothing)
        d1_tr_history.append(d_tr)
        if len(d1_tr_history) < atr_period:
            d1_atr = sum(d1_tr_history) / len(d1_tr_history)
        elif len(d1_tr_history) == atr_period:
            d1_atr = sum(d1_tr_history) / atr_period
        else:
            d1_atr = (d1_atr * (atr_period - 1) + d_tr) / atr_period
        d1_atr_pct = (d1_atr / d_close) * 100

        # Метрики M15 за этот день
        m15_avg_tr = sum(x[6] for x in items15) / len(items15)
        m15_avg_tr_pct = sum(x[7] for x in items15) / len(items15)
        m15_max_tr = max(x[6] for x in items15)

        row_data = [
            int(day_str[:4]),
            day_str,
            round(d_open, 2),
            round(d_close, 2),
            round(d_tr, 2),
            round(d_tr_pct, 2),
            round(d1_atr, 2),
            round(d1_atr_pct, 2),
            round(m15_avg_tr, 2),
            round(m15_avg_tr_pct, 3),
            round(m15_max_tr, 2),
        ]

        # Метрики M5 за этот день (если загружены)
        if m5_candles:
            items5 = daily_m5.get(day_str, [])
            if items5:
                m5_avg_tr = sum(x[6] for x in items5) / len(items5)
                m5_avg_tr_pct = sum(x[7] for x in items5) / len(items5)
                m5_max_tr = max(x[6] for x in items5)
                row_data.extend([round(m5_avg_tr, 2), round(m5_avg_tr_pct, 3), round(m5_max_tr, 2)])
            else:
                row_data.extend([0, 0, 0])

        row_data.append(round(d_vol, 2))
        daily_rows.append(row_data)

    # 4. Сохраняем в Excel (Сводка по годам + Полная таблица по дням + Листы по годам)
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Средний ATR по годам"

    sum_headers = [
        "Год", "Дней", "Средний дневной ход D1 ($)", "Средний дневной ход D1 (%)",
        "Средний ATR M15 ($)", "Средний ATR M15 (%)"
    ]
    if m5_candles:
        sum_headers.extend(["Средний ATR M5 ($)", "Средний ATR M5 (%)"])
    ws_summary.append(sum_headers)
    for cell in ws_summary[1]:
        cell.font = Font(bold=True)

    years = sorted(list(set(r[0] for r in daily_rows)))
    for y in years:
        y_rows = [r for r in daily_rows if r[0] == y]
        cnt = len(y_rows)
        s_row = [
            y,
            cnt,
            round(sum(r[4] for r in y_rows) / cnt, 2),
            round(sum(r[5] for r in y_rows) / cnt, 2),
            round(sum(r[8] for r in y_rows) / cnt, 2),
            round(sum(r[9] for r in y_rows) / cnt, 3),
        ]
        if m5_candles:
            s_row.extend([
                round(sum(r[11] for r in y_rows) / cnt, 2),
                round(sum(r[12] for r in y_rows) / cnt, 3),
            ])
        ws_summary.append(s_row)

    detail_headers = [
        "Год", "Дата (UTC)", "Open D1", "Close D1",
        "Размах дня TR D1 ($)", "Размах дня TR D1 (%)",
        f"ATR({atr_period}) D1 ($)", f"ATR({atr_period}) D1 (%)",
        "Средний ATR M15 за день ($)", "Средний ATR M15 за день (%)", "Макс. свеча M15 ($)"
    ]
    if m5_candles:
        detail_headers.extend([
            "Средний ATR M5 за день ($)", "Средний ATR M5 за день (%)", "Макс. свеча M5 ($)"
        ])
    detail_headers.append(f"Объем за день ({coin_name})")

    ws_all = wb.create_sheet("Все дни 2021-2026")
    ws_all.append(detail_headers)
    for cell in ws_all[1]:
        cell.font = Font(bold=True)
    for r in daily_rows:
        ws_all.append(r)

    for y in years:
        ws_y = wb.create_sheet(str(y))
        ws_y.append(detail_headers)
        for cell in ws_y[1]:
            cell.font = Font(bold=True)
        for r in daily_rows:
            if r[0] == y:
                ws_y.append(r)

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 12)

    out_file = f"{coin_name.lower()}_daily_intraday_atr_2021_2026.xlsx"
    wb.save(out_file)
    print("\n" + "=" * 60)
    print(f"Готово! Проанализировано дней: {len(daily_rows)} (2021–2026 гг.)")
    print(f"Отчет по волатильности сохранен в: {out_file}")
    print("=" * 60)


if __name__ == "__main__":
    print("Выберите монету для расчета ATR:")
    print("  1 — BTC (BTCUSDT)")
    print("  2 — ETH (ETHUSDT)")
    choice = input("Ваш выбор (по умолчанию 1 - BTC): ").strip().upper()
    symbol = "ETHUSDT" if choice in ("2", "ETH", "ETHUSDT") else (
        "BTCUSDT" if choice in ("", "1", "BTC", "BTCUSDT") else (
            choice if choice.endswith("USDT") else f"{choice}USDT"
        )
    )

    # М15 загружается мгновенно из вашего готового кэша
    m15_data = sync_and_load_candles = sync_and_load_tf(symbol, "15m", 900_000)

    include_m5 = input("\nВключить 5-минутный ATR (M5)? [y/N - при первом запуске скачает ~610 пачек]: ").strip().lower()
    m5_data = None
    if include_m5 in ("y", "yes", "д", "да", "1"):
        m5_data = sync_and_load_tf(symbol, "5m", 300_000)

    if m15_data:
        build_daily_volatility_report(symbol, m15_data, m5_data, atr_period=14)