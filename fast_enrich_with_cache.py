import csv
from datetime import datetime, timezone
import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import requests

INPUT_FILE = "vlabs_trades_1000.xlsx"
OUTPUT_FILE = "vlabs_trades_with_pnl_calculated.xlsx"
CACHE_15M = "btc_m15_full_cache.csv"
CACHE_5M = "btc_m5_full_cache.csv"

MONTH_MAP = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12,
}


def parse_window(title: str, default_year: int = 2026):
    m = re.search(
        r"([A-Za-z]+)\s+(\d+),\s+(\d+):(\d+)(AM|PM)-(\d+):(\d+)(AM|PM)\s+ET",
        str(title),
    )
    if not m:
        return None, None
    mon_str, day_s, h1_s, m1_s, ap1, h2_s, m2_s, ap2 = m.groups()
    month = MONTH_MAP.get(mon_str)
    if not month:
        return None, None
    day = int(day_s)
    h1, m1 = int(h1_s), int(m1_s)
    if ap1 == "PM" and h1 != 12:
        h1 += 12
    elif ap1 == "AM" and h1 == 12:
        h1 = 0
    h2, m2 = int(h2_s), int(m2_s)
    if ap2 == "PM" and h2 != 12:
        h2 += 12
    elif ap2 == "AM" and h2 == 12:
        h2 = 0

    utc_offset = 4  # EDT
    t1 = datetime(
        default_year, month, day, h1, m1, 0, tzinfo=timezone.utc
    ).timestamp()
    start_ms = int((t1 + utc_offset * 3600) * 1000)

    dur_min = (h2 * 60 + m2) - (h1 * 60 + m1)
    if dur_min < 0:
        dur_min += 24 * 60
    return start_ms, dur_min


def load_cache_dict(cache_file):
    """Быстро читает локальный кэш свечей в словарь {open_time_ms: 'Up'/'Down'}."""
    candles = {}
    if not os.path.exists(cache_file):
        return candles

    print(f"📖 Чтение локального кэша '{cache_file}'...")
    with open(cache_file, mode="r", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            if row:
                try:
                    ot = int(row[0])
                    o, c = float(row[1]), float(row[4])
                    candles[ot] = "Up" if c >= o else "Down"
                except Exception:
                    continue
    print(f"  ✓ Загружено {len(candles)} свечей из файла {cache_file}")
    return candles


def fetch_missing_candles_batch(start_ms, end_ms, interval_m):
    """Если в кэше нет нужного диапазона, скачивает его за 1-2 быстрых запроса."""
    fetched = {}
    curr_ts = start_ms
    limit = 1000

    while curr_ts <= end_ms:
        url = "https://data-api.binance.vision/api/v3/klines"
        params = {
            "symbol": "BTCUSDT",
            "interval": f"{interval_m}m",
            "startTime": curr_ts,
            "limit": limit,
        }
        try:
            r = requests.get(url, params=params, timeout=5)
            if r.status_code != 200:
                break
            data = r.json()
            if not data or not isinstance(data, list):
                break

            for item in data:
                ts = int(item[0])
                o, c = float(item[1]), float(item[4])
                fetched[ts] = "Up" if c >= o else "Down"

            curr_ts = int(data[-1][0]) + (interval_m * 60 * 1000)
            if len(data) < limit:
                break
        except Exception:
            break

    return fetched


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"Файл {INPUT_FILE} не найден!")
        return

    # Загружаем существующие локальные кэши
    cache_15m = load_cache_dict(CACHE_15M)
    cache_5m = load_cache_dict(CACHE_5M)

    wb_in = load_workbook(INPUT_FILE)
    ws_in = wb_in.active
    rows = list(ws_in.iter_rows(min_row=2, values_only=True))

    parsed_windows = []
    missing_15m_ts = []
    missing_5m_ts = []

    for r in rows:
        title = r[1]
        start_ms, dur = parse_window(title)
        parsed_windows.append((start_ms, dur))

        if start_ms and dur == 15 and start_ms not in cache_15m:
            missing_15m_ts.append(start_ms)
        elif start_ms and dur == 5 and start_ms not in cache_5m:
            missing_5m_ts.append(start_ms)

    # Если в кэше чего-то не хватило — докачиваем недостающее за 1 быстрый запрос
    if missing_15m_ts:
        print(
            f"🔄 В кэше нет свечей за октябрь 2026. Быстро докачиваем диапазон 15m..."
        )
        new_15m = fetch_missing_candles_batch(
            min(missing_15m_ts), max(missing_15m_ts), 15
        )
        cache_15m.update(new_15m)

    if missing_5m_ts:
        print(f"🔄 Быстро докачиваем недостающий диапазон 5m...")
        new_5m = fetch_missing_candles_batch(
            min(missing_5m_ts), max(missing_5m_ts), 5
        )
        cache_5m.update(new_5m)

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Сделки с результатами"
    ws_out.views.sheetView[0].showGridLines = True

    headers = [
        "№",
        "Рынок / Окно",
        "Время сделки (UTC)",
        "Исход",
        "Цена ($)",
        "Shares (Акции)",
        "Сумма ($)",
        "Факт. результат",
        "Итог ставки",
        "Выплата ($)",
        "Чистый PnL ($)",
        "Доходность (%)",
        "Side",
        "Tx Hash",
    ]
    ws_out.append(headers)

    header_fill = PatternFill(
        start_color="1F2937", end_color="1F2937", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    win_fill = PatternFill(
        start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"
    )
    win_font = Font(name="Calibri", size=10, bold=True, color="166534")
    loss_fill = PatternFill(
        start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"
    )
    loss_font = Font(name="Calibri", size=10, bold=True, color="991B1B")
    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB"),
    )

    for col in range(1, len(headers) + 1):
        c = ws_out.cell(row=1, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    total_wins, total_losses, total_pnl = 0, 0, 0.0

    print("📊 Сведение сделок с результатами свечей...")
    for idx, r in enumerate(rows, start=1):
        num, market, dt_str, bet_side, price, size, cost, side, tx_hash = r[
            :9
        ]
        price = float(price or 0.0)
        size = float(size or 0.0)
        cost = float(cost or round(price * size, 2))
        bet_side = str(bet_side).strip().title()

        start_ms, dur = parsed_windows[idx - 1]
        actual = "UNKNOWN"

        if dur == 15 and start_ms in cache_15m:
            actual = cache_15m[start_ms]
        elif dur == 5 and start_ms in cache_5m:
            actual = cache_5m[start_ms]

        verdict = "Неизвестно"
        payout, pnl, roi = 0.0, 0.0, 0.0

        if actual in ["Up", "Down"]:
            if bet_side == actual:
                verdict = "Победа (Win)"
                payout = round(size * 1.0, 2)
                pnl = round(payout - cost, 2)
                roi = round(pnl / cost * 100, 1) if cost > 0 else 0.0
                total_wins += 1
            else:
                verdict = "Поражение (Loss)"
                payout = 0.0
                pnl = round(-cost, 2)
                roi = -100.0
                total_losses += 1
            total_pnl += pnl

        out_row = [
            num,
            market,
            dt_str,
            bet_side,
            price,
            size,
            cost,
            actual,
            verdict,
            payout,
            pnl,
            roi,
            side,
            tx_hash,
        ]
        ws_out.append(out_row)

        curr_row = idx + 1
        for col_idx in range(1, len(out_row) + 1):
            cell = ws_out.cell(row=curr_row, column=col_idx)
            cell.border = thin_border
            if col_idx in [1, 3, 4, 8, 9, 13]:
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in [5, 6, 7, 10, 11, 12]:
                cell.alignment = Alignment(horizontal="right")
                if col_idx == 5:
                    cell.number_format = "$#,##0.000"
                elif col_idx == 6:
                    cell.number_format = "#,##0.0"
                elif col_idx in [7, 10, 11]:
                    cell.number_format = "$#,##0.00"
                elif col_idx == 12:
                    cell.number_format = '0.0"%"'

        c_v = ws_out.cell(row=curr_row, column=9)
        c_p = ws_out.cell(row=curr_row, column=11)
        if verdict == "Победа (Win)":
            c_v.fill = win_fill
            c_v.font = win_font
            c_p.font = win_font
        elif verdict == "Поражение (Loss)":
            c_v.fill = loss_fill
            c_v.font = loss_font
            c_p.font = loss_font

    for col in ws_out.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = col[0].column_letter
        ws_out.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb_out.save(OUTPUT_FILE)
    total_res = total_wins + total_losses
    wr = (total_wins / total_res * 100) if total_res > 0 else 0.0

    print("\n" + "=" * 55)
    print(f"Готово! Результаты записаны в: {OUTPUT_FILE}")
    print(f"Всего рассчитано сделок: {total_res}")
    print(f"  • Побед:     {total_wins} ({wr:.1f}%)")
    print(f"  • Поражений: {total_losses} ({100 - wr:.1f}%)")
    print(f"  • Чистый рассчитанный PnL: ${total_pnl:+.2f}")
    print("=" * 55)


if __name__ == "__main__":
    main()