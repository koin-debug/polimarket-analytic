import argparse
import csv
import os
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

CACHE_FILE = "btc_m15_full_cache.csv"


def load_candles(cache_path=CACHE_FILE):
    if not os.path.exists(cache_path):
        for alt in ["btc_m15_2021_2025_cache.csv"]:
            if os.path.exists(alt):
                cache_path = alt
                break
        else:
            print(f"Файл кэша {cache_path} не найден!")
            return []

    candles = []
    with open(cache_path, mode="r", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            if row:
                candles.append((
                    int(row[0]),
                    float(row[1]),
                    float(row[2]),
                    float(row[3]),
                    float(row[4]),
                    float(row[5]),
                ))
    return candles


def calculate_atr(candles, period=14):
    n = len(candles)
    atr = [0.0] * n
    tr = [0.0] * n

    for i in range(n):
        h, l = candles[i][2], candles[i][3]
        if i == 0:
            tr[i] = h - l
        else:
            prev_close = candles[i - 1][4]
            tr[i] = max(h - l, abs(h - prev_close), abs(l - prev_close))

    if n > period:
        atr[period - 1] = sum(tr[:period]) / period
        for i in range(period, n):
            atr[i] = (atr[i - 1] * (period - 1) + tr[i]) / period

    return atr


def analyze_normalized_expansion(candles, streak_len=4, atr_period=14):
    n = len(candles)
    if n < streak_len + atr_period + 2:
        return []

    atr_values = calculate_atr(candles, period=atr_period)

    colors = []
    for c in candles:
        if c[4] > c[1]:
            colors.append(1)
        elif c[4] < c[1]:
            colors.append(-1)
        else:
            colors.append(0)

    events = []
    i = atr_period

    while i <= n - streak_len - 1:
        s_color = colors[i]
        if s_color == 0:
            i += 1
            continue

        is_streak = all(colors[i + j] == s_color for j in range(streak_len))

        if is_streak:
            if i > 0 and colors[i - 1] == s_color:
                i += 1
                continue

            start_candle = candles[i]
            end_idx = i + streak_len - 1
            end_candle = candles[end_idx]

            net_move_usd = abs(end_candle[4] - start_candle[1])
            current_atr = atr_values[end_idx]

            # Нормализованный показатель: средняя величина бара серии относительно 1 ATR
            expected_total_atr = streak_len * current_atr
            norm_ratio = (
                round(net_move_usd / expected_total_atr, 2)
                if expected_total_atr > 0
                else 0.0
            )

            next_color = colors[end_idx + 1]
            first_is_reversal = (next_color == -s_color)

            start_dt = datetime.fromtimestamp(start_candle[0] / 1000, tz=timezone.utc)
            end_dt = datetime.fromtimestamp(end_candle[0] / 1000, tz=timezone.utc)

            events.append({
                "start_time": start_dt.strftime("%Y-%m-%d %H:%M"),
                "end_time": end_dt.strftime("%Y-%m-%d %H:%M"),
                "direction": "🟢 UP" if s_color == 1 else "🔴 DOWN",
                "net_move_usd": round(net_move_usd, 2),
                "atr": round(current_atr, 2),
                "norm_ratio": norm_ratio,
                "first_reversal": first_is_reversal,
            })

            i += streak_len
        else:
            i += 1

    return events


def print_and_export_norm_report(events, streak_len, out_file):
    if not events:
        print("Серий указанной длины не найдено.")
        return

    total = len(events)

    # Категории нормализованной скорости импульса
    buckets = {
        "1. Слабый дрейф (< 0.6x от ATR/свеча)": {"events": []},
        "2. Средний темп (0.6x - 1.0x от ATR/свеча)": {"events": []},
        "3. Быстрый импульс (1.0x - 1.5x от ATR/свеча)": {"events": []},
        "4. Взрывной вынос (> 1.5x от ATR/свеча)": {"events": []},
    }

    for e in events:
        r = e["norm_ratio"]
        if r < 0.6:
            buckets["1. Слабый дрейф (< 0.6x от ATR/свеча)"]["events"].append(e)
        elif 0.6 <= r < 1.0:
            buckets["2. Средний темп (0.6x - 1.0x от ATR/свеча)"]["events"].append(e)
        elif 1.0 <= r < 1.5:
            buckets["3. Быстрый импульс (1.0x - 1.5x от ATR/свеча)"]["events"].append(e)
        else:
            buckets["4. Взрывной вынос (> 1.5x от ATR/свеча)"]["events"].append(e)

    print("\n" + "=" * 75)
    print(
        f"НОРМАЛИЗОВАННЫЙ ТЕСТ ATR (Делитель: N * ATR) | Серия: {streak_len} свечей"
    )
    print(f"Всего событий в выборке: {total}")
    print("=" * 75)
    print(
        f"{'Категория темпа (Ход / [N*ATR])':<42} | {'Всего':<6} | {'Разворотов':<10} | {'Винрейт %':<10}"
    )
    print("-" * 75)

    summary_rows = []
    for name, data in buckets.items():
        b_events = data["events"]
        cnt = len(b_events)
        if cnt > 0:
            rev_cnt = sum(1 for x in b_events if x["first_reversal"])
            winrate = (rev_cnt / cnt) * 100
            print(f"{name:<42} | {cnt:<6} | {rev_cnt:<10} | {winrate:>8.1f}%")
            summary_rows.append([name, cnt, rev_cnt, round(winrate, 1)])
        else:
            print(f"{name:<42} | 0      | 0          |     0.0%")
            summary_rows.append([name, 0, 0, 0.0])

    print("=" * 75)

    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Сводка (Норм. ATR)"
    ws_summary.append([
        "Категория темпа",
        "Количество серий",
        "Разворотов",
        "Винрейт контртренда (%)",
    ])
    for cell in ws_summary[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for row in summary_rows:
        ws_summary.append(row)

    ws_details = wb.create_sheet("Все серии")
    ws_details.append([
        "Начало",
        "Конец",
        "Направление",
        "Ход ($)",
        "ATR ($)",
        "Индекс темпа (Ход / [N*ATR])",
        "Разворот?",
    ])
    for cell in ws_details[1]:
        cell.font = Font(bold=True)

    for e in events:
        ws_details.append([
            e["start_time"],
            e["end_time"],
            e["direction"],
            e["net_move_usd"],
            e["atr"],
            e["norm_ratio"],
            "ДА" if e["first_reversal"] else "НЕТ",
        ])

    for ws in [ws_summary, ws_details]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max(
                max_len + 3, 12
            )

    wb.save(out_file)
    print(f"Отчет сохранен в: {out_file}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Нормализованный анализ темпа серии к ATR"
    )
    parser.add_argument(
        "--streak", type=int, default=4, help="Длина серии свечей"
    )
    parser.add_argument("--atr", type=int, default=14, help="Период ATR")
    parser.add_argument(
        "--cache", type=str, default=CACHE_FILE, help="Путь к файлу кэша"
    )
    args = parser.parse_args()

    candles = load_candles(args.cache)
    if candles:
        events = analyze_normalized_expansion(
            candles, streak_len=args.streak, atr_period=args.atr
        )
        out_name = f"norm_atr_streak_{args.streak}.xlsx"
        print_and_export_norm_report(
            events, streak_len=args.streak, out_file=out_name
        )