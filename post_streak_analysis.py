import argparse
import csv
import os
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.styles import Font

CACHE_FILE = "btc_m15_full_cache.csv"


def load_candles(cache_path=CACHE_FILE):
    if not os.path.exists(cache_path):
        print(f"Файл кэша {cache_path} не найден!")
        return []
    candles = []
    with open(cache_path, mode="r", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            if row:
                candles.append((
                    int(row[0]),  # open_time
                    float(row[1]),  # open
                    float(row[2]),  # high
                    float(row[3]),  # low
                    float(row[4]),  # close
                    float(row[5]),  # volume
                ))
    return candles


def analyze_post_streaks(candles, streak_len=10, horizon=20):
    n = len(candles)
    if n < streak_len + horizon:
        return []

    # 1 (Up), -1 (Down), 0 (Doji)
    colors = []
    for c in candles:
        if c[4] > c[1]:
            colors.append(1)
        elif c[4] < c[1]:
            colors.append(-1)
        else:
            colors.append(0)

    events = []
    i = 0
    while i <= n - streak_len - horizon:
        s_color = colors[i]
        if s_color == 0:
            i += 1
            continue

        # Проверяем серию из streak_len одноцветных свечей
        is_streak = all(colors[i + j] == s_color for j in range(streak_len))

        if is_streak:
            # Отсекаем продолжение более длинной серии (перед началом должен быть другой цвет)
            if i > 0 and colors[i - 1] == s_color:
                i += 1
                continue

            streak_end_idx = i + streak_len - 1
            next_candles = candles[streak_end_idx + 1 : streak_end_idx + 1 + horizon]
            next_colors = colors[streak_end_idx + 1 : streak_end_idx + 1 + horizon]

            # Если на конце датасета не хватает свечей на полный горизонт
            if len(next_colors) < horizon:
                break

            with_trend = sum(1 for c in next_colors if c == s_color)
            counter_trend = sum(1 for c in next_colors if c == -s_color)
            doji_count = sum(1 for c in next_colors if c == 0)

            first_is_reversal = (next_colors[0] == -s_color)

            # Ценовой ход на всем горизонте
            base_price = candles[streak_end_idx][4]
            final_price = next_candles[-1][4]
            total_change_pct = (final_price - base_price) / base_price * 100
            trend_return_pct = total_change_pct if s_color == 1 else -total_change_pct

            start_dt = datetime.fromtimestamp(candles[i][0] / 1000, tz=timezone.utc)
            end_dt = datetime.fromtimestamp(candles[streak_end_idx][0] / 1000, tz=timezone.utc)

            events.append({
                "start_time": start_dt.strftime("%Y-%m-%d %H:%M"),
                "end_time": end_dt.strftime("%Y-%m-%d %H:%M"),
                "direction": "🟢 UP" if s_color == 1 else "🔴 DOWN",
                "first_reversal": first_is_reversal,
                "with_trend_count": with_trend,
                "counter_trend_count": counter_trend,
                "doji_count": doji_count,
                "counter_trend_pct": round((counter_trend / horizon) * 100, 1),
                "trend_return_pct": round(trend_return_pct, 2),
                "seq_colors": "".join(["T" if c == s_color else ("C" if c == -s_color else "D") for c in next_colors]),
            })

            # Смещаемся за пределы проверенного окна
            i += streak_len
        else:
            i += 1

    return events


def print_and_export_summary(events, streak_len, horizon, out_file):
    if not events:
        print("Серий указанной длины не найдено.")
        return

    total = len(events)
    first_rev_count = sum(1 for e in events if e["first_reversal"])
    at_least_one_rev = sum(1 for e in events if e["counter_trend_count"] >= 1)

    avg_counter = sum(e["counter_trend_count"] for e in events) / total
    avg_with_trend = sum(e["with_trend_count"] for e in events) / total
    avg_doji = sum(e["doji_count"] for e in events) / total

    pct_counter = (avg_counter / horizon) * 100
    pct_trend = (avg_with_trend / horizon) * 100

    print("\n" + "=" * 60)
    print(f"РЕЗУЛЬТАТЫ АНАЛИЗА: {total} серий из {streak_len} свечей")
    print(f"Горизонт проверки: следующие {horizon} свечей")
    print("=" * 60)
    print(f"• 1-я свеча сразу дала разворот (контртренд): {first_rev_count}/{total} ({first_rev_count / total * 100:.1f}%)")
    print(f"• Хотя бы 1 откат за {horizon} свечей произошёл: {at_least_one_rev}/{total} ({at_least_one_rev / total * 100:.1f}%)")
    print(f"• Среднее распределение на отрезке {horizon} свечей:")
    print(f"    - Контртрендовых свечей: {avg_counter:.2f} ({pct_counter:.1f}%)")
    print(f"    - Свечей по тренду:      {avg_with_trend:.2f} ({pct_trend:.1f}%)")
    if avg_doji > 0.01:
        print(f"    - Доджи (нейтральных):  {avg_doji:.2f} ({(avg_doji / horizon) * 100:.1f}%)")
    print("=" * 60)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Post_{streak_len}_H{horizon}"
    headers = [
        "Начало серии", "Конец серии", "Направление", "1-я свеча разворот?",
        f"Контртренд (из {horizon})", f"По тренду (из {horizon})", "Доджи",
        "% Контртренда", "Суммарный ход по тренду (%)",
        f"Паттерн {horizon} свечей (T=тренд, C=контр, D=доджи)"
    ]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)

    for e in events:
        ws.append([
            e["start_time"], e["end_time"], e["direction"],
            "ДА" if e["first_reversal"] else "НЕТ",
            e["counter_trend_count"], e["with_trend_count"], e["doji_count"],
            e["counter_trend_pct"], e["trend_return_pct"], e["seq_colors"]
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(out_file)
    print(f"Детальный отчет сохранен в: {out_file}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Анализ поведения свечей после серии")
    parser.add_argument("--streak", type=int, default=10, help="Длина серии свечей")
    parser.add_argument("--horizon", type=int, default=20, help="Горизонт проверки свечей")
    parser.add_argument("--cache", type=str, default=CACHE_FILE, help="Путь к файлу кэша")
    args = parser.parse_args()

    candles = load_candles(args.cache)
    if candles:
        events = analyze_post_streaks(candles, streak_len=args.streak, horizon=args.horizon)
        out_name = f"streak_{args.streak}_horizon_{args.horizon}.xlsx"
        print_and_export_summary(events, streak_len=args.streak, horizon=args.horizon, out_file=out_name)