import argparse
import csv
import math
import os
from collections import defaultdict
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

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
                    int(row[0]),    # open_time
                    float(row[1]),  # open
                    float(row[2]),  # high
                    float(row[3]),  # low
                    float(row[4]),  # close
                    float(row[5])   # volume
                ))
    return candles

def get_session(hour_utc):
    if 0 <= hour_utc < 8:
        return "Азия (00:00 - 08:00 UTC)"
    elif 8 <= hour_utc < 13:
        return "Европа (08:00 - 13:00 UTC)"
    elif 13 <= hour_utc < 21:
        return "США / Перекрытие (13:00 - 21:00 UTC)"
    else:
        return "Тихий вечер (21:00 - 24:00 UTC)"

def analyze_time_distribution(candles, streak_len=10):
    n = len(candles)
    if n < streak_len:
        return [], {}

    # Цвета свечей: 1 (Up), -1 (Down), 0 (Doji)
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
    while i <= n - streak_len:
        s_color = colors[i]
        if s_color == 0:
            i += 1
            continue

        # Проверка серии
        is_streak = all(colors[i + j] == s_color for j in range(streak_len))

        if is_streak:
            # Отсекаем нахождение внутри более длинной серии
            if i > 0 and colors[i - 1] == s_color:
                i += 1
                continue

            start_dt = datetime.fromtimestamp(candles[i][0] / 1000, tz=timezone.utc)
            end_dt = datetime.fromtimestamp(candles[i + streak_len - 1][0] / 1000, tz=timezone.utc)

            events.append({
                "start_time": start_dt,
                "end_time": end_dt,
                "start_hour": start_dt.hour,
                "end_hour": end_dt.hour,
                "weekday": start_dt.strftime("%A"),
                "weekday_num": start_dt.weekday(),
                "session_start": get_session(start_dt.hour),
                "direction": "🟢 UP" if s_color == 1 else "🔴 DOWN"
            })
            i += streak_len
        else:
            i += 1

    return events

def print_and_export_stats(events, streak_len, out_file="time_distribution_streaks.xlsx"):
    total = len(events)
    if total == 0:
        print(f"Серий длиной {streak_len} не найдено.")
        return

    # Распределение по часам старта
    hour_counts = defaultdict(int)
    hour_up = defaultdict(int)
    hour_down = defaultdict(int)
    
    # Распределение по сессиям
    session_counts = defaultdict(int)
    # Распределение по дням недели
    dow_counts = defaultdict(int)

    dow_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    dow_ru = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]

    for e in events:
        h = e["start_hour"]
        hour_counts[h] += 1
        if "UP" in e["direction"]:
            hour_up[h] += 1
        else:
            hour_down[h] += 1

        session_counts[e["session_start"]] += 1
        dow_counts[e["weekday_num"]] += 1

    # Хи-квадрат тест на равномерность по часам (k = 24, ожидаемое = total / 24)
    expected_per_hour = total / 24.0
    chi2_stat = sum(((hour_counts[h] - expected_per_hour) ** 2) / expected_per_hour for h in range(24))

    print("\n" + "=" * 65)
    print(f"АНАЛИЗ ВРЕМЕННЫХ ЗОН ДЛЯ СЕРИЙ ИЗ {streak_len} СВЕЧЕЙ (Всего серий: {total})")
    print("=" * 65)
    print(f"Ожидаемое количество на 1 час при случайном шуме: {expected_per_hour:.2f}")
    print(f"Статистика Chi-Square (по часам): {chi2_stat:.2f} (критическое ~35.17 при p=0.05)")
    if chi2_stat > 35.17:
        print("💡 ВЫВОД: Распределение НЕ случайно! Есть статистически выраженные часы активности.")
    else:
        print("⚖️ ВЫВОД: Распределение близко к равномерному (случайное возникновение).")
    print("-" * 65)

    print("\nРАСПРЕДЕЛЕНИЕ ПО СЕССИЯМ:")
    session_hours = {
        "Азия (00:00 - 08:00 UTC)": 8,
        "Европа (08:00 - 13:00 UTC)": 5,
        "США / Перекрытие (13:00 - 21:00 UTC)": 8,
        "Тихий вечер (21:00 - 24:00 UTC)": 3
    }
    for sess, h_len in session_hours.items():
        cnt = session_counts[sess]
        rate_per_hour = cnt / h_len
        pct = (cnt / total) * 100
        print(f"  • {sess:<36}: {cnt:>3} шт ({pct:>5.1f}%) | Плотность: {rate_per_hour:.2f} серий/час")

    print("\nТОП-3 ЧАСА С НАИВЫСШИМ ЧИСЛОМ СЕРИЙ (UTC):")
    sorted_hours = sorted(range(24), key=lambda h: hour_counts[h], reverse=True)
    for h in sorted_hours[:3]:
        print(f"  • {h:02d}:00 - {h:02d}:59 UTC -> {hour_counts[h]} серий ({hour_counts[h]/total*100:.1f}%) [🟢 {hour_up[h]} | 🔴 {hour_down[h]}]")

    print("\nТОП-3 ЧАСА С НАИМЕНЬШИМ ЧИСЛОМ СЕРИЙ (UTC):")
    for h in sorted_hours[-3:]:
        print(f"  • {h:02d}:00 - {h:02d}:59 UTC -> {hour_counts[h]} серий ({hour_counts[h]/total*100:.1f}%)")
    print("=" * 65)

    # Генерация Excel-отчета
    wb = Workbook()
    
    # Лист 1: Почасовая сетка
    ws_hours = wb.active
    ws_hours.title = "По часам (UTC)"
    ws_hours.append(["Час (UTC)", "Всего серий", "🟢 Зеленых", "🔴 Красных", "% от всех", "Отклонение от среднего"])
    for cell in ws_hours[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for h in range(24):
        cnt = hour_counts[h]
        diff_pct = ((cnt - expected_per_hour) / expected_per_hour) * 100 if expected_per_hour > 0 else 0
        ws_hours.append([
            f"{h:02d}:00", cnt, hour_up[h], hour_down[h],
            f"{(cnt/total)*100:.1f}%", f"{diff_pct:+.1f}%"
        ])

    # Лист 2: По дням недели
    ws_dow = wb.create_sheet("По дням недели")
    ws_dow.append(["День недели", "Всего серий", "% от общего числа"])
    for cell in ws_dow[1]:
        cell.font = Font(bold=True)

    expected_dow = total / 7.0
    for idx, name_ru in enumerate(dow_ru):
        cnt = dow_counts[idx]
        ws_dow.append([name_ru, cnt, f"{(cnt/total)*100:.1f}%"])

    # Лист 3: Все серии с отметкой времени
    ws_all = wb.create_sheet("Все найденные серии")
    ws_all.append(["Старт (UTC)", "Конец (UTC)", "Направление", "Час старта", "Торговая сессия", "День недели"])
    for cell in ws_all[1]:
        cell.font = Font(bold=True)

    for e in events:
        ws_all.append([
            e["start_time"].strftime("%Y-%m-%d %H:%M"),
            e["end_time"].strftime("%Y-%m-%d %H:%M"),
            e["direction"],
            e["start_hour"],
            e["session_start"],
            dow_ru[e["weekday_num"]]
        ])

    # Подгонка колонок
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(out_file)
    print(f"\nДетальная таблица с расчетами сохранена в: {out_file}\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Анализ временных зон возникновения серий свечей")
    parser.add_argument("--streak", type=int, default=10, help="Длина серии свечей (по умолчанию 10)")
    parser.add_argument("--cache", type=str, default=CACHE_FILE, help="Путь к файлу кэша")
    args = parser.parse_args()

    candles = load_candles(args.cache)
    if candles:
        events = analyze_time_distribution(candles, streak_len=args.streak)
        out_name = f"time_zones_streak_{args.streak}.xlsx"
        print_and_export_stats(events, streak_len=args.streak, out_file=out_name)