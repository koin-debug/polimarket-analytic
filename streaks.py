import csv
from datetime import datetime, timezone
import os
import time
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import requests

BINANCE_VISION_URL = "https://data-api.binance.vision/api/v3/klines"
BINANCE_FALLBACK_URL = "https://api.binance.com/api/v3/klines"


def sync_and_load_candles(
    symbol="BTCUSDT", tf="15m", start_str="2021-01-01 00:00:00"
) -> list:
  """Загружает свечи выбранного таймфрейма из кэша и докачивает недостающие."""
  coin_prefix = symbol.replace("USDT", "").lower()
  interval_ms = 5 * 60 * 1000 if tf == "5m" else 15 * 60 * 1000

  # Имя кэша под таймфрейм
  cache_path = f"{coin_prefix}_{tf}_streaks_cache.csv"

  # Проверка наличия старых файлов кэша
  if not os.path.exists(cache_path):
    possible_old_names = []
    if tf == "5m":
      possible_old_names = [f"{coin_prefix}_m5_full_cache.csv"]
    else:
      possible_old_names = [
          f"{coin_prefix}_m15_2021_2025_cache.csv",
          f"{coin_prefix}_m15_full_cache.csv",
      ]

    for old_name in possible_old_names:
      if os.path.exists(old_name):
        try:
          # Если файл с 6 колонками OHLCV — используем его напрямую
          with open(old_name, mode="r", encoding="utf-8") as f_check:
            reader = csv.reader(f_check)
            hdr = next(reader, None)
            if hdr and len(hdr) >= 6:
              cache_path = old_name
              print(
                  f"📦 Подключен существующий кэш: '{old_name}'", flush=True
              )
              break
        except Exception:
          pass

  candles = []
  if os.path.exists(cache_path):
    print(f"Чтение локального кэша '{cache_path}'...", flush=True)
    t0 = time.time()
    corrupted = False
    with open(cache_path, mode="r", encoding="utf-8") as f:
      reader = csv.reader(f)
      header = next(reader, None)

      if not header or len(header) < 6:
        corrupted = True
      else:
        for row in reader:
          if row and len(row) >= 6:
            try:
              candles.append((
                  int(row[0]),
                  float(row[1]),
                  float(row[2]),
                  float(row[3]),
                  float(row[4]),
                  float(row[5]),
              ))
            except (ValueError, IndexError):
              continue

    if corrupted or not candles:
      print(
          f"⚠️ Файл '{cache_path}' имел несовместимый формат. Создаем чистую"
          " базу...",
          flush=True,
      )
      candles = []
      cache_path = f"{coin_prefix}_{tf}_streaks_cache.csv"
    else:
      print(
          f"Из кэша загружено {len(candles)} свечей ({tf}) по {symbol} за"
          f" {time.time() - t0:.2f} сек.",
          flush=True,
      )

  # Определение начальной точки дозагрузки
  if candles:
    next_start_ts = candles[-1][0] + interval_ms
  else:
    next_start_ts = int(
        datetime.strptime(start_str, "%Y-%m-%d %H:%M:%S")
        .replace(tzinfo=timezone.utc)
        .timestamp()
        * 1000
    )

  now_ms = int(time.time() * 1000)
  current_open_candle_ts = (now_ms // interval_ms) * interval_ms

  if next_start_ts >= current_open_candle_ts:
    last_dt = datetime.fromtimestamp(
        candles[-1][0] / 1000, tz=timezone.utc
    ).strftime("%Y-%m-%d %H:%M")
    print(
        f"Кэш {symbol} [{tf}] полностью актуален (последняя закрытая свеча:"
        f" {last_dt} UTC).\n",
        flush=True,
    )
    return candles

  new_candles = []
  current_ts = next_start_ts
  batch = 1
  total_est_batches = max(
      1, (current_open_candle_ts - next_start_ts) // (1000 * interval_ms) + 1
  )

  from_dt = datetime.fromtimestamp(
      next_start_ts / 1000, tz=timezone.utc
  ).strftime("%Y-%m-%d %H:%M")
  print(
      f"Докачиваем свечи {symbol} [{tf}] начиная с {from_dt} UTC...", flush=True
  )

  while current_ts < current_open_candle_ts:
    params = {
        "symbol": symbol,
        "interval": tf,
        "startTime": current_ts,
        "endTime": current_open_candle_ts - 1,
        "limit": 1000,
    }
    data = None
    for endpoint in [BINANCE_VISION_URL, BINANCE_FALLBACK_URL]:
      try:
        response = requests.get(endpoint, params=params, timeout=10)
        if response.status_code == 200:
          data = response.json()
          break
      except Exception:
        continue

    if not data or not isinstance(data, list):
      time.sleep(1.5)
      continue

    for c in data:
      open_time = int(c[0])
      if open_time < current_open_candle_ts:
        new_candles.append((
            open_time,
            float(c[1]),
            float(c[2]),
            float(c[3]),
            float(c[4]),
            float(c[5]),
        ))

    last_date = datetime.fromtimestamp(
        data[-1][0] / 1000, tz=timezone.utc
    ).strftime("%Y-%m-%d %H:%M")
    print(
        f" [{batch:03d}/{total_est_batches}] {symbol} [{tf}] — скачано новых"
        f" свечей: {len(new_candles)} (до {last_date} UTC)",
        flush=True,
    )

    current_ts = int(data[-1][0]) + interval_ms
    batch += 1
    time.sleep(0.04)

  # Сохранение докачанных свечей в CSV
  if new_candles:
    file_exists = os.path.exists(cache_path) and len(candles) > 0
    mode = "a" if file_exists else "w"
    with open(cache_path, mode=mode, newline="", encoding="utf-8") as f:
      writer = csv.writer(f)
      if not file_exists:
        writer.writerow(
            ["open_time", "open", "high", "low", "close", "volume"]
        )
      writer.writerows(new_candles)

    candles.extend(new_candles)
    print(
        f"💾 В кэш '{cache_path}' добавлено {len(new_candles)} свечей. Всего в"
        f" базе: {len(candles)}.\n",
        flush=True,
    )

  return candles


def find_streaks(candles: list, min_length: int = 8) -> list:
  """Находит серии одноцветных свечей и замеряет внутрисерийные откаты."""
  n = len(candles)
  streaks = []
  i = 0

  while i < n:
    op, cl = candles[i][1], candles[i][4]
    if cl > op:
      color = "Green"
    elif cl < op:
      color = "Red"
    else:
      i += 1
      continue

    j = i + 1
    while j < n:
      next_op, next_cl = candles[j][1], candles[j][4]
      next_color = (
          "Green" if next_cl > next_op else ("Red" if next_cl < next_op else "D")
      )
      if next_color == color:
        j += 1
      else:
        break

    streak_len = j - i
    if streak_len >= min_length:
      streak_candles = candles[i:j]
      start_ts = streak_candles[0][0]
      end_ts = streak_candles[-1][0]
      first_open = streak_candles[0][1]
      last_close = streak_candles[-1][4]

      if color == "Green":
        net_move = last_close - first_open
        running_peak = streak_candles[0][2]
        max_pullback = 0.0
        for c in streak_candles[1:]:
          pb = running_peak - c[3]
          if pb > max_pullback:
            max_pullback = pb
          if c[2] > running_peak:
            running_peak = c[2]
      else:
        net_move = first_open - last_close
        running_trough = streak_candles[0][3]
        max_pullback = 0.0
        for c in streak_candles[1:]:
          pb = c[2] - running_trough
          if pb > max_pullback:
            max_pullback = pb
          if c[3] < running_trough:
            running_trough = c[3]

      net_pct = (net_move / first_open) * 100
      pb_pct = (max_pullback / first_open) * 100

      if j < n:
        rev_op, rev_cl = candles[j][1], candles[j][4]
        if color == "Green":
          reversal_success = rev_cl < rev_op
          reversal_move = rev_op - rev_cl
        else:
          reversal_success = rev_cl > rev_op
          reversal_move = rev_cl - rev_op
      else:
        reversal_success = None
        reversal_move = 0.0

      streaks.append({
          "color": color,
          "length": streak_len,
          "start_ts": start_ts,
          "end_ts": end_ts,
          "start_dt": datetime.fromtimestamp(
              start_ts / 1000, tz=timezone.utc
          ).strftime("%Y-%m-%d %H:%M"),
          "end_dt": datetime.fromtimestamp(
              end_ts / 1000, tz=timezone.utc
          ).strftime("%Y-%m-%d %H:%M"),
          "year": datetime.fromtimestamp(
              start_ts / 1000, tz=timezone.utc
          ).year,
          "first_open": first_open,
          "last_close": last_close,
          "net_move": round(net_move, 2),
          "net_pct": round(net_pct, 2),
          "max_pullback": round(max_pullback, 2),
          "pb_pct": round(pb_pct, 2),
          "reversal_success": reversal_success,
          "reversal_move": round(reversal_move, 2),
      })

    i = j

  return streaks


def export_to_excel(
    streaks: list, symbol: str, tf: str, min_len: int, outfile: str
):
  """Формирует отчет в Excel с разбивкой по годам."""
  wb = Workbook()
  thin_border = Border(
      left=Side(style="thin", color="E5E7EB"),
      right=Side(style="thin", color="E5E7EB"),
      top=Side(style="thin", color="E5E7EB"),
      bottom=Side(style="thin", color="E5E7EB"),
  )
  header_fill = PatternFill(
      start_color="1F2937", end_color="1F2937", fill_type="solid"
  )
  header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  green_font = Font(name="Calibri", size=11, color="166534", bold=True)
  red_font = Font(name="Calibri", size=11, color="991B1B", bold=True)

  headers = [
      "№",
      "Год",
      "Начало серии (UTC)",
      "Окончание серии (UTC)",
      "Тип",
      "Длина свечей",
      "Цена входа ($)",
      "Цена закрытия ($)",
      "Чистый ход ($)",
      "Чистый ход (%)",
      "Макс. откат ($)",
      "Откат (%)",
      "1-я свеча контртренд?",
      "Дельта отката ($)",
  ]

  ws_sum = wb.active
  ws_sum.title = "Сравнение по годам"
  ws_sum.views.sheetView[0].showGridLines = True

  sum_headers = [
      "Год",
      "Всего серий",
      "🟢 Зеленых (Up)",
      "🔴 Красных (Down)",
      "Разворотов на 1-й свече",
      "Винрейт контртренда (%)",
      "Макс. длина серии",
  ]
  ws_sum.append(sum_headers)
  for col_idx in range(1, len(sum_headers) + 1):
    c = ws_sum.cell(row=1, column=col_idx)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")

  years = sorted(list(set(s["year"] for s in streaks)))
  r_idx = 2
  for y in years:
    y_streaks = [s for s in streaks if s["year"] == y]
    grn = sum(1 for s in y_streaks if s["color"] == "Green")
    red = sum(1 for s in y_streaks if s["color"] == "Red")
    revs = sum(1 for s in y_streaks if s["reversal_success"])
    wr = (revs / len(y_streaks) * 100) if y_streaks else 0.0
    max_l = max((s["length"] for s in y_streaks), default=0)

    row_vals = [y, len(y_streaks), grn, red, revs, round(wr, 1), max_l]
    ws_sum.append(row_vals)
    for c_i, val in enumerate(row_vals, 1):
      cell = ws_sum.cell(row=r_idx, column=c_i)
      cell.border = thin_border
      cell.alignment = Alignment(horizontal="center", vertical="center")
    r_idx += 1

  tot_grn = sum(1 for s in streaks if s["color"] == "Green")
  tot_red = sum(1 for s in streaks if s["color"] == "Red")
  tot_revs = sum(1 for s in streaks if s["reversal_success"])
  tot_wr = (tot_revs / len(streaks) * 100) if streaks else 0.0
  tot_max_l = max((s["length"] for s in streaks), default=0)
  total_row = [
      "ВСЕ ГОДЫ",
      len(streaks),
      tot_grn,
      tot_red,
      tot_revs,
      round(tot_wr, 1),
      tot_max_l,
  ]
  ws_sum.append(total_row)
  for c_i in range(1, len(total_row) + 1):
    cell = ws_sum.cell(row=r_idx, column=c_i)
    cell.font = Font(name="Calibri", size=11, bold=True)
    cell.border = thin_border
    cell.alignment = Alignment(horizontal="center", vertical="center")

  def append_streak_sheet(ws, items, sheet_name):
    ws.title = sheet_name
    ws.views.sheetView[0].showGridLines = True
    ws.append(headers)
    for col_i in range(1, len(headers) + 1):
      c = ws.cell(row=1, column=col_i)
      c.fill = header_fill
      c.font = header_font
      c.alignment = Alignment(horizontal="center", vertical="center")

    for r_i, s in enumerate(items, start=2):
      rev_text = (
          "ДА"
          if s["reversal_success"] is True
          else ("НЕТ" if s["reversal_success"] is False else "-")
      )
      row_v = [
          r_i - 1,
          s["year"],
          s["start_dt"],
          s["end_dt"],
          "🟢 UP" if s["color"] == "Green" else "🔴 DOWN",
          s["length"],
          s["first_open"],
          s["last_close"],
          s["net_move"],
          s["net_pct"],
          s["max_pullback"],
          s["pb_pct"],
          rev_text,
          s["reversal_move"],
      ]
      ws.append(row_v)
      for c_i, val in enumerate(row_v, 1):
        cell = ws.cell(row=r_i, column=c_i)
        cell.border = thin_border
        if c_i in [1, 2, 3, 4, 5, 6, 13]:
          cell.alignment = Alignment(horizontal="center", vertical="center")
          if c_i == 5:
            cell.font = green_font if "UP" in val else red_font
        else:
          cell.alignment = Alignment(horizontal="right", vertical="center")
          if c_i in [7, 8, 9, 11, 14]:
            cell.number_format = "$#,##0.00"
          elif c_i in [10, 12]:
            cell.number_format = '0.00"%"'

    for col in ws.columns:
      max_len = max(len(str(c.value or "")) for c in col)
      col_letter = get_column_letter(col[0].column)
      ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

  ws_all = wb.create_sheet("Все серии")
  append_streak_sheet(ws_all, streaks, f"Все серии ({min_len}+)")

  for y in sorted(years, reverse=True):
    ws_y = wb.create_sheet(str(y))
    append_streak_sheet(ws_y, [s for s in streaks if s["year"] == y], str(y))

  for col in ws_sum.columns:
    max_len = max(len(str(c.value or "")) for c in col)
    col_letter = get_column_letter(col[0].column)
    ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 14)

  wb.save(outfile)
  print(f"\n🎉 Анализ сохранен в файл: {outfile}")


if __name__ == "__main__":
  print("=" * 60)
  print("АНАЛИЗАТОР СЕРИЙ ОДНОЦВЕТНЫХ СВЕЧЕЙ (M15 / M5)")
  print("=" * 60)
  print("Шаг 1: Выберите монету:")
  print("  1 — BTC (BTCUSDT)")
  print("  2 — ETH (ETHUSDT)")
  print("  Или введите тикер вручную (SOL, BNB, XRP)")

  raw_choice = input("Ваш выбор (по умолчанию 1 - BTC): ").strip().upper()
  if raw_choice == "2" or raw_choice == "ETH":
    ticker = "ETHUSDT"
    prefix = "eth"
  elif raw_choice in ["", "1", "BTC"]:
    ticker = "BTCUSDT"
    prefix = "btc"
  else:
    prefix = raw_choice.replace("USDT", "").lower()
    ticker = f"{prefix.upper()}USDT"

  print("\nШаг 2: Выберите таймфрейм свечей:")
  print("  1 — 15 минут (15m) [По умолчанию]")
  print("  2 — 5 минут (5m)")
  tf_input = input("Ваш выбор (1 или 2): ").strip()
  selected_tf = "5m" if tf_input in ["2", "5", "5m", "5M"] else "15m"

  default_len = 10 if selected_tf == "5m" else 8
  raw_len = input(
      f"\nШаг 3: Введите минимальную длину серии свечей (по умолчанию"
      f" {default_len}): "
  ).strip()
  min_streak = (
      int(raw_len) if raw_len.isdigit() and int(raw_len) > 0 else default_len
  )

  # Загрузка / докачка кэша выбранного таймфрейма
  candles_data = sync_and_load_candles(symbol=ticker, tf=selected_tf)

  if candles_data:
    print(f"🔍 Поиск серий [{selected_tf}] длиной >= {min_streak} свечей...")
    found_streaks = find_streaks(candles_data, min_length=min_streak)
    print(f"   Найдено серий: {len(found_streaks)} шт.")

    out_file = f"{prefix}_{selected_tf}_streaks_{min_streak}plus_report.xlsx"
    export_to_excel(found_streaks, ticker, selected_tf, min_streak, out_file)