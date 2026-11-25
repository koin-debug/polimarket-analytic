import argparse
from collections import defaultdict
import csv
from datetime import datetime, timedelta, timezone
import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import requests

BINANCE_KLINES_URL = "https://data-api.binance.vision/api/v3/klines"
FALLBACK_BINANCE_URL = "https://api.binance.com/api/v3/klines"


def normalize_cache_key(val):
  """Приводит любой формат времени (мс, секунды или строку) к виду 'YYYY-MM-DD HH:MM'."""
  val_str = str(val).strip()
  if not val_str or val_str.lower() in ["open_time", "none"]:
    return None

  if val_str.isdigit():
    ts = int(val_str)
    if ts > 1e11:  # миллисекунды
      ts = ts // 1000
    try:
      return datetime.fromtimestamp(ts, tz=timezone.utc).strftime(
          "%Y-%m-%d %H:%M"
      )
    except Exception:
      return None

  return val_str[:16]


def parse_candle_window(title: str, expiry_str: str):
  """Извлекает пару, длительность свечи и время открытия свечи (UTC)."""
  try:
    end_dt = datetime.strptime(
        str(expiry_str).strip()[:16], "%Y-%m-%d %H:%M"
    ).replace(tzinfo=timezone.utc)
  except Exception:
    return "BTCUSDT", None, 15

  symbol = (
      "ETHUSDT"
      if "eth" in str(title).lower() or "ethereum" in str(title).lower()
      else "BTCUSDT"
  )

  dur = 15
  m = re.search(r"(\d+):(\d+)(AM|PM)-(\d+):(\d+)(AM|PM)", str(title))
  if m:
    h1, min1, p1, h2, min2, p2 = m.groups()
    h1, min1, h2, min2 = int(h1), int(min1), int(h2), int(min2)
    if p1 == "PM" and h1 != 12:
      h1 += 12
    if p1 == "AM" and h1 == 12:
      h1 = 0
    if p2 == "PM" and h2 != 12:
      h2 += 12
    if p2 == "AM" and h2 == 12:
      h2 = 0
    diff = (h2 * 60 + min2) - (h1 * 60 + min1)
    if diff < 0:
      diff += 24 * 60
    if diff in [5, 15, 30, 60]:
      dur = diff
  elif "5m" in str(title).lower():
    dur = 5

  start_dt = end_dt - timedelta(minutes=dur)
  return symbol, start_dt, dur


def load_or_fetch_candles(
    window_meta: list, cache_file: str = "btc_m15_full_cache.csv"
) -> dict:
  """Читает кэш свечей и быстро докачивает недостающие пачками по 1000 штук."""
  cache = {}
  target_file = cache_file if os.path.exists(cache_file) else None
  if not target_file and os.path.exists("btc_m15_streaks_cache.csv"):
    target_file = "btc_m15_streaks_cache.csv"

  if target_file:
    print(f"📖 Чтение локального кэша '{target_file}'...", flush=True)
    t0 = datetime.now()
    with open(target_file, mode="r", encoding="utf-8") as f:
      reader = csv.reader(f)
      header = next(reader, None)
      idx_open = 1
      idx_close = 4 if header and len(header) >= 5 else 2

      for row in reader:
        if row and len(row) >= 3:
          norm_k = normalize_cache_key(row[0])
          if norm_k:
            try:
              # Ключ сохраняем как (символ, дата-строка)
              cache[("BTCUSDT", norm_k)] = (
                  float(row[idx_open]),
                  float(row[idx_close]),
              )
            except (ValueError, IndexError):
              continue

    sec = (datetime.now() - t0).total_seconds()
    print(f"   ✓ Распознано {len(cache)} свечей из кэша за {sec:.2f} сек.")

  # Находим недостающие окна с группировкой по символу и таймфрейму
  missing_by_group = defaultdict(list)
  for sym, start_dt, dur, key in window_meta:
    if not start_dt or not key:
      continue
    if (sym, key) not in cache:
      missing_by_group[(sym, dur)].append(start_dt)

  total_missing = sum(len(v) for v in missing_by_group.values())
  new_candles_to_save = []

  if total_missing > 0:
    print(
        f"🌐 Не найдено в кэше: {total_missing} свечей. Быстрая пакетная"
        " загрузка...",
        flush=True,
    )

    for (sym, dur), dts in missing_by_group.items():
      tf_str = f"{dur}m"
      interval_ms = dur * 60 * 1000
      min_dt = min(dts)
      max_dt = max(dts)

      start_ts = int(min_dt.timestamp() * 1000)
      end_ts = int(max_dt.timestamp() * 1000) + interval_ms

      curr_ts = start_ts
      batch_cnt = 0

      while curr_ts <= end_ts:
        batch_cnt += 1
        params = {
            "symbol": sym,
            "interval": tf_str,
            "startTime": curr_ts,
            "endTime": end_ts,
            "limit": 1000,
        }

        data = None
        for base_url in [BINANCE_KLINES_URL, FALLBACK_BINANCE_URL]:
          try:
            r = requests.get(base_url, params=params, timeout=8)
            if r.status_code == 200:
              data = r.json()
              break
          except Exception:
            continue

        if not data or not isinstance(data, list):
          break

        for c in data:
          c_open_ms = int(c[0])
          c_dt_str = datetime.fromtimestamp(
              c_open_ms / 1000, tz=timezone.utc
          ).strftime("%Y-%m-%d %H:%M")
          o_val, h_val, l_val, c_val, v_val = (
              float(c[1]),
              float(c[2]),
              float(c[3]),
              float(c[4]),
              float(c[5]),
          )

          cache[(sym, c_dt_str)] = (o_val, c_val)

          # Добавляем в список для сохранения в CSV
          if sym == "BTCUSDT" and dur == 15:
            new_candles_to_save.append(
                (c_open_ms, o_val, h_val, l_val, c_val, v_val)
            )

        last_open_ts = int(data[-1][0])
        curr_ts = last_open_ts + interval_ms

        print(
            f"   ✓ [{sym} {tf_str}] Загружена пачка #{batch_cnt} (+{len(data)}"
            " свечей)",
            flush=True,
        )

        if len(data) < 1000:
          break

    # Дописываем новые свечи в файл кэша на диске
    if new_candles_to_save and target_file:
      try:
        with open(target_file, mode="a", newline="", encoding="utf-8") as f_app:
          writer = csv.writer(f_app)
          writer.writerows(new_candles_to_save)
        print(
            f"💾 В кэш '{target_file}' успешно дозаписано"
            f" {len(new_candles_to_save)} свечей. Больше качаться не будут!",
            flush=True,
        )
      except Exception as e:
        print(f"⚠️ Не удалось дописать в кэш: {e}")

  return cache


def run_analysis(
    input_file: str, output_file: str, cache_file: str = "btc_m15_full_cache.csv"
):
  if not os.path.exists(input_file):
    print(f"❌ Файл {input_file} не найден!")
    return

  print(f"📊 Чтение ончейн-сделок из {input_file}...")
  wb_in = load_workbook(input_file, data_only=True)
  ws_in = wb_in.active

  raw_rows = list(ws_in.iter_rows(values_only=True))
  if not raw_rows or len(raw_rows) < 2:
    print("❌ Файл пуст или содержит только заголовок!")
    return

  headers_in = [str(c or "").strip() for c in raw_rows[0]]
  idx_time = next(
      (i for i, h in enumerate(headers_in) if "время" in h.lower()), 1
  )
  idx_market = next(
      (i for i, h in enumerate(headers_in) if "рынок" in h.lower()), 2
  )
  idx_expiry = next(
      (i for i, h in enumerate(headers_in) if "окно" in h.lower()), 3
  )
  idx_outcome = next(
      (i for i, h in enumerate(headers_in) if "исход" in h.lower()), 4
  )
  idx_shares = next(
      (i for i, h in enumerate(headers_in) if "shares" in h.lower()), 5
  )
  idx_price = next(
      (i for i, h in enumerate(headers_in) if "цена" in h.lower()), 6
  )
  idx_cost = next(
      (i for i, h in enumerate(headers_in) if "сумма" in h.lower()), 7
  )
  idx_hash = next((i for i, h in enumerate(headers_in) if "tx" in h.lower()), 8)

  trades_list = []
  for row in raw_rows[1:]:
    if not row or not any(row):
      continue
    t_str = str(row[idx_time]).strip()
    try:
      dt = datetime.strptime(t_str, "%Y-%m-%d %H:%M:%S")
    except Exception:
      try:
        dt = datetime.strptime(t_str[:16], "%Y-%m-%d %H:%M")
      except Exception:
        dt = datetime.min
    trades_list.append((dt, row))

  # Сортировка строго по хронологии
  trades_list.sort(key=lambda x: x[0])

  window_meta = []
  for _, r in trades_list:
    mkt = str(r[idx_market] or "")
    exp = str(r[idx_expiry] or "")
    sym, start_dt, dur = parse_candle_window(mkt, exp)
    if start_dt:
      key = start_dt.strftime("%Y-%m-%d %H:%M")
      window_meta.append((sym, start_dt, dur, key))
    else:
      window_meta.append((sym, None, dur, None))

  candles_cache = load_or_fetch_candles(window_meta, cache_file)

  print("⚡ Расчёт финансовых результатов и серий...")
  processed_trades = []
  cum_pnl = 0.0

  for idx, (_, r) in enumerate(trades_list):
    sym, start_dt, dur, key = window_meta[idx]
    chosen_outcome = str(r[idx_outcome] or "").strip()
    shares = float(r[idx_shares] or 0.0)
    entry_price = float(r[idx_price] or 0.0)
    cost = float(r[idx_cost] or 0.0)

    candle = candles_cache.get((sym, key)) if key else None
    if candle:
      op, cl = candle
      candle_fact = "Up" if cl >= op else "Down"
      diff_pts = cl - op
      candle_str = f"{cl:.1f} ({diff_pts:+.1f})"
    else:
      candle_fact = "UNKNOWN"
      candle_str = "N/A"

    if candle_fact == "UNKNOWN":
      win_status = "UNKNOWN"
      payout = 0.0
      pnl = 0.0
    elif chosen_outcome.lower() == candle_fact.lower():
      win_status = "Победа"
      payout = round(shares * 1.0, 2)
      pnl = round(payout - cost, 2)
    else:
      win_status = "Поражение"
      payout = 0.0
      pnl = round(-cost, 2)

    cum_pnl = round(cum_pnl + pnl, 2)
    roi_pct = (pnl / cost * 100) if cost > 0 else 0.0

    processed_trades.append({
        "№": idx + 1,
        "Время сделки (UTC)": str(r[idx_time] or "-"),
        "Рынок": str(r[idx_market] or "-"),
        "Окно экспирации (UTC)": str(r[idx_expiry] or "-"),
        "Свеча (UTC)": key if key else "-",
        "Исход": chosen_outcome,
        "Факт свечи": candle_fact,
        "Свеча Close (pts)": candle_str,
        "Shares": shares,
        "Цена входа ($)": entry_price,
        "Сумма ставки ($)": cost,
        "Результат": win_status,
        "Выплата ($)": payout,
        "Чистый PnL ($)": pnl,
        "ROI (%)": round(roi_pct, 1),
        "Накопленный PnL ($)": cum_pnl,
        "Tx Hash": str(r[idx_hash] or "-"),
    })

  ladder_series = []
  curr_series_id = 1
  curr_streak = 0
  curr_series_spend = 0.0

  for idx, t in enumerate(processed_trades):
    curr_streak += 1
    curr_series_spend += t["Сумма ставки ($)"]
    is_win = t["Результат"] == "Победа"

    if is_win or idx == len(processed_trades) - 1:
      series_pnl = (
          t["Выплата ($)"] - curr_series_spend if is_win else -curr_series_spend
      )
      ladder_series.append({
          "Серия №": curr_series_id,
          "Время старта": t["Время сделки (UTC)"],
          "Шагов в серии": curr_streak,
          "Итог серии": "Закрыта в плюс" if is_win else "Не закрыта",
          "Общий расход ($)": round(curr_series_spend, 2),
          "Итоговая выплата ($)": t["Выплата ($)"] if is_win else 0.0,
          "Чистый PnL серии ($)": round(series_pnl, 2),
      })
      curr_series_id += 1
      curr_streak = 0
      curr_series_spend = 0.0

  monthly_data = defaultdict(list)
  for t in processed_trades:
    m = t["Время сделки (UTC)"][:7] if len(t["Время сделки (UTC)"]) >= 7 else "-"
    monthly_data[m].append(t)

  timeline_rows = []
  for m, items in sorted(monthly_data.items()):
    wins = len([x for x in items if x["Результат"] == "Победа"])
    losses = len([x for x in items if x["Результат"] == "Поражение"])
    tot = wins + losses
    wr = (wins / tot * 100) if tot > 0 else 0.0
    turnover = sum(x["Сумма ставки ($)"] for x in items)
    m_pnl = sum(x["Чистый PnL ($)"] for x in items)
    gp = sum(x["Чистый PnL ($)"] for x in items if x["Чистый PnL ($)"] > 0)
    gl = abs(sum(x["Чистый PnL ($)"] for x in items if x["Чистый PnL ($)"] < 0))
    pf = round(gp / gl, 2) if gl > 0 else 999.0

    timeline_rows.append({
        "Период (Месяц)": m,
        "Сделок": len(items),
        "Побед": wins,
        "Поражений": losses,
        "Винрейт (%)": round(wr, 2),
        "Оборот ($)": round(turnover, 2),
        "Чистый PnL ($)": round(m_pnl, 2),
        "Profit Factor": pf,
    })

  total_trades = len(processed_trades)
  wins_count = len([x for x in processed_trades if x["Результат"] == "Победа"])
  losses_count = len(
      [x for x in processed_trades if x["Результат"] == "Поражение"]
  )
  overall_wr = (
      (wins_count / (wins_count + losses_count) * 100)
      if (wins_count + losses_count) > 0
      else 0.0
  )
  gross_profit = sum(
      x["Чистый PnL ($)"]
      for x in processed_trades
      if x["Чистый PnL ($)"] > 0
  )
  gross_loss = abs(
      sum(
          x["Чистый PnL ($)"]
          for x in processed_trades
          if x["Чистый PnL ($)"] < 0
      )
  )
  total_pnl = sum(x["Чистый PnL ($)"] for x in processed_trades)
  profit_factor = (
      round(gross_profit / gross_loss, 2) if gross_loss > 0 else 999.0
  )
  max_ladder = (
      max(x["Шагов в серии"] for x in ladder_series) if ladder_series else 0
  )
  all_costs = [x["Сумма ставки ($)"] for x in processed_trades]
  avg_bet = (sum(all_costs) / len(all_costs)) if all_costs else 0.0
  max_bet = max(all_costs) if all_costs else 0.0

  summary_rows = [
      ("Период торговли", f"{processed_trades[0]['Время сделки (UTC)'][:10]} — {processed_trades[-1]['Время сделки (UTC)'][:10]}"),
      ("Всего совершено ставок", f"{total_trades} шт."),
      ("Побед (Wins)", f"{wins_count} шт."),
      ("Поражений (Losses)", f"{losses_count} шт."),
      ("Итоговый Винрейт (Win Rate)", f"{overall_wr:.2f}%"),
      ("Суммарный оборот ставок", f"${sum(all_costs):.2f}"),
      ("Валовая прибыль (Gross Profit)", f"${gross_profit:.2f}"),
      ("Валовый убыток (Gross Loss)", f"-${gross_loss:.2f}"),
      ("Чистый PnL (Net Profit)", f"${total_pnl:+.2f}"),
      ("Профит-фактор (Profit Factor)", f"{profit_factor}"),
      ("Средняя ставка", f"${avg_bet:.2f}"),
      ("Максимальная ставка", f"${max_bet:.2f}"),
      ("Макс. шагов в серии (Мартингейл)", f"{max_ladder} подряд"),
  ]

  wb_out = Workbook()
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
  win_fill = PatternFill(
      start_color="ECFDF5", end_color="ECFDF5", fill_type="solid"
  )
  win_font = Font(name="Calibri", size=11, color="065F46", bold=True)
  loss_fill = PatternFill(
      start_color="FEF2F2", end_color="FEF2F2", fill_type="solid"
  )
  loss_font = Font(name="Calibri", size=11, color="991B1B", bold=True)

  def format_and_append(ws, dict_list, sheet_title, is_journal=False):
    ws.title = sheet_title
    ws.views.sheetView[0].showGridLines = True
    if not dict_list:
      return

    keys = list(dict_list[0].keys())
    ws.append(keys)

    for col_idx in range(1, len(keys) + 1):
      c = ws.cell(row=1, column=col_idx)
      c.fill = header_fill
      c.font = header_font
      c.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, item in enumerate(dict_list, start=2):
      row_vals = [item[k] for k in keys]
      ws.append(row_vals)
      row_res = item.get("Результат") if is_journal else None

      for col_idx, val in enumerate(row_vals, start=1):
        cell = ws.cell(row=r_idx, column=col_idx)
        cell.border = thin_border
        k_name = keys[col_idx - 1]

        if is_journal and row_res == "Победа" and col_idx in [12, 14]:
          cell.fill = win_fill
          cell.font = win_font
        elif is_journal and row_res == "Поражение" and col_idx in [12, 14]:
          cell.fill = loss_fill
          cell.font = loss_font

        if isinstance(val, (int, float)):
          cell.alignment = Alignment(horizontal="right", vertical="center")
          if "($)" in k_name:
            cell.number_format = "$#,##0.00"
          elif "(%)" in k_name:
            cell.number_format = "0.0%"
        else:
          cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
      max_len = max(len(str(cell.value or "")) for cell in col)
      col_letter = get_column_letter(col[0].column)
      ws.column_dimensions[col_letter].width = max(min(max_len + 3, 40), 12)

  # Вкладки книги
  ws1 = wb_out.active
  format_and_append(ws1, processed_trades, "Журнал сделок (PnL)", is_journal=True)

  ws2 = wb_out.create_sheet()
  format_and_append(ws2, ladder_series, "Анализ серий (Мартингейл)")

  ws3 = wb_out.create_sheet()
  format_and_append(ws3, timeline_rows, "Динамика по месяцам")

  ws4 = wb_out.create_sheet()
  ws4.title = "Сводная аналитика"
  ws4.views.sheetView[0].showGridLines = True
  ws4.append(["Метрика", "Значение"])
  for col_idx in (1, 2):
    c = ws4.cell(row=1, column=col_idx)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")

  for r_idx, (m, v) in enumerate(summary_rows, start=2):
    ws4.append([m, v])
    c1 = ws4.cell(row=r_idx, column=1)
    c2 = ws4.cell(row=r_idx, column=2)
    c1.border = thin_border
    c2.border = thin_border
    c1.alignment = Alignment(horizontal="left", vertical="center")
    c2.alignment = Alignment(horizontal="right", vertical="center")

  for col in ws4.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws4.column_dimensions[col_letter].width = max(min(max_len + 3, 45), 18)

  try:
    wb_out.save(output_file)
    print("\n" + "=" * 60)
    print(f"🎉 АНАЛИЗ ЗАВЕРШЁН! Отчёт сохранён в: {output_file}")
    print(f"   • Обработано ставок: {total_trades}")
    print(f"   • Винрейт:           {overall_wr:.2f}%")
    print(f"   • Итоговый PnL:      ${total_pnl:+.2f}")
    print(f"   • Профит-фактор:     {profit_factor}")
    print(f"   • Макс. шаг лесенки: {max_ladder}")
    print("=" * 60)
  except PermissionError:
    fallback_name = (
        output_file.replace(".xlsx", "")
        + f"_{datetime.now().strftime('%H%M%S')}.xlsx"
    )
    wb_out.save(fallback_name)
    print(
        f"\n⚠️ Файл '{output_file}' открыт в Excel! Результат записан в:"
        f" {fallback_name}"
    )


if __name__ == "__main__":
  parser = argparse.ArgumentParser()
  parser.add_argument(
      "--file", type=str, default="vlabs_all_onchain_trades.xlsx"
  )
  parser.add_argument("--out", type=str, default="vlabs_final_pnl_report.xlsx")
  parser.add_argument("--cache", type=str, default="btc_m15_full_cache.csv")
  args = parser.parse_args()

  run_analysis(args.file, args.out, args.cache)