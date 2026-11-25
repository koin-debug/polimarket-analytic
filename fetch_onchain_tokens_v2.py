import argparse
from datetime import datetime, timezone
import json
import os
import time
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}

PRIMARY_URL = "https://api.etherscan.io/v2/api"
GAMMA_URL = "https://gamma-api.polymarket.com"
DATA_API = "https://data-api.polymarket.com"
CACHE_FILE = "tokens_market_cache.json"


def load_token_cache() -> dict:
    """Загружает локальный кэш, отфильтровывая ошибочные заглушки."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                c = json.load(f)
                valid = {}
                for k, v in c.items():
                    title = v.get("title", "")
                    outcome = v.get("outcome", "-")
                    if (
                        "Xi Jinping" not in title
                        and title != "Polymarket Order"
                        and outcome != "-"
                    ):
                        valid[k] = v
                return valid
        except Exception:
            return {}
    return {}


def save_token_cache(cache: dict):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ Ошибка сохранения кэша: {e}")


def resolve_token_market(token_id: str, approx_ts: int, cache: dict) -> dict:
    """Определяет рынок, окно экспирации и исход (Up/Down) с валидацией токена."""
    token_str = str(token_id).strip()
    if token_str in cache:
        cached = cache[token_str]
        if (
            "Xi Jinping" not in cached.get("title", "")
            and cached.get("title") != "Polymarket Order"
            and cached.get("outcome") != "-"
        ):
            return cached

    # 1. Поиск по временным окнам вокруг сделки (15m и 5m пулы)
    w_curr_15 = (approx_ts // 900) * 900
    w_curr_5 = (approx_ts // 300) * 300
    candidate_windows = [
        w_curr_15 + 900,
        w_curr_15,
        w_curr_15 - 900,
        w_curr_5 + 300,
        w_curr_5,
    ]

    for win_ts in candidate_windows:
        for prefix in [
            "btc-updown-15m-",
            "eth-updown-15m-",
            "btc-updown-5m-",
            "eth-updown-5m-",
        ]:
            slug = f"{prefix}{win_ts}"
            try:
                r = requests.get(
                    f"{GAMMA_URL}/events",
                    params={"slug": slug},
                    headers=HEADERS,
                    timeout=5,
                )
                if r.status_code == 200:
                    events = r.json()
                    if isinstance(events, list) and len(events) > 0:
                        markets = events[0].get("markets", [])
                        for m in markets:
                            clob_ids = json.loads(m.get("clobTokenIds") or "[]")
                            if token_str in clob_ids:
                                outcomes = json.loads(m.get("outcomes") or "[]")
                                idx = clob_ids.index(token_str)
                                outcome = (
                                    outcomes[idx]
                                    if idx < len(outcomes)
                                    else "Up/Down"
                                )
                                title = (
                                    m.get("question")
                                    or m.get("groupItemTitle")
                                    or events[0].get("title")
                                )

                                end_iso = m.get("endDate") or events[0].get(
                                    "endDate"
                                )
                                win_str = "-"
                                if end_iso:
                                    win_str = str(end_iso)[:16].replace("T", " ")

                                res = {
                                    "title": title,
                                    "outcome": outcome,
                                    "window": win_str,
                                }
                                cache[token_str] = res
                                return res
            except Exception:
                pass

    # 2. Поиск через Data API trades со строгой сверкой asset_id
    try:
        r = requests.get(
            f"{DATA_API}/trades",
            params={"asset_id": token_str, "limit": 1},
            headers=HEADERS,
            timeout=6,
        )
        if r.status_code == 200:
            trades = r.json()
            if isinstance(trades, list) and len(trades) > 0:
                t = trades[0]
                if (
                    str(t.get("asset_id", "")).strip() == token_str
                    or str(t.get("token_id", "")).strip() == token_str
                ):
                    title = t.get("title") or t.get("market_slug") or "Market"
                    outcome = t.get("outcome") or "-"
                    res = {"title": title, "outcome": outcome, "window": "-"}
                    cache[token_str] = res
                    return res
    except Exception:
        pass

    # 3. Резервный формат для архивных рынков
    dt_readable = datetime.fromtimestamp(approx_ts, tz=timezone.utc).strftime(
        "%b %d, %H:%M UTC"
    )
    fallback = {
        "title": f"BTC 15m Market ({dt_readable})",
        "outcome": "Up/Down",
        "window": "-",
    }
    cache[token_str] = fallback
    return fallback


def fetch_all_usdc_expenses(address: str, apikey: str) -> dict[str, float]:
    """Скачивает историю списаний USDC (Native + Bridged) с нормализацией хэшей."""
    target = address.strip().lower()
    print(f"💵 Загрузка ончейн-списаний USDC для {target}...", flush=True)
    tx_spend_map = {}
    page = 1
    total_events = 0

    while True:
        params = {
            "chainid": 137,
            "module": "account",
            "action": "tokentx",
            "address": target,
            "page": page,
            "offset": 1000,
            "sort": "desc",
            "apikey": apikey,
        }
        try:
            r = requests.get(
                PRIMARY_URL, params=params, headers=HEADERS, timeout=15
            )
            if r.status_code != 200:
                time.sleep(1)
                continue

            data = r.json()
            items = data.get("result", [])
            if not isinstance(items, list) or len(items) == 0:
                break

            for tx in items:
                symbol = str(tx.get("tokenSymbol", "")).upper()
                if "USD" in symbol:
                    if str(tx.get("from", "")).lower() == target:
                        h = str(tx.get("hash", "")).strip().lower()
                        decimals = int(tx.get("tokenDecimal") or 6)
                        val = float(tx.get("value") or 0) / (10**decimals)
                        if val > 0 and h:
                            tx_spend_map[h] = tx_spend_map.get(h, 0.0) + val
                            total_events += 1

            if len(items) < 1000:
                break
            page += 1
            time.sleep(0.2)
        except Exception:
            time.sleep(1.5)
            continue

    print(
        f"   ✓ Найдено {len(tx_spend_map)} транзакций расхода USDC"
        f" ({total_events} переводов).\n",
        flush=True,
    )
    return tx_spend_map


def fetch_onchain_history(address: str, apikey: str, outfile: str):
    target_wallet = address.strip().lower()
    token_cache = load_token_cache()

    # Шаг 1: Карта фактических расходов в USDC
    usdc_spends = fetch_all_usdc_expenses(target_wallet, apikey)

    # Шаг 2: Выгрузка токенов исходов (ERC-1155)
    print(
        f"📡 Выгрузка ончейн-ставок (ERC-1155) для {target_wallet}...",
        flush=True,
    )
    raw_txs = []
    end_block = 999999999
    page_count = 1
    seen_keys = set()

    while True:
        params = {
            "chainid": 137,
            "module": "account",
            "action": "token1155tx",
            "address": target_wallet,
            "startblock": 0,
            "endblock": end_block,
            "page": 1,
            "offset": 1000,
            "sort": "desc",
            "apikey": apikey,
        }
        try:
            r = requests.get(
                PRIMARY_URL, params=params, headers=HEADERS, timeout=15
            )
            if r.status_code != 200:
                time.sleep(2)
                continue
            data = r.json()
            items = data.get("result", [])
            if (
                str(data.get("status", "0")) != "1"
                or not isinstance(items, list)
                or len(items) == 0
            ):
                break

            for tx in items:
                if str(tx.get("to", "")).lower() == target_wallet:
                    tx_h = str(tx.get("hash", "")).strip().lower()
                    t_id = str(tx.get("tokenID") or "-")
                    ukey = f"{tx_h}_{t_id}"
                    if ukey in seen_keys:
                        continue
                    seen_keys.add(ukey)
                    raw_txs.append(tx)

            oldest_block = int(items[-1].get("blockNumber", 0))
            print(
                f"   [Пакет {page_count:02d}] Получено {len(items)} событий"
                f" | Всего покупок: {len(raw_txs)} | Блок: {oldest_block}",
                flush=True,
            )

            if len(items) < 1000 or oldest_block <= 0:
                break
            end_block = oldest_block - 1
            page_count += 1
            time.sleep(0.25)
        except Exception:
            time.sleep(2)
            continue

    if not raw_txs:
        print("❌ Сделок не найдено.")
        return

    # Шаг 3: Резолвинг рынков и исходов
    unique_tokens = list(set(str(tx.get("tokenID")) for tx in raw_txs))
    print(
        f"\n🔍 Резолвинг рынков и исходов для {len(unique_tokens)} уникальных"
        " токенов...",
        flush=True,
    )

    token_time_map = {}
    for tx in raw_txs:
        t_id = str(tx.get("tokenID"))
        ts = int(tx.get("timeStamp") or 0)
        if t_id not in token_time_map and ts > 0:
            token_time_map[t_id] = ts

    for idx, t_id in enumerate(unique_tokens, 1):
        if (
            t_id not in token_cache
            or "Xi Jinping" in token_cache[t_id].get("title", "")
            or token_cache[t_id].get("outcome") == "-"
        ):
            approx_ts = token_time_map.get(t_id, int(time.time()))
            resolve_token_market(t_id, approx_ts, token_cache)
            time.sleep(0.12)
        if idx % 10 == 0 or idx == len(unique_tokens):
            print(
                f"   Проверено токенов: {idx}/{len(unique_tokens)}", flush=True
            )

    save_token_cache(token_cache)

    # Шаг 4: Расчет математики сделок (без ложных $500 на минималке)
    records = []
    matched_exact = 0

    for tx in raw_txs:
        tx_hash = str(tx.get("hash", "")).strip().lower()
        token_id = str(tx.get("tokenID") or "-")

        ts = int(tx.get("timeStamp", 0))
        dt_str = (
            datetime.fromtimestamp(ts, tz=timezone.utc).strftime(
                "%Y-%m-%d %H:%M:%S"
            )
            if ts
            else "-"
        )

        raw_val = float(tx.get("tokenValue") or 0)
        shares = round(raw_val / 1e6, 3)

        if tx_hash in usdc_spends and shares > 0:
            cost = round(usdc_spends[tx_hash], 2)
            price = round(cost / shares, 3)
            # Допустимый диапазон цен от $0.001 (минимальный шаг биржи) до $0.999
            if 0.001 <= price <= 0.999:
                matched_exact += 1
            else:
                price = 0.50
                cost = round(shares * price, 2)
        else:
            # Если транзакция не списывала USDC напрямую (микро-покупка или пул)
            if shares >= 500.0:
                price = 0.001
                cost = round(shares * price, 2)
            elif shares < 0.05:
                price = 0.00
                cost = 0.00
            else:
                price = 0.50
                cost = round(shares * price, 2)

        m_info = token_cache.get(
            token_id,
            {"title": "BTC 15m Market", "outcome": "-", "window": "-"},
        )

        records.append([
            len(records) + 1,
            dt_str,
            m_info["title"],
            m_info["window"],
            m_info["outcome"],
            shares,
            price,
            cost,
            tx_hash,
            tx.get("blockNumber", "-"),
        ])

    print("\n" + "-" * 55)
    print("📊 ИТОГИ ОБРАБОТКИ ДАННЫХ:")
    print(f"   • Всего записей в таблице: {len(records)}")
    print(f"   • Сделок с точной ончейн-ценой: {matched_exact}")
    print("-" * 55)

    # Шаг 5: Экспорт в Excel через openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = "Ончейн-сделки"
    ws.views.sheetView[0].showGridLines = True

    headers = [
        "№",
        "Время сделки (UTC)",
        "Рынок",
        "Окно экспирации (UTC)",
        "Исход",
        "Shares (Акции)",
        "Факт. цена ($)",
        "Сумма ставки ($)",
        "Tx Hash",
        "Block Number",
    ]
    ws.append(headers)

    header_fill = PatternFill(
        start_color="1F2937", end_color="1F2937", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB"),
    )

    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row in enumerate(records, start=2):
        ws.append(row)
        for col_idx in range(1, len(row) + 1):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.border = thin_border
            if col_idx in [1, 2, 4, 5, 9, 10]:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx in [6, 7, 8]:
                cell.alignment = Alignment(
                    horizontal="right", vertical="center"
                )
                if col_idx == 6:
                    cell.number_format = (
                        "#,##0.000" if row[5] < 1 else "#,##0.0"
                    )
                elif col_idx == 7:
                    cell.number_format = "$#,##0.000"
                else:
                    cell.number_format = "$#,##0.00"

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, 50), 12)

    try:
        wb.save(outfile)
        print(f"\n🎉 Файл успешно сохранен: {outfile}")
    except PermissionError:
        fallback_name = (
            outfile.replace(".xlsx", "")
            + f"_{datetime.now().strftime('%H%M%S')}.xlsx"
        )
        wb.save(fallback_name)
        print(
            f"\n⚠️ Файл '{outfile}' сейчас открыт в Excel!"
            f"\n👉 Результат сохранен в: {fallback_name}"
        )


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--address", type=str, required=True)
    parser.add_argument("--apikey", type=str, required=True)
    parser.add_argument(
        "--out", type=str, default="vlabs_all_onchain_trades.xlsx"
    )
    args = parser.parse_args()

    fetch_onchain_history(args.address, args.apikey, args.out)