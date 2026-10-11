import os
import sys
import time
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional

import requests
from openpyxl import Workbook, load_workbook

print("[*] Starting Pump.fun Advanced Indexer Movers Bot...")

# ==========================================
# CONFIGURATION & TRADING RULES
# ==========================================
MIN_TRADERS = 10
MIN_MCAP_USD = 12000
MAX_MCAP_USD = 35000

TRADE_SIZE_SOL = 0.05
TAKE_PROFIT_PCT = 0.50
STOP_LOSS_PCT = 0.20
MAX_HOLD_SECONDS = 900
POLL_INTERVAL_SECONDS = 5

INDEXER_URL = "https://advanced-indexer.pump.fun/boards/movers"
INDEXER_PARAMS = {
    "tier": "web",
    "surface": "WEB",
    "platform": "WEB",
    "limit": 150,
    "userId": "F62jhbtofxZJnVDXvv1DgwLmQUZRLgK9UvZn317D9W4i",
    "pumpUserId": "6e79099c-0da0-4726-9528-310de70b719b",
    "session_id": "pump_session_mv3n2a9v_6o4u1k6d2b101a52146t0t2j",
    "device_id": "device-7c932531-619d-4acc-a869-72d6d923e694",
    "session_start_timestamp_ms": "1791711998851",
    "chains": "solana,eip155:4663,eip155:8453,eip155:56,eip155:1,eip155:5042",
}

EXCEL_FILE = "pump_trading_simulation_results.xlsx"
ACTIVE_POSITIONS: Dict[str, Dict[str, Any]] = {}
SEEN_TOKENS = set()


def get_default_headers() -> Dict[str, str]:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Origin": "https://pump.fun",
        "Referer": "https://pump.fun/",
        "Sec-Fetch-Site": "same-site",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Dest": "empty",
        "Connection": "keep-alive",
    }


def init_excel() -> None:
    """Ensure the Excel workbook and Trade_Log sheet exist."""
    try:
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            summary = wb.active
            summary.title = "Summary Dashboard"
            summary["A1"] = "Pump.fun Advanced Indexer Simulation & Audit"
            summary["A2"] = "Automated Paper Trading Log | Target MCap: $12k-$35k"
            summary["A4"] = "Last refreshed:"
            summary["A5"] = "Status: Ready"

            trade_log = wb.create_sheet(title="Trade_Log")
            trade_log.append(
                [
                    "Trade ID",
                    "Timestamp",
                    "Token Mint",
                    "Symbol",
                    "Entry MCap ($)",
                    "Exit MCap ($)",
                    "Peak MCap ($)",
                    "Position (SOL)",
                    "Price Change (%)",
                    "Simulated PnL (SOL)",
                    "Exit Reason",
                ]
            )
            wb.save(EXCEL_FILE)
            print(f"[✓] Initialized fresh Excel file: {EXCEL_FILE}")
        else:
            wb = load_workbook(EXCEL_FILE)
            if "Trade_Log" not in wb.sheetnames:
                ws = wb.create_sheet(title="Trade_Log")
                ws.append(
                    [
                        "Trade ID",
                        "Timestamp",
                        "Token Mint",
                        "Symbol",
                        "Entry MCap ($)",
                        "Exit MCap ($)",
                        "Peak MCap ($)",
                        "Position (SOL)",
                        "Price Change (%)",
                        "Simulated PnL (SOL)",
                        "Exit Reason",
                    ]
                )
                wb.save(EXCEL_FILE)
            print(f"[✓] Existing Excel audit file detected: {EXCEL_FILE}")
    except Exception as exc:  # pragma: no cover - operational safeguard
        print(f"[!] Warning: Could not initialize Excel file: {exc}")


def record_trade_to_excel(trade_data: Dict[str, Any]) -> None:
    """Append a completed simulation trade to the Excel audit log."""
    try:
        if not os.path.exists(EXCEL_FILE):
            init_excel()

        wb = load_workbook(EXCEL_FILE)
        if "Trade_Log" not in wb.sheetnames:
            ws = wb.create_sheet(title="Trade_Log")
            ws.append(
                [
                    "Trade ID",
                    "Timestamp",
                    "Token Mint",
                    "Symbol",
                    "Entry MCap ($)",
                    "Exit MCap ($)",
                    "Peak MCap ($)",
                    "Position (SOL)",
                    "Price Change (%)",
                    "Simulated PnL (SOL)",
                    "Exit Reason",
                ]
            )
        else:
            ws = wb["Trade_Log"]

        next_trade_id = ws.max_row
        row = [
            next_trade_id,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            str(trade_data["mint"])[:8] + "...pump",
            trade_data["symbol"],
            trade_data["entry_mcap"],
            trade_data["exit_mcap"],
            trade_data["peak_mcap"],
            trade_data["size_sol"],
            trade_data["pnl_pct"],
            trade_data["pnl_sol"],
            trade_data["reason"],
        ]
        ws.append(row)
        wb.save(EXCEL_FILE)
        print(f"[📊] Recorded trade for {trade_data['symbol']} to {EXCEL_FILE}")
    except Exception as exc:
        print(f"[!] Error writing to Excel file: {exc}")


def fetch_movers_board() -> List[Dict[str, Any]]:
    """Query the exact Pump.fun Advanced Indexer Movers API using browser-like headers."""
    headers = get_default_headers()
    try:
        response = requests.get(
            INDEXER_URL,
            params=INDEXER_PARAMS,
            headers=headers,
            timeout=10,
        )
        if response.status_code != 200:
            print(f"[!] Indexer returned HTTP status: {response.status_code}")
            return []

        data = response.json()
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for key in ("rows", "data", "coins", "results", "movers"):
                value = data.get(key)
                if isinstance(value, list):
                    return value
            return []

        return []
    except Exception as exc:
        print(f"[!] Fetch error: {exc}")
        return []


def extract_metric(coin: Dict[str, Any], *keys: str, default: float = 0.0) -> float:
    for key in keys:
        if key in coin and coin[key] is not None:
            try:
                return float(coin[key])
            except (TypeError, ValueError):
                pass
    return float(default)


def normalize_coin(coin: Dict[str, Any]) -> Dict[str, Any]:
    norm = dict(coin)
    for key in ["mint", "address", "id"]:
        if key in norm and norm[key] is not None:
            norm["mint"] = norm[key]
            break
    return norm


def evaluate_and_trade(coins: Iterable[Dict[str, Any]]) -> None:
    """Scan indexer data for candidates and manage open positions."""
    current_market_data: Dict[str, Dict[str, Any]] = {}

    for raw_coin in coins:
        if not isinstance(raw_coin, dict):
            continue
        coin = normalize_coin(raw_coin)
        mint = str(coin.get("mint") or coin.get("address") or coin.get("id") or "").strip()
        if not mint:
            continue

        symbol = str(coin.get("symbol") or coin.get("name") or "UNKNOWN").strip() or "UNKNOWN"
        mcap_usd = extract_metric(
            coin,
            "usd_market_cap",
            "usdMarketCap",
            "market_cap",
            "marketCap",
            "mcap_usd",
            "mcapUsd",
            default=0.0,
        )
        tx_count = extract_metric(
            coin,
            "tx_count",
            "txCount",
            "num_transactions",
            "transactions",
            "reply_count",
            "activity_index",
            "activityIndex",
            "traders",
            default=0.0,
        )

        current_market_data[mint] = {"symbol": symbol, "mcap_usd": mcap_usd}

        if mint not in ACTIVE_POSITIONS and mint not in SEEN_TOKENS:
            if MIN_MCAP_USD <= mcap_usd <= MAX_MCAP_USD and tx_count >= MIN_TRADERS:
                SEEN_TOKENS.add(mint)
                simulate_buy(mint=mint, symbol=symbol, mcap=mcap_usd)

    for mint in list(ACTIVE_POSITIONS.keys()):
        if mint in current_market_data:
            monitor_and_simulate_sell(mint, current_market_data[mint]["mcap_usd"])


def simulate_buy(mint: str, symbol: str, mcap: float) -> None:
    """Open a simulated paper-trading position."""
    ACTIVE_POSITIONS[mint] = {
        "symbol": symbol,
        "entry_mcap": float(mcap),
        "peak_mcap": float(mcap),
        "amount_sol": TRADE_SIZE_SOL,
        "entry_time": time.time(),
    }

    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🟢 [INDEXER BUY SIMULATION]")
    print(f" -> Token : {symbol} ({mint[:8]}...)")
    print(f" -> Entry MCap : ${float(mcap):,.2f}")
    print(f" -> Size : {TRADE_SIZE_SOL} SOL")
    print(f" -> Active Positions Open: {len(ACTIVE_POSITIONS)}\n")


def monitor_and_simulate_sell(mint: str, current_mcap: float) -> None:
    """Monitor price changes and trigger take-profit, stop-loss, trailing stop, or timeout exit."""
    if current_mcap <= 0:
        return

    pos = ACTIVE_POSITIONS[mint]
    entry_mcap = float(pos["entry_mcap"])
    symbol = pos["symbol"]

    if current_mcap > pos["peak_mcap"]:
        pos["peak_mcap"] = float(current_mcap)

    pnl_pct = (current_mcap - entry_mcap) / entry_mcap if entry_mcap else 0.0
    peak_drawdown = (pos["peak_mcap"] - current_mcap) / pos["peak_mcap"] if pos["peak_mcap"] else 0.0
    hold_time = time.time() - pos["entry_time"]

    exit_reason = None
    if pnl_pct >= TAKE_PROFIT_PCT:
        exit_reason = f"🚀 TAKE PROFIT HIT (+{pnl_pct * 100:.1f}%)"
    elif pnl_pct <= -STOP_LOSS_PCT:
        exit_reason = f"🛑 STOP LOSS HIT ({pnl_pct * 100:.1f}%)"
    elif pos["peak_mcap"] > entry_mcap * 1.25 and peak_drawdown >= 0.15:
        exit_reason = (
            f"📉 TRAILING STOP HIT (-15% from Peak ${pos['peak_mcap']:,.0f})"
        )
    elif hold_time >= MAX_HOLD_SECONDS:
        exit_reason = f"⏱ TIME EXPIRED (Held > 15m | PnL: {pnl_pct * 100:+.1f}%)"

    if exit_reason:
        pnl_sol = pos["amount_sol"] * pnl_pct
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🔴 [SELL SIMULATION EXECUTED]")
        print(f" -> Reason : {exit_reason}")
        print(f" -> Token : {symbol}")
        print(f" -> Entry MCap : ${entry_mcap:,.2f} -> Exit MCap: ${current_mcap:,.2f}")
        print(f" -> Simulated PnL: {pnl_sol:+.3f} SOL\n")

        trade_summary = {
            "mint": mint,
            "symbol": symbol,
            "entry_mcap": entry_mcap,
            "exit_mcap": current_mcap,
            "peak_mcap": pos["peak_mcap"],
            "size_sol": pos["amount_sol"],
            "pnl_pct": pnl_pct,
            "pnl_sol": pnl_sol,
            "reason": exit_reason,
        }
        record_trade_to_excel(trade_summary)
        del ACTIVE_POSITIONS[mint]


def main() -> None:
    init_excel()
    print("=" * 60)
    print("🚀 PUMP.FUN ADVANCED INDEXER MOVERS BOT")
    print(f"[*] Rules & Targets:")
    print(f"    - Target Market Cap : ${MIN_MCAP_USD:,} - ${MAX_MCAP_USD:,}")
    print(f"    - Min Activity Index: > {MIN_TRADERS} transactions")
    print(f"    - Take Profit Target: +{int(TAKE_PROFIT_PCT * 100)}%")
    print(f"    - Stop Loss Limit   : -{int(STOP_LOSS_PCT * 100)}%")
    print(f"    - Excel Audit Log   : {EXCEL_FILE}")
    print("=" * 60 + "\n")

    while True:
        try:
            coins = fetch_movers_board()
            if coins:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] [*] Fetched {len(coins)} movers. Evaluating rules...")
                evaluate_and_trade(coins)
            else:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] [*] Polling movers board...")

            time.sleep(POLL_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            print("\n[*] Bot stopped by user.")
            sys.exit(0)
        except Exception as exc:
            print(f"[!] Error in main loop: {exc}")
            time.sleep(POLL_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
