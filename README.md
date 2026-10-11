# pump-fun-trading-bot

Automated paper-trading bot for Pump.fun that polls the Advanced Indexer Movers API every 5 seconds, screens for qualifying tokens, simulates paper entries/exits, and logs trades to an Excel audit workbook.

## Features

- Uses the exact Pump.fun Advanced Indexer `boards/movers` endpoint with the required browser-style headers and query params
- Polls every 5 seconds
- Filters by configurable market cap and activity threshold
- Simulates paper trades with 0.05 SOL position sizing
- Applies take-profit, stop-loss, trailing-stop, and 15-minute exit rules
- Produces a live Excel workbook with a `Trade_Log` sheet for audit tracking

## Requirements

```bash
pip install -r requirements.txt
```

## Run

```bash
python pump_bot.py
```

## Output

The script writes a workbook named `pump_trading_simulation_results.xlsx` in the current working directory.

## Notes

This script is a simulation and does not place real orders. It is designed for paper-trading and audit logging only.
