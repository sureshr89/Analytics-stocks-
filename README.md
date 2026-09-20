# Trading Journal Automation

Google Sheets + EOD trading-report automation for Stocks, Futures and Options.

## Workflow
1. Upload EOD Excel/CSV/PDF to a configured Google Drive folder.
2. Google Apps Script detects new files.
3. The parser imports normalized trades and charges.
4. Duplicate reports are skipped safely.
5. Dashboard sheets recalculate automatically.

## Initial implementation
- Google Apps Script backend
- Raw Imports, Trades, Charges, Daily Summary, Strategy Analytics, Dashboard Data
- Broker-reported P&L reconciliation
- Duplicate-file protection
- Anomaly/spike flags

See `google-apps-script/Code.gs` and `google-apps-script/README.md`.
