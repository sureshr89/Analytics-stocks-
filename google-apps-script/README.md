# Google Sheets setup

1. Create a Google Sheet.
2. Create a Google Drive folder for EOD files.
3. Put the Sheet ID and Drive folder ID into `Code.gs`.
4. Open Extensions → Apps Script and paste `Code.gs`.
5. Run `setupTradingJournal()` once and authorize it.
6. Run `installHourlyTrigger()` once.
7. Drop daily EOD files into the Drive folder.

The current starter includes safe duplicate detection and analytics aggregation. A broker-specific parser adapter should be added for the exact EOD workbook/contract-note formats you use; Excel/CSV is preferred for deterministic extraction.
