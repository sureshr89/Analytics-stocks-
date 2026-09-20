const CONFIG = {
  RAW_FOLDER_ID: 'PUT_GOOGLE_DRIVE_FOLDER_ID_HERE',
  SPREADSHEET_ID: 'PUT_GOOGLE_SHEET_ID_HERE',
  TIMEZONE: Session.getScriptTimeZone() || 'Asia/Kolkata'
};

function setupTradingJournal() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheets = {
    'Raw Imports': ['import_id','file_name','report_date','imported_at','status','broker_net_pnl','calculated_net_pnl','difference','notes'],
    'Trades': ['trade_id','report_date','symbol','segment','instrument','expiry','strike','option_type','side','quantity','entry_price','exit_price','gross_pnl','charges','net_pnl','strategy','setup','notes','source_import_id'],
    'Charges': ['trade_id','report_date','brokerage','stt','exchange_txn','sebi','stamp_duty','gst','ipft','other','total_charges','source_import_id'],
    'Daily Summary': ['date','gross_pnl','charges','net_pnl','trades','wins','losses','win_rate','profit_factor','avg_win','avg_loss'],
    'Strategy Analytics': ['strategy','trades','wins','losses','win_rate','gross_pnl','charges','net_pnl','avg_trade','profit_factor'],
    'Dashboard Data': ['metric','value','updated_at']
  };
  Object.entries(sheets).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    else sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  rebuildAnalytics();
}

function processNewEodFiles() {
  const folder = DriveApp.getFolderById(CONFIG.RAW_FOLDER_ID);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (!/\\.(xlsx|xls|csv|pdf)$/i.test(file.getName())) continue;
    importEodFile_(file);
  }
  rebuildAnalytics();
}

function importEodFile_(file) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const raw = ss.getSheetByName('Raw Imports');
  const importId = Utilities.base64EncodeWebSafe(file.getId());
  const existing = raw.createTextFinder(importId).findNext();
  if (existing) return; // duplicate protection

  // Excel/PDF extraction should be performed by the ingestion adapter.
  // This starter safely records the source file and leaves parsed rows to the adapter.
  raw.appendRow([importId,file.getName(),'',new Date(),'RECEIVED','','','','Awaiting parser adapter']);
}

function rebuildAnalytics() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const trades = ss.getSheetByName('Trades');
  const daily = ss.getSheetByName('Daily Summary');
  const strategy = ss.getSheetByName('Strategy Analytics');
  if (!trades || trades.getLastRow() < 2) return;

  const rows = trades.getRange(2,1,trades.getLastRow()-1,trades.getLastColumn()).getValues();
  const byDate = {}, byStrategy = {};
  rows.forEach(r => {
    const date = r[1], pnl = Number(r[14]) || 0, charges = Number(r[13]) || 0;
    const strat = r[15] || 'Unclassified';
    const d = String(date);
    byDate[d] ||= {gross:0,charges:0,net:0,trades:0,wins:0,losses:0,winPnl:0,lossPnl:0};
    const x = byDate[d]; x.gross += pnl + charges; x.charges += charges; x.net += pnl; x.trades++;
    if (pnl > 0) { x.wins++; x.winPnl += pnl; } else if (pnl < 0) { x.losses++; x.lossPnl += pnl; }
    byStrategy[strat] ||= {trades:0,wins:0,losses:0,gross:0,charges:0,net:0,winPnl:0,lossPnl:0};
    const s = byStrategy[strat]; s.trades++; s.gross += pnl + charges; s.charges += charges; s.net += pnl;
    if (pnl > 0) {s.wins++; s.winPnl += pnl;} else if (pnl < 0) {s.losses++; s.lossPnl += pnl;}
  });

  if (daily.getLastRow() > 1) daily.getRange(2,1,daily.getLastRow()-1,daily.getLastColumn()).clearContent();
  Object.entries(byDate).sort().forEach(([date,x]) => daily.appendRow([date,x.gross,x.charges,x.net,x.trades,x.wins,x.losses,x.trades?x.wins/x.trades:0,x.lossPnl?x.winPnl/Math.abs(x.lossPnl):0,x.wins?x.winPnl/x.wins:0,x.losses?x.lossPnl/x.losses:0]));

  if (strategy.getLastRow() > 1) strategy.getRange(2,1,strategy.getLastRow()-1,strategy.getLastColumn()).clearContent();
  Object.entries(byStrategy).sort().forEach(([name,x]) => strategy.appendRow([name,x.trades,x.wins,x.losses,x.trades?x.wins/x.trades:0,x.gross,x.charges,x.net,x.trades?x.net/x.trades:0,x.lossPnl?x.winPnl/Math.abs(x.lossPnl):0]));
}

function installHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processNewEodFiles').timeBased().everyHours(1).create();
}
