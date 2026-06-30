// ── CONFIG ────────────────────────────────────────────────────────────────────
const GEMINI_API_KEY = "AIzaSyC8pvMI0EcqvxXKWn9HoieM7mBuW92w5kU";
const GEMINI_URL     = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + GEMINI_API_KEY;
const ITRUST_API     = "https://api.itrust.co.tz/api/market/summary";
const HISTORY_SHEET  = "Historical_Log";
const DIV_SHEET      = "Dividends_Table";
const PRICES_SHEET   = "Current_Prices";   // ← new dedicated sheet
const MY_TICKERS     = ["CRDB", "NMB", "KCB", "NICO", "AFRIPRISE",
                        "DSE", "DCB", "MKCB", "SWIS", "IEACLC-ETF"];
const REFERER = "https://yohanaraphael19.github.io/wealthledger/"; // for API key referrer restrictions


// ── MASTER TRIGGER ────────────────────────────────────────────────────────────
function dailyDSEUpdate() {
  snapshotFromITrust();       // 1. log OHLC to Historical_Log
  updateCurrentPricesSheet(); // 2. update Current_Prices sheet
  updateStockHoldingsPrices();// 3. push latest prices into Stock Holdings col L
  update52WHighLow();         // 4. refresh 52-week high/low from history
  runAIAnalysis();            // 5. AI daily report
}


// ── STEP 2: UPDATE Current_Prices SHEET ──────────────────────────────────────
// This is the single source of truth for latest prices.
// Stock Holdings col L (Current Price) reads from here via VLOOKUP or this function.
function updateCurrentPricesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PRICES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PRICES_SHEET);
    sheet.getRange("A1:D1").setValues([["Ticker", "Company", "Price (TZS)", "Last Updated"]]);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#1A5F5F").setFontColor("#ffffff");
  }

  let response;
  try {
    response = UrlFetchApp.fetch(ITRUST_API);
  } catch(e) {
    console.error("Current_Prices: iTrust fetch failed: " + e.message);
    return;
  }

  const parsed = JSON.parse(response.getContentText());
  const data   = Array.isArray(parsed) ? parsed :
                 parsed.data || parsed.stocks || parsed.summary ||
                 parsed.result || Object.values(parsed)[0];

  const priceMap = {};
  data.forEach(stock => {
    const ticker = String(stock.Symbol || "").trim().toUpperCase();
    if (ticker && stock.Close > 0) {
      priceMap[ticker] = { company: stock.Name || stock.CompanyName || ticker, price: stock.Close };
    }
  });

  // Collect ALL tickers we care about: MY_TICKERS + any ticker in Stock Holdings
  const stockSheet   = ss.getSheetByName("Stock Holdings");
  const holdingTicks = stockSheet
    ? stockSheet.getRange(5, 4, 50, 1).getValues().flat()
        .map(t => String(t).trim().toUpperCase()).filter(t => t)
    : [];
  const allTickers   = [...new Set([...MY_TICKERS, ...holdingTicks])];

  const now       = new Date();
  const rows      = [];
  allTickers.forEach(tk => {
    if (priceMap[tk]) {
      rows.push([tk, priceMap[tk].company, priceMap[tk].price, now]);
    }
  });

  if (rows.length === 0) { console.warn("Current_Prices: no prices received."); return; }

  // Rewrite data rows (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 180);
  console.log("Current_Prices: updated " + rows.length + " tickers.");
}


// ── STEP 3: PUSH PRICES INTO STOCK HOLDINGS COL L ────────────────────────────
// Reads from Current_Prices sheet so new tickers are picked up automatically.
function updateStockHoldingsPrices() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Stock Holdings");
  const priceSheet = ss.getSheetByName(PRICES_SHEET);

  if (!priceSheet) {
    console.warn("Current_Prices sheet not found — run updateCurrentPricesSheet() first.");
    return;
  }

  // Build price map from Current_Prices sheet
  const priceData = priceSheet.getRange(2, 1, priceSheet.getLastRow() - 1, 3).getValues();
  const priceMap  = {};
  priceData.forEach(row => {
    const ticker = String(row[0]).trim().toUpperCase();
    if (ticker && row[2]) priceMap[ticker] = row[2];
  });

  // Scan Stock Holdings rows 5–54 (col D = Ticker, col L = Current Price = col 12)
  const lastDataRow = stockSheet.getLastRow();
  const maxRows     = Math.max(0, Math.min(lastDataRow - 4, 50));
  if (maxRows === 0) return;

  const tickerRange = stockSheet.getRange(5, 4, maxRows, 1).getValues();
  tickerRange.forEach((row, i) => {
    const ticker = String(row[0]).trim().toUpperCase();
    if (ticker && priceMap[ticker] !== undefined) {
      stockSheet.getRange(5 + i, 12).setValue(priceMap[ticker]); // col L = 12
      console.log("Stock Holdings: " + ticker + " → " + priceMap[ticker]);
    }
  });
}


// ── FORMULA AUTO-FILL HELPER ──────────────────────────────────────────────────
// After insertRowBefore(5), the new row 5 is blank.
// The old row 5 (with formulas) is now at row 6.
// We copy formulas from row 6 → row 5 for the specified columns.
function copyFormulasDown(sheet, fromRow, toRow, startCol, endCol) {
  for (let col = startCol; col <= endCol; col++) {
    const sourceCell = sheet.getRange(fromRow, col);
    const formula    = sourceCell.getFormula();
    if (formula) {
      // Shift all row references up by (fromRow - toRow) to point to correct row
      const shift      = fromRow - toRow;
      const adjusted   = shiftFormulaRows(formula, shift);
      sheet.getRange(toRow, col).setFormula(adjusted);
    }
  }
}

// Shift numeric row references inside a formula by -shift
// e.g. shiftFormulaRows("=E6*F6", 1) → "=E5*F5"
function shiftFormulaRows(formula, shift) {
  // Replace all cell references like A6, Z99, etc.
  return formula.replace(/([A-Z]+)(\d+)/g, (match, col, rowStr) => {
    const row = parseInt(rowStr, 10);
    return col + (row - shift);
  });
}


// ── LOG STOCK BUY (with formula auto-fill) ────────────────────────────────────
function logStockBuy(control_number, ticker, company, qty, price) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Stock Holdings");

  // Insert new row at row 5
  sheet.insertRowBefore(5);

  // Fill in the blue (user-input) columns: A=1, B=2, C=3, D=4, E=5, F=6
  const dateStr = Utilities.formatDate(new Date(), "Africa/Dar_es_Salaam", "dd-MMM-yyyy");
  sheet.getRange(5, 1).setValue(control_number || "");
  sheet.getRange(5, 2).setValue(dateStr);
  sheet.getRange(5, 3).setValue(company || "");
  sheet.getRange(5, 4).setValue(ticker  || "");
  sheet.getRange(5, 5).setValue(Number(qty)   || 0);
  sheet.getRange(5, 6).setValue(Number(price) || 0);

  // Copy formulas from row 6 → row 5 for auto-calculated columns G through O (7–15)
  copyFormulasDown(sheet, 6, 5, 7, 15);

  // Also try to fill current price from Current_Prices sheet (col L = 12)
  const priceSheet = ss.getSheetByName(PRICES_SHEET);
  if (priceSheet) {
    const priceData = priceSheet.getRange(2, 1, Math.max(1, priceSheet.getLastRow()-1), 3).getValues();
    priceData.forEach(row => {
      if (String(row[0]).trim().toUpperCase() === String(ticker).trim().toUpperCase() && row[2]) {
        sheet.getRange(5, 12).setValue(row[2]);
      }
    });
  }

  console.log("Buy logged: " + qty + " " + ticker + " @ " + price);
  return "Buy Trade Logged: " + qty + " " + ticker + " @ TZS " + Number(price).toLocaleString();
}


// ── LOG STOCK SELL (with formula auto-fill) ───────────────────────────────────
function logStockSell(ticker, company, qty, price, sell_date) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sold Trades");

  sheet.insertRowBefore(5);

  const dateStr = sell_date ||
                  Utilities.formatDate(new Date(), "Africa/Dar_es_Salaam", "dd-MMM-yyyy");
  sheet.getRange(5, 1).setValue(dateStr);
  sheet.getRange(5, 2).setValue(company || "");
  sheet.getRange(5, 3).setValue(ticker  || "");
  sheet.getRange(5, 4).setValue(Number(qty)   || 0);
  sheet.getRange(5, 5).setValue(Number(price) || 0);

  // Copy formulas from row 6 → row 5 for cols F–H (6–8): Broker Fee, True Sell Price, Net Proceeds
  copyFormulasDown(sheet, 6, 5, 6, 8);

  console.log("Sell logged: " + qty + " " + ticker + " @ " + price);
  return "Sell Trade Logged: " + qty + " " + ticker + " @ TZS " + Number(price).toLocaleString();
}


// ── LOG LAND BUY ──────────────────────────────────────────────────────────────
function logLandBuy(block, plot_no, region, district, street, area_sqm, date_purchased, title_deed_status, purchase_price, proc_fee) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Land & Plots");

  const dateStr       = date_purchased ||
                        Utilities.formatDate(new Date(), "Africa/Dar_es_Salaam", "dd-MMM-yyyy");
  const purchasePrice = Number(purchase_price) || 0;
  const procFeeVal    = Number(proc_fee)        || 0;
  const trueCost      = purchasePrice + procFeeVal;

  sheet.insertRowBefore(5);
  sheet.getRange(5, 1).setValue(block    || "");
  sheet.getRange(5, 2).setValue(plot_no  || "");
  sheet.getRange(5, 3).setValue(region   || "");
  sheet.getRange(5, 4).setValue(district || "");
  sheet.getRange(5, 5).setValue(street   || "");
  sheet.getRange(5, 6).setValue(area_sqm || "");
  sheet.getRange(5, 7).setValue(dateStr);
  sheet.getRange(5, 8).setValue(title_deed_status || "Pending Title Deed");
  sheet.getRange(5, 9).setValue(purchasePrice);
  sheet.getRange(5, 10).setValue(procFeeVal);
  sheet.getRange(5, 11).setValue(trueCost);
  sheet.getRange(5, 12).setValue(trueCost); // Current Est. Value starts at true cost

  // Copy any remaining formulas (cols 13+) from row 6
  copyFormulasDown(sheet, 6, 5, 13, 16);

  console.log("Land logged: " + street + ", " + region);
  return "Land/Plot Logged: " + street + ", " + region + " — TZS " + trueCost.toLocaleString();
}


// ── FETCH FROM ITRUST API AND SNAPSHOT ───────────────────────────────────────
function snapshotFromITrust() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(HISTORY_SHEET);
  const divSheet = ss.getSheetByName(DIV_SHEET);

  const divData = divSheet.getRange("A2:B20").getValues();
  const divMap  = {};
  divData.forEach(row => {
    const tk  = String(row[0]).trim().toUpperCase();
    const div = row[1];
    if (tk && typeof div === 'number' && div >= 0) divMap[tk] = div;
  });

  let response;
  try {
    response = UrlFetchApp.fetch(ITRUST_API);
  } catch(e) {
    console.error("Failed to fetch iTrust API: " + e.message);
    return;
  }

  const parsed = JSON.parse(response.getContentText());
  const data   = Array.isArray(parsed) ? parsed :
                 parsed.data || parsed.stocks ||
                 parsed.summary || parsed.result ||
                 Object.values(parsed)[0];

  const today      = new Date();
  const rowsToLog  = [];

  data.forEach(stock => {
    const ticker = String(stock.Symbol).trim().toUpperCase();
    if (!MY_TICKERS.includes(ticker)) return;
    const close = stock.Close || 0;
    if (close <= 0) return;
    const dividend = divMap[ticker] !== undefined ? divMap[ticker] : 0;
    const yieldPct = dividend > 0 && close > 0
                     ? parseFloat(((dividend / close) * 100).toFixed(2)) : 0;
    rowsToLog.push([
      today, ticker,
      stock.Open || 0, close, stock.High || 0, stock.Low || 0,
      stock.Change || 0, stock.Turnover || 0, dividend, yieldPct
    ]);
  });

  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, 10).setValues([[
      'Date','Ticker','Open','Close','High','Low','Change (%)','Turnover','Dividend (TZS)','Yield (%)'
    ]]);
    logSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1A5F5F').setFontColor('#ffffff');
  }

  if (rowsToLog.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, rowsToLog.length, 10).setValues(rowsToLog);
    console.log("Logged " + rowsToLog.length + " tickers.");
  } else {
    console.warn("No valid data from iTrust API.");
  }
}


// ── GENERATE SIMULATION DATA ──────────────────────────────────────────────────
function generateSimulationData() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(HISTORY_SHEET);
  const divSheet = ss.getSheetByName(DIV_SHEET);

  const divData = divSheet.getRange("A2:B20").getValues();
  const divMap  = {};
  divData.forEach(row => {
    const tk  = String(row[0]).trim().toUpperCase();
    const div = row[1];
    if (tk && typeof div === 'number' && div >= 0) divMap[tk] = div;
  });

  const basePrice  = { "CRDB":3020,"NMB":14260,"KCB":1790,"NICO":3690,"AFRIPRISE":830,"DSE":6670,"DCB":820,"MKCB":5000,"SWIS":2580,"IEACLC-ETF":1280 };
  const volatility = { "CRDB":0.012,"NMB":0.008,"KCB":0.010,"NICO":0.015,"AFRIPRISE":0.018,"DSE":0.014,"DCB":0.020,"MKCB":0.016,"SWIS":0.012,"IEACLC-ETF":0.006 };

  const totalDays  = 126;
  const allRows    = [];
  const today      = new Date();
  today.setHours(0, 0, 0, 0);

  const startPrice = {};
  MY_TICKERS.forEach(tk => {
    startPrice[tk] = basePrice[tk] / Math.pow(1.0003, totalDays);
  });

  const tradingDates = [];
  let d = new Date(today);
  d.setDate(d.getDate() - 1);
  while (tradingDates.length < totalDays) {
    const day = d.getDay();
    if (day !== 0 && day !== 6) tradingDates.unshift(new Date(d));
    d.setDate(d.getDate() - 1);
  }

  MY_TICKERS.forEach(tk => {
    let price = startPrice[tk];
    const vol = volatility[tk] || 0.012;
    tradingDates.forEach(date => {
      price = price * (1 + (Math.random() - 0.48) * vol * 2);
      const close    = Math.round(price / 10) * 10;
      const open     = Math.round(close * (1 + (Math.random() - 0.5) * vol) / 10) * 10;
      const high     = Math.round(Math.max(open, close) * (1 + Math.random() * vol * 0.5) / 10) * 10;
      const low      = Math.round(Math.min(open, close) * (1 - Math.random() * vol * 0.5) / 10) * 10;
      const change   = parseFloat(((close - open) / open * 100).toFixed(2));
      const turnover = Math.round(close * (Math.random() * 50000 + 5000));
      const dividend = divMap[tk] !== undefined ? divMap[tk] : 0;
      const yieldPct = dividend > 0 && close > 0 ? parseFloat(((dividend / close) * 100).toFixed(2)) : 0;
      allRows.push([date, tk, open, close, high, low, change, turnover, dividend, yieldPct]);
    });
  });

  allRows.sort((a, b) => a[0] - b[0] || a[1].localeCompare(b[1]));
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) logSheet.getRange(2, 1, lastRow - 1, 10).clearContent();
  logSheet.getRange(2, 1, allRows.length, 10).setValues(allRows);
  console.log("Generated " + allRows.length + " simulation rows.");
}


// ── HELPER: Build rich portfolio context from source sheets ───────────────────
function buildPortfolioContext() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const holdings = ss.getSheetByName("Stock Holdings");
  const sold     = ss.getSheetByName("Sold Trades");
  const land     = ss.getSheetByName("Land & Plots");
  const history  = ss.getSheetByName("Historical_Log");

  const holdingsData = holdings.getRange(5, 1, 38, 14).getValues().filter(r => r[3]);
  const soldData     = sold.getRange(5, 1, 30, 8).getValues().filter(r => r[2]);
  const landData     = land.getRange(5, 1, 6, 13).getValues().filter(r => r[2]);

  const totalTrueCost     = holdings.getRange("I44").getValue();
  const totalSoldProceeds = sold.getRange("H36").getValue();
  const costBasis         = totalTrueCost - totalSoldProceeds;
  const landValue         = land.getRange("L29").getValue();
  const stockValue        = holdingsData.reduce((sum, r) => sum + (Number(r[12]) || 0), 0);
  const totalNetWorth     = stockValue + landValue;
  const unrealizedPL      = stockValue - costBasis;

  const allHoldingTickers = holdings.getRange(5, 4, 38, 1).getValues().flat();
  const allHoldingShares  = holdings.getRange(5, 5, 38, 1).getValues().flat();
  const allSoldTickers    = sold.getRange(5, 3, 30, 1).getValues().flat();
  const allSoldShares     = sold.getRange(5, 4, 30, 1).getValues().flat();

  function sumIfTicker(ticker, tickers, values) {
    return tickers.reduce((sum, t, i) => t === ticker ? sum + (Number(values[i]) || 0) : sum, 0);
  }

  const crdbShares  = sumIfTicker("CRDB", allHoldingTickers, allHoldingShares) - sumIfTicker("CRDB", allSoldTickers, allSoldShares);
  const etfShares   = sumIfTicker("IEACLC-ETF", allHoldingTickers, allHoldingShares) - sumIfTicker("IEACLC-ETF", allSoldTickers, allSoldShares);
  const totalShares = crdbShares + etfShares;

  const holdingsText = holdingsData
    .map(r => `  ${r[1]} | ${r[3]} | ${r[4]} shares @ TZS ${r[5]} | True Cost: TZS ${r[8]} | Current Value: TZS ${r[12]} | P&L: TZS ${r[13]} | Return: ${r[14]}`)
    .join("\n");

  const soldText = soldData.length
    ? soldData.map(r => `  ${r[0]} | ${r[2]} | ${r[3]} shares @ TZS ${r[4]} | Net Proceeds: TZS ${r[7]}`).join("\n")
    : "No sold trades yet";

  const landText = landData
    .map(r => `  ${r[2]}, ${r[3]} | ${r[4]} | ${r[5]} | ${r[6]} | Purchase: TZS ${r[8]} | Proc Fee: TZS ${r[9]} | True Cost: TZS ${r[10]} | Current Value: TZS ${r[11]} | Gain: TZS ${r[12]}`)
    .join("\n");

  const histLastRow = history.getLastRow();
  const histRows    = Math.min(900, histLastRow - 1);
  const histStart   = Math.max(2, histLastRow - histRows + 1);
  const histText    = history.getRange(histStart, 1, histRows, 7).getValues()
    .filter(r => r[1])
    .map(r => `  ${r[0]} | ${r[1]} | O:${r[2]} C:${r[3]} H:${r[4]} L:${r[5]} Chg:${r[6]}%`)
    .join("\n");

  return {
    context:
`=== PORTFOLIO SUMMARY ===
Total Net Worth: TZS ${totalNetWorth.toLocaleString()}
Stock Portfolio Value: TZS ${stockValue.toLocaleString()}
Cost Basis (Stocks): TZS ${costBasis.toLocaleString()}
Unrealized P&L: TZS ${unrealizedPL.toLocaleString()}
Land Portfolio Value: TZS ${landValue.toLocaleString()}
CRDB Shares Held: ${crdbShares} | IEACLC-ETF Shares Held: ${etfShares} | Total: ${totalShares}

=== STOCK HOLDINGS (${holdingsData.length} trades) ===
Date | Ticker | Shares | Buy Price | True Cost | Current Value | P&L | Return
${holdingsText}

=== SOLD TRADES ===
Date | Ticker | Shares | Sell Price | Net Proceeds
${soldText}

=== LAND & PLOTS (${landData.length} plots) ===
Region | District | Street | Size | Date | Purchase Price | Proc Fee | True Cost | Current Value | Gain
${landText}

=== FUNDAMENTALS ===
Ticker | EPS (TZS) | BVPS (TZS) | ROE (%) | 52W High | 52W Low | % from High
${(()=>{try{const f=ss.getSheetByName('Fundamentals');if(!f)return'Not available.';const d=f.getRange(2,1,10,6).getValues().filter(r=>r[0]);if(!d.length)return'No data yet.';return d.map(r=>{const tk=String(r[0]).trim().toUpperCase();const eps=r[1]||'—';const bv=r[2]||'—';const roe=r[3]||'—';const h52=r[4]||'—';const l52=r[5]||'—';return '  '+tk+': EPS=TZS '+eps+' | BVPS=TZS '+bv+' | ROE='+roe+'% | 52W H='+h52+' L='+l52;}).join('\n');}catch(e){return'Error reading Fundamentals.';}})()}

=== MARKET HISTORY (last ${histRows} entries ≈ 3 months) ===
Date | Ticker | Open | Close | High | Low | Change%
${histText}`,
    totalNetWorth, stockValue, costBasis, unrealizedPL,
    landValue, crdbShares, etfShares, totalShares
  };
}


// ── RUN AI ANALYSIS (daily) ───────────────────────────────────────────────────
function runAIAnalysis() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(HISTORY_SHEET);

  const { context, totalNetWorth, stockValue, costBasis,
          unrealizedPL, landValue, crdbShares, etfShares, totalShares } = buildPortfolioContext();

  const allHistory  = logSheet.getDataRange().getValues();
  const rows        = allHistory.slice(1);
  const tickerStats = {};
  rows.forEach(row => {
    const ticker = String(row[1]).trim().toUpperCase();
    const close  = row[3];
    const div    = row[8] || 0;
    if (!ticker || typeof close !== 'number' || close <= 0) return;
    if (!tickerStats[ticker]) tickerStats[ticker] = { prices: [], dividend: div };
    tickerStats[ticker].prices.push(close);
  });

  const summaries = [];
  Object.keys(tickerStats).forEach(tk => {
    const prices = tickerStats[tk].prices;
    const first  = prices[0];
    const last   = prices[prices.length - 1];
    const avg    = Math.round(prices.reduce((a, b) => a + b, 0) / prices.length);
    const trend  = ((last - first) / first * 100).toFixed(1);
    const div    = tickerStats[tk].dividend;
    const yld    = div > 0 && last > 0 ? ((div / last) * 100).toFixed(2) : 0;
    summaries.push(`${tk}: Last=TZS ${last}, 6M-Trend=${trend}%, Yield=${yld}%, Min=${Math.min(...prices)}, Max=${Math.max(...prices)}, Avg=${avg}`);
  });

  const totalReturn = costBasis > 0 ? ((unrealizedPL / costBasis) * 100).toFixed(1) : 0;

  const prompt = `You are a professional financial analyst specializing in the Dar es Salaam Stock Exchange (DSE), Tanzania.

## MY CURRENT PORTFOLIO
- Total Net Worth: TZS ${totalNetWorth.toLocaleString()}
- Stock Portfolio Value: TZS ${stockValue.toLocaleString()}
- Land Portfolio Value: TZS ${landValue.toLocaleString()}
- True Cost Basis (inc. 2.4% broker fees, net of proceeds): TZS ${costBasis.toLocaleString()}
- Unrealized Stock P&L: TZS ${unrealizedPL.toLocaleString()} (${totalReturn}% return)
- CRDB shares held: ${crdbShares.toLocaleString()}
- IEACLC-ETF shares held: ${etfShares.toLocaleString()}
- Total shares held: ${totalShares.toLocaleString()}

## 6-MONTH DSE MARKET DATA
${summaries.join('\n')}

## ANALYSIS REQUESTED

**1. PORTFOLIO HEALTH**
Assess concentration risk, sector exposure, and overall balance. Be specific to DSE/Tanzania context.

**2. 5-YEAR PROJECTION**
Project portfolio value under three scenarios: Conservative (6% CAGR), Moderate (10% CAGR), Optimistic (15% CAGR). Show values in TZS.

**3. TOP BUY OPPORTUNITIES**
From the 10 tracked stocks, recommend 2-3 to add based on trend, yield, and value. Give specific reasons with TZS figures.

**4. DIVIDEND INCOME ANALYSIS**
Calculate estimated annual dividend income from current holdings. Which tracked stocks offer best yield?

**5. RISK ALERTS**
Flag any concerning trends. Note DSE-specific risks (liquidity, sector concentration, currency).

**6. 30-DAY ACTION PLAN**
Give 3 specific, actionable steps for the next 30 days.

Keep analysis concise, use TZS throughout, tailor all advice to the DSE/Tanzania market context.`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.7, maxOutputTokens: 8000 }
  };

  try {
    const response = UrlFetchApp.fetch(GEMINI_URL, {
      method: 'post', contentType: 'application/json',
      headers: { 'Referer': REFERER },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });

    const code   = response.getResponseCode();
    const result = JSON.parse(response.getContentText());
    if (code !== 200) { console.error("Gemini error: " + result.error.message); return; }

    const aiResponse = result.candidates[0].content.parts[0].text;

    let insightSheet = ss.getSheetByName("AI_Insights");
    if (!insightSheet) insightSheet = ss.insertSheet("AI_Insights");
    insightSheet.clearContents();
    insightSheet.getRange("A1").setValue("WealthLedger AI Analysis — Powered by Gemini 2.0 Flash");
    insightSheet.getRange("A2").setValue("Generated: " + new Date().toLocaleString());
    insightSheet.getRange("A1").setFontSize(14).setFontWeight("bold").setFontColor("#1A5F5F");
    insightSheet.getRange("A2").setFontSize(10).setFontColor("#8A8680");
    insightSheet.getRange("A4").setValue(aiResponse).setWrap(true);
    insightSheet.setColumnWidth(1, 900);
    writeAIParsed(aiResponse);
    console.log("AI Analysis complete.");

  } catch(e) {
    console.error("Script error: " + e.message);
  }
}


// ── WRITE STRUCTURED DATA TO AI_Parsed SHEET ─────────────────────────────────
function writeAIParsed(fullText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AI_Parsed');
  if (!sheet) sheet = ss.insertSheet('AI_Parsed');
  sheet.clearContents();

  const healthMatch    = fullText.match(/Status:\s*\*?\*?([^\n*]+)/i);
  const health         = healthMatch ? healthMatch[1].trim() : 'See full report';
  const concMatch      = fullText.match(/Concentration Risk[^:]*:\s*\*?\*?([^\n*]+)/i);
  const healthSub      = concMatch ? concMatch[1].trim().slice(0, 100) : '';
  const buyMatch       = fullText.match(/1\.\s*\*\*([^*(]+)/);
  const buy            = buyMatch ? buyMatch[1].trim() : '—';
  const buyReasonMatch = fullText.match(/Reason[^:]*:\*?\*?\s*([^\n*]+)/i);
  const buySub         = buyReasonMatch ? buyReasonMatch[1].trim().slice(0, 100) : '';
  const actionMatch    = fullText.match(/1\.\s*\*\*([^*:]+)[*:]/);
  const action         = actionMatch ? actionMatch[1].trim() : '—';

  sheet.getRange('A1:B6').setValues([
    ['Generated', new Date().toLocaleString()],
    ['Health',    health],
    ['HealthSub', healthSub],
    ['TopBuy',    buy],
    ['BuySub',    buySub],
    ['Action',    action],
  ]);
  console.log('AI_Parsed written.');
}


// ── doGet (Telegram webhook verification) ────────────────────────────────────
function doGet(e) {
  return ContentService.createTextOutput("WealthLedger Bot is running.");
}


// ── doPost (routes all incoming requests) ────────────────────────────────────
function doPost(e) {
  const contents = e.postData.contents;

  // Route 1: Dashboard log request (NO AI involved — direct sheet write)
  try {
    const parsed = JSON.parse(contents);
    if (parsed.logType) {
      const result = handleDashboardLog(parsed);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // ★ Route 2b: Regenerate AI analysis from frontend ↻ button
    if (parsed.action === 'regenerate_analysis') {
      runAIAnalysis();
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, message: 'Analysis regenerated' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // ★ Route 2c: Extract fundamentals from PDF reports in Drive
    if (parsed.action === 'process_reports') {
      refreshFundamentals();
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, message: 'Fundamentals updated from PDF reports' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {}

  // Route 3: Dashboard chat proxy (uses Gemini)
  const body     = JSON.parse(contents);
  const question = body.question || '';
  const context  = body.context  || '';

  const prompt = `You are a DSE (Dar es Salaam Stock Exchange) financial advisor for a Tanzanian investor.

Portfolio context:
${context}

Investor asks: "${question}"

Rules:
- Plain prose only, no tables, no bullet symbols
- Use TZS values from context
- Match answer length to question complexity: simple questions get 1-2 sentences, analysis gets full paragraphs
- Always finish your sentences completely`;

  try {
    const response = UrlFetchApp.fetch(GEMINI_URL, {
      method: 'post', contentType: 'application/json',
      headers: { 'Referer': REFERER },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 8000 }
      }),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    // ── Quota handling: return structured error so frontend can show reset info
    if (responseCode === 429) {
      return ContentService
        .createTextOutput(JSON.stringify({ answer: null, quotaError: true, code: 429 }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (responseCode !== 200) {
      return ContentService
        .createTextOutput(JSON.stringify({ answer: 'API Error ' + responseCode + ': ' + responseText.slice(0, 200) }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result    = JSON.parse(responseText);
    const candidate = result.candidates?.[0];
    if (!candidate) {
      return ContentService
        .createTextOutput(JSON.stringify({ answer: 'Response blocked: ' + (result.promptFeedback?.blockReason || 'Unknown') }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ answer: candidate.content?.parts?.[0]?.text || 'No text in response.' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ answer: 'Script error: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ── DASHBOARD LOG HANDLER (NO AI — direct sheet writes) ──────────────────────
function handleDashboardLog(body) {
  const type = body.logType;
  try {
    if (type === "BUY") {
      const msg = logStockBuy(
        body.control_number, body.ticker, body.company, body.qty, body.price
      );
      return { ok: true, message: msg };

    } else if (type === "SELL") {
      const msg = logStockSell(
        body.ticker, body.company, body.qty, body.price, body.sell_date
      );
      return { ok: true, message: msg };

    } else if (type === "LAND_BUY") {
      const msg = logLandBuy(
        body.block, body.plot_no, body.region, body.district,
        body.street, body.area_sqm, body.date_purchased,
        body.title_deed_status, body.purchase_price, body.proc_fee
      );
      return { ok: true, message: msg };

    } else if (type === "LAND_SELL") {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("Land & Plots");
      const data  = sheet.getRange(5, 1, 20, 5).getValues();
      let deletedRow = null;
      for (let i = 0; i < data.length; i++) {
        const matchRegion = String(data[i][2]).toLowerCase().includes(body.region.toLowerCase().trim());
        const matchStreet = body.street  ? String(data[i][4]).toLowerCase().includes(body.street.toLowerCase().trim())  : true;
        const matchPlot   = body.plot_no ? String(data[i][1]).toLowerCase() === String(body.plot_no).toLowerCase().trim() : true;
        if (matchRegion && matchStreet && matchPlot) { deletedRow = i + 5; break; }
      }
      if (deletedRow) {
        sheet.deleteRow(deletedRow);
        return { ok: true, message: "Plot sold and removed: " + body.street + ", " + body.region };
      }
      return { ok: false, message: "Plot not found. Check region/street/plot number." };
    }

    return { ok: false, message: "Unknown log type." };

  } catch(err) {
    return { ok: false, message: "Error: " + err.message };
  }
}


// ── EXTRACT FUNDAMENTALS FROM PDF REPORTS IN DRIVE ──────────────────────────
// 1. User uploads PDF reports to the Google Drive folder "WealthLedger Reports"
// 2. Name files like "CRDB_2024_Annual_Report.pdf" so the ticker is in the filename
// 3. Run this function (or trigger from frontend) to extract EPS/BVPS/ROE via Gemini
// 4. Processed PDFs are moved to a "Processed" subfolder
function extractFundamentalsFromPDFs() {
  // Use the shared folder ID directly (viewable at https://drive.google.com/drive/folders/1ddIjc865YlAH0F63vXL_chPTGlUhKuW9)
  let folder;
  try {
    folder = DriveApp.getFolderById("1ddIjc865YlAH0F63vXL_chPTGlUhKuW9");
  } catch(e) {
    console.log("Could not access the shared Drive folder. Check permissions.");
    return { ok: false, message: "Folder access error: " + e.message };
  }
  const files = folder.getFilesByType(MimeType.PDF);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Fundamentals");
  if (!sheet) {
    sheet = ss.insertSheet("Fundamentals");
    sheet.getRange("A1:F1").setValues([["Ticker", "EPS (TZS)", "BVPS (TZS)", "ROE (%)", "52W High", "52W Low"]]);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#1A5F5F").setFontColor("#ffffff");
  }

  let processedFolder;
  const subFolders = folder.getFoldersByName("Processed");
  if (subFolders.hasNext()) { processedFolder = subFolders.next(); }
  else { processedFolder = folder.createFolder("Processed"); }

  const results = {};
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const tickerMatch = fileName.match(/^([A-Z]+[-\w]*)/);
    if (!tickerMatch) { console.log("Skipping " + fileName + ": ticker not found in filename"); continue; }
    const ticker = tickerMatch[1].toUpperCase();

    const blob = file.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());

    const prompt = `You are a financial data extractor for Tanzanian-listed companies. Analyze this annual/quarterly report PDF.
Extract the following data for ticker "${ticker}" in TZS:
1. EPS (Earnings Per Share) in TZS
2. BVPS (Book Value Per Share) in TZS
3. ROE (Return on Equity) as percentage (e.g., 15.2 for 15.2%)

Return ONLY valid JSON, no markdown, no code blocks. Example:
{"eps":550,"bvps":3200,"roe":17.2}
Use null if a value is not found.`;

    try {
      const response = UrlFetchApp.fetch(GEMINI_URL, {
        method: 'post', contentType: 'application/json',
        headers: { 'Referer': REFERER },
        payload: JSON.stringify({
          contents: [{
            parts: [
              { inline_data: { mime_type: "application/pdf", data: base64Data } },
              { text: prompt }
            ]
          }],
          generationConfig: { temperature: 0.1, maxOutputTokens: 2000 }
        }),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      const result = JSON.parse(response.getContentText());
      if (code !== 200) { console.error(`Gemini error for ${fileName}: ${result.error?.message || code}`); continue; }

      const aiText = result.candidates[0].content.parts[0].text;
      const data = JSON.parse(aiText.replace(/```json|```/g, '').trim());
      results[ticker] = { eps: data.eps, bvps: data.bvps, roe: data.roe };
      file.moveTo(processedFolder);
      count++;
      console.log(`Extracted fundamentals for ${ticker} from ${fileName}`);
    } catch(err) {
      console.error(`Error processing ${fileName}: ${err.message}`);
    }
  }

  if (count === 0) { console.log("No new PDFs found to process."); return { ok: true, message: "No new reports" }; }

  // Merge with existing Fundamentals sheet data
  const existingData = sheet.getRange(2, 1, 10, 6).getValues();
  const existingMap = {};
  existingData.forEach(row => { if (row[0]) existingMap[String(row[0]).trim().toUpperCase()] = row; });
  Object.keys(results).forEach(tk => {
    existingMap[tk] = [tk, results[tk].eps, results[tk].bvps, results[tk].roe, existingMap[tk]?.[4] || '', existingMap[tk]?.[5] || ''];
  });
  const sorted = Object.values(existingMap).sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  sheet.getRange(2, 1, Math.max(1, sorted.length), 6).clearContent();
  sheet.getRange(2, 1, sorted.length, 6).setValues(sorted);
  console.log(`Fundamentals updated: ${count} reports processed`);
  return { ok: true, message: `Processed ${count} report(s)` };
}

// ── UPDATE 52-WEEK HIGH/LOW FROM HISTORICAL DATA ────────────────────────────
// Runs daily with dailyDSEUpdate to keep 52W High/Low current
function update52WHighLow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const histSheet = ss.getSheetByName("Historical_Log");
  if (!histSheet) { console.log("Historical_Log not found"); return; }

  // Auto-create Fundamentals sheet if missing
  let fundSheet = ss.getSheetByName("Fundamentals");
  if (!fundSheet) {
    fundSheet = ss.insertSheet("Fundamentals");
    fundSheet.getRange("A1:F1").setValues([["Ticker", "EPS (TZS)", "BVPS (TZS)", "ROE (%)", "52W High", "52W Low"]]);
    fundSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#1A5F5F").setFontColor("#ffffff");
    console.log("Created Fundamentals sheet");
  }

  const now = new Date();
  const oneYearAgo = new Date(now);
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

  const allData = histSheet.getRange(2, 1, Math.max(1, histSheet.getLastRow() - 1), 4).getValues();
  const priceMap = {};
  allData.forEach(row => {
    const date = row[0];
    const ticker = String(row[1]).trim().toUpperCase();
    const close = row[3];
    if (!date || !ticker || typeof close !== 'number' || close <= 0) return;
    if (date instanceof Date && date < oneYearAgo) return;
    if (!priceMap[ticker]) priceMap[ticker] = [];
    priceMap[ticker].push(close);
  });

  const fundData = fundSheet.getRange(2, 1, 10, 6).getValues();
  let changed = false;
  fundData.forEach(row => {
    if (!row[0]) return;
    const ticker = String(row[0]).trim().toUpperCase();
    const prices = priceMap[ticker];
    if (prices && prices.length > 0) {
      const hi = Math.round(Math.max(...prices));
      const lo = Math.round(Math.min(...prices));
      if (row[4] !== hi || row[5] !== lo) { row[4] = hi; row[5] = lo; changed = true; }
    }
  });
  if (changed) fundSheet.getRange(2, 1, fundData.length, 6).setValues(fundData);
  console.log("52W High/Low updated");
}

// ── REFRESH ALL FUNDAMENTALS (PDF extraction + 52W calculation) ──────────────
function refreshFundamentals() {
  extractFundamentalsFromPDFs();
  update52WHighLow();
  console.log("Fundamentals refresh complete");
}