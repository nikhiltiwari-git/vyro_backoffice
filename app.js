/* ================================================
   Vyro Backoffice — app.js
   Pure Vanilla JS — No frameworks
   
   Dependencies (CDN): SheetJS (xlsx), ApexCharts
   ================================================ */

(function () {
  'use strict';

  // ──────────────────────────────────────────────
  // GLOBALS
  // ──────────────────────────────────────────────
  let uploadedFiles = { trade: null, pl: null, fifo: null };
  let rawTrades = [];         // Store parsed trades for fast re-filtering
  let analysisResult = null;
  let charts = {};
  let currentEquityView = 'equity'; // 'equity' | 'drawdown'
  let calCurrentMonth = null; // {year, month} for calendar navigation
  let calPnlMap = {};         // dateKey -> dailyPnl entry for quick lookup

  // ──────────────────────────────────────────────
  // UTILITY FUNCTIONS
  // ──────────────────────────────────────────────
  function formatINR(num) {
    if (num == null || isNaN(num)) return '—';
    const abs = Math.abs(num);
    const sign = num < 0 ? '-' : '';
    // Indian numbering system: last 3 digits, then pairs of 2
    let s = abs.toFixed(2);
    let [intPart, dec] = s.split('.');
    let lastThree = intPart.slice(-3);
    let rest = intPart.slice(0, -3);
    if (rest.length > 0) {
      lastThree = ',' + lastThree;
      rest = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',');
    }
    return sign + '₹' + rest + lastThree + '.' + dec;
  }

  function formatINRShort(num) {
    if (num == null || isNaN(num)) return '—';
    const abs = Math.abs(num);
    const sign = num < 0 ? '-' : '';
    if (abs >= 1e7) return sign + '₹' + (num / 1e7).toFixed(2) + ' Cr';
    if (abs >= 1e5) return sign + '₹' + (num / 1e5).toFixed(2) + ' L';
    if (abs >= 1e3) return sign + '₹' + (num / 1e3).toFixed(1) + ' K';
    return sign + '₹' + abs.toFixed(0);
  }

  function formatPct(num) {
    if (num == null || isNaN(num)) return '—';
    return (num >= 0 ? '+' : '') + num.toFixed(2) + '%';
  }

  function formatDate(d) {
    if (!d) return '—';
    const dt = new Date(d);
    return dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  function parseExcelDate(val) {
    if (val == null) return null;
    // If it's already a Date
    if (val instanceof Date) return val;
    // If it's a number (Excel serial date)
    if (typeof val === 'number') {
      // Excel serial date to JS date
      const epoch = new Date(1899, 11, 30);
      return new Date(epoch.getTime() + val * 86400000);
    }
    // String date — try dd/mm/yyyy or dd-mm-yyyy
    if (typeof val === 'string') {
      val = val.trim();
      // dd/mm/yyyy or dd-mm-yyyy
      let m = val.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (m) {
        return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
      }
      // yyyy-mm-dd
      m = val.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (m) {
        return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
      }
      // Try native parse as last resort
      const d = new Date(val);
      return isNaN(d.getTime()) ? null : d;
    }
    return null;
  }

  function toNum(val) {
    if (val == null) return 0;
    const n = parseFloat(val);
    return isNaN(n) ? 0 : n;
  }

  function dayOfWeekName(d) {
    return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][d.getDay()];
  }

  function dateKey(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function weekKey(d) {
    // ISO week number
    const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    dt.setDate(dt.getDate() + 3 - ((dt.getDay() + 6) % 7));
    const week1 = new Date(dt.getFullYear(), 0, 4);
    const wn = 1 + Math.round(((dt.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
    return dt.getFullYear() + '-W' + String(wn).padStart(2, '0');
  }

  function monthKey(d) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[d.getMonth()] + ' ' + d.getFullYear();
  }

  // ──────────────────────────────────────────────
  // 1. PARSE TRADE REPORT
  // ──────────────────────────────────────────────
  function parseTradeReport(workbook) {
    // Sheet name may have trailing space
    let sheetName = workbook.SheetNames.find(s => s.trim().toLowerCase().includes('trade'));
    if (!sheetName) sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert to JSON, header at row 3 (0-indexed row 2)
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    // Find header row — look for 'Trade Date' in first column
    let headerIdx = -1;
    for (let i = 0; i < Math.min(10, rawData.length); i++) {
      if (rawData[i] && rawData[i][0] && String(rawData[i][0]).trim() === 'Trade Date') {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx === -1) throw new Error('Could not find header row in Trade Report');

    const headers = rawData[headerIdx].map(h => h ? String(h).trim() : '');
    const trades = [];

    for (let i = headerIdx + 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (!row || !row[0]) continue;

      const tradeDate = parseExcelDate(row[headers.indexOf('Trade Date')]);
      if (!tradeDate) continue; // Skip non-date rows (client name, totals)

      const quantity = toNum(row[headers.indexOf('Quantity')]);
      if (quantity === 0) continue;

      const instrument = String(row[headers.indexOf('Instrument')] || '').trim().toUpperCase();
      const symbol = String(row[headers.indexOf('Symbol')] || '').trim().toUpperCase();
      const expiryRaw = row[headers.indexOf('Expiry Date')];
      const expiryDate = parseExcelDate(expiryRaw);
      const strikePrice = toNum(row[headers.indexOf('Strike Price')]);
      const optionType = String(row[headers.indexOf('Option Type')] || '').trim().toUpperCase();
      const rate = toNum(row[headers.indexOf('Rate')]);
      const brokerage = toNum(row[headers.indexOf('Brokerage')]);
      const tradeTime = row[headers.indexOf('Trade Time')];

      const side = quantity > 0 ? 'Buy' : 'Sell';
      const absQty = Math.abs(quantity);
      const exchange = symbol.includes('SENSEX') ? 'BSE' : 'NSE';

      trades.push({
        tradeDate, tradeTime, instrument, symbol, expiryDate,
        strikePrice, optionType, side, absQty, rate, brokerage, exchange,
        quantity
      });
    }

    // Sort by date and time
    trades.sort((a, b) => {
      const da = a.tradeDate.getTime();
      const db = b.tradeDate.getTime();
      if (da !== db) return da - db;
      return String(a.tradeTime || '').localeCompare(String(b.tradeTime || ''));
    });

    return trades;
  }

  function applyGlobalFilterAndRender() {
    if (rawTrades.length === 0) return;
    
    // Get filter dates
    const fromVal = document.getElementById('global-date-from').value;
    const toVal = document.getElementById('global-date-to').value;
    
    let filteredTrades = rawTrades;
    if (fromVal && toVal) {
      const from = new Date(fromVal);
      const to = new Date(toVal);
      to.setHours(23, 59, 59);
      filteredTrades = rawTrades.filter(t => t.tradeDate >= from && t.tradeDate <= to);
    }
    
    if (filteredTrades.length === 0) {
      alert("No trades found in this date range.");
      return;
    }

    const initialCapital = parseFloat(document.getElementById('initial-capital').value) || 3000000;
    analysisResult = runAnalysis(filteredTrades, initialCapital);
    
    // Set default dates to filter inputs if empty
    if (!fromVal || !toVal) {
      const first = analysisResult.dailyPnl[0].date;
      const last = analysisResult.dailyPnl[analysisResult.dailyPnl.length - 1].date;
      document.getElementById('global-date-from').value = first.toISOString().split('T')[0];
      document.getElementById('global-date-to').value = last.toISOString().split('T')[0];
    }
    
    renderDashboard(analysisResult);
  }

  // ──────────────────────────────────────────────
  // 2. CHARGE CALCULATION
  // Matches the Python script's calculate_charges()
  // ──────────────────────────────────────────────
  function calculateCharges(trade) {
    const brokerage = trade.brokerage;
    const turnover = trade.absQty * trade.rate;
    const side = trade.side;
    const instrument = trade.instrument;
    const exchange = trade.exchange;

    let stt = 0, exchangeCharge = 0, gst = 0, sebi = 0, stamp = 0;

    if (instrument === 'OPTIDX') {
      // STT: 0.05% on sell side
      if (side === 'Sell') {
        stt = 0.0005 * turnover;
      }
      // Exchange transaction charges
      if (exchange === 'NSE') {
        exchangeCharge = 0.0000295 * turnover;
      } else {
        exchangeCharge = 0.0000325 * turnover;
      }
      // Stamp duty: 0.003% on buy side
      if (side === 'Buy') {
        stamp = 0.00003 * turnover;
      }
    } else if (instrument === 'FUTIDX') {
      // STT: 0.01% on sell side
      if (side === 'Sell') {
        stt = 0.0001 * turnover;
      }
      if (exchange === 'NSE') {
        exchangeCharge = 0.000019 * turnover;
      } else {
        exchangeCharge = 0.000021 * turnover;
      }
      if (side === 'Buy') {
        stamp = 0.00002 * turnover;
      }
    }

    // SEBI: ₹10 per crore = 0.0001%
    sebi = 0.000001 * turnover;

    // GST: 18% on (brokerage + exchange charge + SEBI)
    gst = 0.18 * (brokerage + exchangeCharge + sebi);

    const totalCharges = brokerage + stt + exchangeCharge + gst + sebi + stamp;

    return {
      brokerage, stt, exchangeCharge, gst, sebi, stamp, totalCharges
    };
  }

  // ──────────────────────────────────────────────
  // 3. FIFO PnL MATCHING
  // Direct translation of Python compute_realized_pnl()
  // ──────────────────────────────────────────────
  function computeRealizedPnl(trades) {
    // Using Maps with string keys
    const longLots = {};  // key -> array of {qty, price, date}
    const shortLots = {};
    const realizedEvents = [];

    for (const trade of trades) {
      const key = [trade.symbol, trade.expiryDate ? trade.expiryDate.toISOString() : '',
        trade.strikePrice, trade.optionType, trade.instrument].join('|');
      const side = trade.side;
      let qty = trade.absQty;
      const price = trade.rate;
      const tradeDate = trade.tradeDate;

      if (!longLots[key]) longLots[key] = [];
      if (!shortLots[key]) shortLots[key] = [];

      if (side === 'Buy') {
        // Close short positions first
        while (qty > 0 && shortLots[key].length > 0) {
          const lot = shortLots[key][0];
          const matchQty = Math.min(qty, lot.qty);
          const pnl = (lot.price - price) * matchQty; // short cover
          realizedEvents.push({
            date: tradeDate,
            tradeTime: trade.tradeTime,
            contract: key,
            symbol: trade.symbol,
            optionType: trade.optionType,
            strikePrice: trade.strikePrice,
            pnl: pnl,
            qty: matchQty,
            entryPrice: lot.price,
            exitPrice: price
          });
          lot.qty -= matchQty;
          qty -= matchQty;
          if (lot.qty <= 0) shortLots[key].shift();
        }
        if (qty > 0) {
          longLots[key].push({ qty, price, date: tradeDate });
        }
      } else {
        // Close long positions first
        while (qty > 0 && longLots[key].length > 0) {
          const lot = longLots[key][0];
          const matchQty = Math.min(qty, lot.qty);
          const pnl = (price - lot.price) * matchQty; // sell
          realizedEvents.push({
            date: tradeDate,
            tradeTime: trade.tradeTime,
            contract: key,
            symbol: trade.symbol,
            optionType: trade.optionType,
            strikePrice: trade.strikePrice,
            pnl: pnl,
            qty: matchQty,
            entryPrice: lot.price,
            exitPrice: price
          });
          lot.qty -= matchQty;
          qty -= matchQty;
          if (lot.qty <= 0) longLots[key].shift();
        }
        if (qty > 0) {
          shortLots[key].push({ qty, price, date: tradeDate });
        }
      }
    }

    // Open positions
    const openPositions = [];
    for (const key in longLots) {
      for (const lot of longLots[key]) {
        if (lot.qty > 0) {
          openPositions.push({ contract: key, side: 'Long', qty: lot.qty, avgPrice: lot.price, date: lot.date });
        }
      }
    }
    for (const key in shortLots) {
      for (const lot of shortLots[key]) {
        if (lot.qty > 0) {
          openPositions.push({ contract: key, side: 'Short', qty: lot.qty, avgPrice: lot.price, date: lot.date });
        }
      }
    }

    return { realizedEvents, openPositions };
  }

  // ──────────────────────────────────────────────
  // 4. MAIN ANALYSIS ENGINE
  // ──────────────────────────────────────────────
  function runAnalysis(trades, initialCapital) {
    // --- Calculate charges for each trade ---
    const tradesWithCharges = trades.map(t => {
      const charges = calculateCharges(t);
      return { ...t, ...charges };
    });

    // --- FIFO PnL ---
    const { realizedEvents, openPositions } = computeRealizedPnl(trades);

    // --- Daily Gross PnL from realized events ---
    const dailyGrossPnl = {};
    for (const ev of realizedEvents) {
      const dk = dateKey(ev.date);
      dailyGrossPnl[dk] = (dailyGrossPnl[dk] || 0) + ev.pnl;
    }

    // --- Daily Total Charges ---
    const dailyCharges = {};
    const totalChargesBreakdown = { brokerage: 0, stt: 0, exchangeCharge: 0, gst: 0, sebi: 0, stamp: 0, totalCharges: 0 };
    for (const t of tradesWithCharges) {
      const dk = dateKey(t.tradeDate);
      dailyCharges[dk] = (dailyCharges[dk] || 0) + t.totalCharges;

      totalChargesBreakdown.brokerage += t.brokerage;
      totalChargesBreakdown.stt += t.stt;
      totalChargesBreakdown.exchangeCharge += t.exchangeCharge;
      totalChargesBreakdown.gst += t.gst;
      totalChargesBreakdown.sebi += t.sebi;
      totalChargesBreakdown.stamp += t.stamp;
      totalChargesBreakdown.totalCharges += t.totalCharges;
    }

    // --- Merge into daily PnL ---
    const allDates = new Set([...Object.keys(dailyGrossPnl), ...Object.keys(dailyCharges)]);
    const sortedDates = Array.from(allDates).sort();

    let cumulativeNetPnl = 0;
    const dailyPnl = [];

    for (const dk of sortedDates) {
      const gross = dailyGrossPnl[dk] || 0;
      const charges = dailyCharges[dk] || 0;
      const net = gross - charges;
      cumulativeNetPnl += net;
      const equity = initialCapital + cumulativeNetPnl;

      dailyPnl.push({
        date: new Date(dk),
        dateStr: dk,
        grossPnl: gross,
        totalCharges: charges,
        netPnl: net,
        cumulativeNetPnl,
        equity
      });
    }

    // --- Drawdown ---
    let peak = initialCapital;
    for (const day of dailyPnl) {
      if (day.equity > peak) peak = day.equity;
      day.peak = peak;
      day.drawdown = day.equity - peak;
      day.drawdownPct = peak > 0 ? (day.drawdown / peak) * 100 : 0;
    }

    const maxDrawdown = dailyPnl.length > 0 ? Math.min(...dailyPnl.map(d => d.drawdown)) : 0;
    const maxDrawdownPct = dailyPnl.length > 0 ? Math.min(...dailyPnl.map(d => d.drawdownPct)) : 0;

    // --- Statistics ---
    const totalNetPnl = dailyPnl.reduce((s, d) => s + d.netPnl, 0);
    const totalGrossPnl = dailyPnl.reduce((s, d) => s + d.grossPnl, 0);
    const totalChargesSum = dailyPnl.reduce((s, d) => s + d.totalCharges, 0);

    const winDays = dailyPnl.filter(d => d.netPnl > 0);
    const lossDays = dailyPnl.filter(d => d.netPnl < 0);
    const flatDays = dailyPnl.filter(d => d.netPnl === 0);

    const winRate = dailyPnl.length > 0 ? (winDays.length / dailyPnl.length) * 100 : 0;
    const avgWin = winDays.length > 0 ? winDays.reduce((s, d) => s + d.netPnl, 0) / winDays.length : 0;
    const avgLoss = lossDays.length > 0 ? lossDays.reduce((s, d) => s + d.netPnl, 0) / lossDays.length : 0;

    const totalWinAmount = winDays.reduce((s, d) => s + d.netPnl, 0);
    const totalLossAmount = Math.abs(lossDays.reduce((s, d) => s + d.netPnl, 0));
    const profitFactor = totalLossAmount > 0 ? totalWinAmount / totalLossAmount : Infinity;

    // Expectancy = (WinRate * AvgWin) + ((1-WinRate) * AvgLoss)
    const expectancy = (winRate / 100 * avgWin) + ((1 - winRate / 100) * avgLoss);

    // Daily returns based on initial capital for volatility calculation
    const dailyReturns = dailyPnl.map(d => d.netPnl / initialCapital);

    // Daily Volatility (standard deviation of daily returns)
    const meanReturn = dailyReturns.length > 0 ? dailyReturns.reduce((s, r) => s + r, 0) / dailyReturns.length : 0;
    const dailyVariance = dailyReturns.length > 1
      ? dailyReturns.reduce((s, r) => s + Math.pow(r - meanReturn, 2), 0) / (dailyReturns.length - 1)
      : 0;
    const dailyVolatility = Math.sqrt(dailyVariance);

    // Annualized volatility: daily vol * sqrt(252)
    const annualizedVolatility = dailyVolatility * Math.sqrt(252);

    // Sharpe Ratio: (mean daily return / daily std dev) * sqrt(252)
    // Using risk-free rate = 0 for simplicity
    const sharpe = dailyVolatility > 0 ? (meanReturn / dailyVolatility) * Math.sqrt(252) : 0;

    // Sortino Ratio: uses downside deviation only
    const downsideReturns = dailyReturns.filter(r => r < 0);
    const downsideVariance = downsideReturns.length > 1
      ? downsideReturns.reduce((s, r) => s + Math.pow(r, 2), 0) / (downsideReturns.length - 1)
      : 0;
    const downsideDev = Math.sqrt(downsideVariance);
    const sortino = downsideDev > 0 ? (meanReturn * Math.sqrt(252)) / (downsideDev * Math.sqrt(252)) * Math.sqrt(252) : 0;
    // Simplified: sortino = (meanReturn / downsideDev) * sqrt(252)
    const sortinoFixed = downsideDev > 0 ? (meanReturn / downsideDev) * Math.sqrt(252) : 0;

    // Best / worst day
    const bestDay = dailyPnl.length > 0 ? dailyPnl.reduce((best, d) => d.netPnl > best.netPnl ? d : best, dailyPnl[0]) : null;
    const worstDay = dailyPnl.length > 0 ? dailyPnl.reduce((worst, d) => d.netPnl < worst.netPnl ? d : worst, dailyPnl[0]) : null;

    // --- Streaks ---
    let maxWinStreak = 0, currentWinStreak = 0;
    let maxLossStreak = 0, currentLossStreak = 0;
    for (const d of dailyPnl) {
      if (d.netPnl > 0) {
        currentWinStreak++;
        maxWinStreak = Math.max(maxWinStreak, currentWinStreak);
        currentLossStreak = 0;
      } else if (d.netPnl < 0) {
        currentLossStreak++;
        maxLossStreak = Math.max(maxLossStreak, currentLossStreak);
        currentWinStreak = 0;
      } else {
        // Flat day resets streaks
        currentWinStreak = 0;
        currentLossStreak = 0;
      }
    }

    // --- Day of week average PnL ---
    const dayOfWeekPnl = { Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [] };
    for (const d of dailyPnl) {
      const dow = dayOfWeekName(d.date);
      if (dayOfWeekPnl[dow] !== undefined) {
        dayOfWeekPnl[dow].push(d.netPnl);
      }
    }
    const dayOfWeekAvg = {};
    for (const dow in dayOfWeekPnl) {
      const arr = dayOfWeekPnl[dow];
      dayOfWeekAvg[dow] = arr.length > 0 ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;
    }

    // --- Weekly PnL ---
    const weeklyMap = {};
    for (const d of dailyPnl) {
      const wk = weekKey(d.date);
      if (!weeklyMap[wk]) weeklyMap[wk] = { grossPnl: 0, totalCharges: 0, netPnl: 0 };
      weeklyMap[wk].grossPnl += d.grossPnl;
      weeklyMap[wk].totalCharges += d.totalCharges;
      weeklyMap[wk].netPnl += d.netPnl;
    }
    const weeklyPnl = Object.entries(weeklyMap).sort((a, b) => a[0].localeCompare(b[0])).map(([k, v]) => ({ week: k, ...v }));

    // --- Monthly PnL ---
    const monthlyMap = {};
    for (const d of dailyPnl) {
      const mk = monthKey(d.date);
      if (!monthlyMap[mk]) monthlyMap[mk] = { grossPnl: 0, totalCharges: 0, netPnl: 0 };
      monthlyMap[mk].grossPnl += d.grossPnl;
      monthlyMap[mk].totalCharges += d.totalCharges;
      monthlyMap[mk].netPnl += d.netPnl;
    }
    // Sort by actual date order
    const monthOrder = {};
    for (const d of dailyPnl) {
      const mk = monthKey(d.date);
      if (!monthOrder[mk]) monthOrder[mk] = d.date.getTime();
    }
    const monthlyPnl = Object.entries(monthlyMap)
      .sort((a, b) => (monthOrder[a[0]] || 0) - (monthOrder[b[0]] || 0))
      .map(([k, v]) => ({ month: k, ...v }));

    // --- Segment PnL (Instrument x Symbol) ---
    const segmentMap = {};
    for (const ev of realizedEvents) {
      const parts = ev.contract.split('|');
      const segKey = parts[4] + ' — ' + parts[0]; // Instrument — Symbol
      segmentMap[segKey] = (segmentMap[segKey] || 0) + ev.pnl;
    }
    const segmentPnl = Object.entries(segmentMap).map(([k, v]) => ({ segment: k, pnl: v }));

    // --- Time of Day Analysis ---
    const timeOfDayMap = {};
    for (const ev of realizedEvents) {
      if (!ev.tradeTime) continue;
      const parts = ev.tradeTime.split(':');
      if (parts.length >= 1) {
        let hr = parseInt(parts[0], 10);
        if (hr >= 9 && hr <= 15) {
          const slot = String(hr).padStart(2, '0') + ':00';
          timeOfDayMap[slot] = (timeOfDayMap[slot] || 0) + ev.pnl;
        }
      }
    }
    const timeOfDayPnl = Object.entries(timeOfDayMap).map(([k, v]) => ({
      time: k, pnl: v
    })).sort((a, b) => a.time.localeCompare(b.time));

    return {
      trades: tradesWithCharges,
      realizedEvents,
      openPositions,
      dailyPnl,
      weeklyPnl,
      monthlyPnl,
      segmentPnl,
      dayOfWeekAvg,
      totalChargesBreakdown,
      timeOfDayPnl,
      stats: {
        initialCapital,
        totalNetPnl,
        totalGrossPnl,
        totalCharges: totalChargesSum,
        winRate,
        winDays: winDays.length,
        lossDays: lossDays.length,
        flatDays: flatDays.length,
        avgWin,
        avgLoss,
        profitFactor,
        expectancy,
        maxDrawdown,
        maxDrawdownPct,
        sharpe,
        sortino: sortinoFixed,
        dailyVolatility,
        annualizedVolatility,
        bestDay,
        worstDay,
        totalTradingDays: dailyPnl.length,
        totalTrades: trades.length,
        maxWinStreak,
        maxLossStreak
      }
    };
  }

  // ──────────────────────────────────────────────
  // 5. UI RENDERING
  // ──────────────────────────────────────────────
  function renderDashboard(result) {
    analysisResult = result;
    const s = result.stats;

    // --- Switch screens ---
    document.getElementById('upload-screen').classList.add('hidden');
    document.getElementById('dashboard-screen').classList.remove('hidden');

    // --- Nav capital ---
    document.getElementById('nav-capital').textContent = formatINRShort(s.initialCapital);

    // --- Stat cards ---
    const setVal = (id, val, cls) => {
      const el = document.getElementById(id);
      if (el) { el.textContent = val; el.className = 'stat-value' + (cls ? ' ' + cls : ''); }
    };
    const setSub = (id, val, cls) => {
      const el = document.getElementById(id);
      if (el) { el.textContent = val; el.className = 'stat-sub' + (cls ? ' ' + cls : ''); }
    };

    setVal('stat-net-pnl', formatINR(s.totalNetPnl), s.totalNetPnl >= 0 ? 'profit' : 'loss');
    setSub('stat-net-pnl-pct', formatPct(s.totalNetPnl / s.initialCapital * 100), s.totalNetPnl >= 0 ? 'profit' : 'loss');

    setVal('stat-gross-pnl', formatINR(s.totalGrossPnl), s.totalGrossPnl >= 0 ? 'profit' : 'loss');
    setSub('stat-total-charges-inline', 'Charges: ' + formatINR(s.totalCharges));

    setVal('stat-win-rate', s.winRate.toFixed(1) + '%');
    setSub('stat-win-loss-count', s.winDays + ' W / ' + s.lossDays + ' L');

    setVal('stat-profit-factor', s.profitFactor === Infinity ? '∞' : s.profitFactor.toFixed(2));

    setVal('stat-total-trades', s.totalTrades);

    setVal('stat-expectancy-main', formatINR(s.expectancy), s.expectancy >= 0 ? 'profit' : 'loss');

    setVal('stat-win-streak', s.maxWinStreak + ' Days', 'profit');
    setVal('stat-loss-streak', s.maxLossStreak + ' Days', 'loss');

    const avgTradesPerDay = s.totalTradingDays > 0 ? (s.totalTrades / s.totalTradingDays) : 0;
    setVal('stat-avg-trades', avgTradesPerDay.toFixed(1));

    setVal('stat-max-dd', formatINR(s.maxDrawdown), 'loss');
    setSub('stat-max-dd-pct', formatPct(s.maxDrawdownPct), 'loss');

    setVal('stat-sharpe', s.sharpe.toFixed(2));
    setSub('stat-sortino', 'Sortino: ' + s.sortino.toFixed(2));

    setVal('stat-avg-win', formatINR(s.avgWin), 'profit');
    setSub('stat-total-win-days', 'Win days: ' + s.winDays);

    setVal('stat-avg-loss', formatINR(s.avgLoss), 'loss');
    setSub('stat-total-loss-days', 'Loss days: ' + s.lossDays);

    setVal('stat-volatility', (s.dailyVolatility * 100).toFixed(3) + '%');
    setSub('stat-annual-volatility', 'Annualized: ' + (s.annualizedVolatility * 100).toFixed(2) + '%');

    if (s.bestDay) {
      setVal('stat-best-day', formatINR(s.bestDay.netPnl), 'profit');
      setSub('stat-best-day-date', formatDate(s.bestDay.date));
    }
    if (s.worstDay) {
      setVal('stat-worst-day', formatINR(s.worstDay.netPnl), 'loss');
      setSub('stat-worst-day-date', formatDate(s.worstDay.date));
    }

    // --- Charges tab ---
    document.getElementById('stat-brokerage').textContent = formatINR(result.totalChargesBreakdown.brokerage);
    document.getElementById('stat-stt').textContent = formatINR(result.totalChargesBreakdown.stt);
    document.getElementById('stat-exchange').textContent = formatINR(result.totalChargesBreakdown.exchangeCharge);
    document.getElementById('stat-gst').textContent = formatINR(result.totalChargesBreakdown.gst);
    document.getElementById('stat-sebi').textContent = formatINR(result.totalChargesBreakdown.sebi);
    document.getElementById('stat-stamp').textContent = formatINR(result.totalChargesBreakdown.stamp);

    // --- Render charts ---
    renderEquityChart(result);
    renderWeekdayChart(result);
    renderMonthlyChart(result);
    renderTimeOfDayChart(result);
    renderCalendar(result);
    renderWeeklyChart(result);
    renderChargesDonut(result);
    renderChargesVsPnl(result);

    // --- Render tables ---
    renderTradeTable(result);
    renderDailyPnlTable(result);
  }

  // ──────────────────────────────────────────────
  // 6. CHARTS — ApexCharts
  // ──────────────────────────────────────────────
  const chartDefaults = {
    chart: {
      background: 'transparent',
      foreColor: '#8b8ba3',
      fontFamily: "'Inter', sans-serif",
      toolbar: { show: true, tools: { download: true, selection: false, zoom: true, zoomin: true, zoomout: true, pan: true, reset: true } },
      animations: { enabled: true, easing: 'easeinout', speed: 600 }
    },
    grid: { borderColor: 'rgba(255,255,255,0.04)', strokeDashArray: 3 },
    tooltip: { theme: 'dark' },
    colors: ['#6366f1', '#06b6d4', '#22c55e', '#f59e0b', '#ef4444', '#8b5cf6']
  };

  function destroyChart(name) {
    if (charts[name]) {
      charts[name].destroy();
      charts[name] = null;
    }
  }

  function renderEquityChart(result) {
    destroyChart('equity');
    const daily = result.dailyPnl;
    if (daily.length === 0) return;

    const equityData = daily.map(d => ({ x: d.date.getTime(), y: d.equity }));
    const drawdownData = daily.map(d => ({ x: d.date.getTime(), y: d.drawdown }));

    const showEquity = currentEquityView === 'equity';

    const opts = {
      ...chartDefaults,
      chart: {
        ...chartDefaults.chart,
        type: 'area',
        height: 340,
      },
      series: [{
        name: showEquity ? 'Equity' : 'Drawdown',
        data: showEquity ? equityData : drawdownData
      }],
      colors: showEquity ? ['#6366f1'] : ['#ef4444'],
      fill: {
        type: 'gradient',
        gradient: {
          shadeIntensity: 1,
          opacityFrom: 0.4,
          opacityTo: 0.05,
          stops: [0, 100]
        }
      },
      stroke: { curve: 'smooth', width: 2.5 },
      xaxis: {
        type: 'datetime',
        labels: { style: { colors: '#5a5a72', fontSize: '11px' } },
        axisBorder: { show: false },
        axisTicks: { show: false }
      },
      yaxis: {
        labels: {
          style: { colors: '#5a5a72', fontSize: '11px' },
          formatter: v => formatINRShort(v)
        }
      },
      dataLabels: { enabled: false },
      tooltip: {
        theme: 'dark',
        x: { format: 'dd MMM yyyy' },
        y: { formatter: v => formatINR(v) }
      }
    };

    charts.equity = new ApexCharts(document.getElementById('chart-equity'), opts);
    charts.equity.render();
  }

  function renderWeekdayChart(result) {
    destroyChart('weekday');
    const dow = result.dayOfWeekAvg;
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
    const values = days.map(d => dow[d] || 0);
    const colors = values.map(v => v >= 0 ? '#22c55e' : '#ef4444');

    const opts = {
      ...chartDefaults,
      chart: { ...chartDefaults.chart, type: 'bar', height: 300 },
      series: [{ name: 'Avg Net P&L', data: values }],
      plotOptions: {
        bar: {
          distributed: true,
          borderRadius: 6,
          columnWidth: '55%',
          colors: { ranges: values.map((v, i) => ({ from: v >= 0 ? 0 : v, to: v >= 0 ? v : 0, color: colors[i] })) }
        }
      },
      colors: colors,
      xaxis: {
        categories: days.map(d => d.slice(0, 3)),
        labels: { style: { colors: '#8b8ba3', fontSize: '12px', fontWeight: 600 } },
        axisBorder: { show: false },
        axisTicks: { show: false }
      },
      yaxis: {
        labels: {
          style: { colors: '#5a5a72', fontSize: '11px' },
          formatter: v => formatINRShort(v)
        }
      },
      dataLabels: { enabled: false },
      legend: { show: false },
      tooltip: { theme: 'dark', y: { formatter: v => formatINR(v) } }
    };

    charts.weekday = new ApexCharts(document.getElementById('chart-weekday'), opts);
    charts.weekday.render();
  }

  function renderMonthlyChart(result) {
    destroyChart('monthly');
    const monthly = result.monthlyPnl;
    if (monthly.length === 0) return;

    const categories = monthly.map(m => m.month);
    const netValues = monthly.map(m => m.netPnl);
    const colors = netValues.map(v => v >= 0 ? '#22c55e' : '#ef4444');

    const opts = {
      ...chartDefaults,
      chart: { ...chartDefaults.chart, type: 'bar', height: 300 },
      series: [{ name: 'Net P&L', data: netValues }],
      plotOptions: {
        bar: {
          distributed: true,
          borderRadius: 6,
          columnWidth: '60%'
        }
      },
      colors: colors,
      xaxis: {
        categories,
        labels: { style: { colors: '#8b8ba3', fontSize: '11px', fontWeight: 500 }, rotate: -45 },
        axisBorder: { show: false },
        axisTicks: { show: false }
      },
      yaxis: {
        labels: {
          style: { colors: '#5a5a72', fontSize: '11px' },
          formatter: v => formatINRShort(v)
        }
      },
      dataLabels: { enabled: false },
      legend: { show: false },
      tooltip: { theme: 'dark', y: { formatter: v => formatINR(v) } }
    };

    charts.monthly = new ApexCharts(document.getElementById('chart-monthly'), opts);
    charts.monthly.render();
  }

  function renderTimeOfDayChart(result) {
    destroyChart('timeofday');
    const tod = result.timeOfDayPnl;
    if (!tod || tod.length === 0) return;

    const categories = tod.map(t => t.time);
    const values = tod.map(t => t.pnl);
    const colors = values.map(v => v >= 0 ? '#22c55e' : '#ef4444');

    const opts = {
      ...chartDefaults,
      chart: { ...chartDefaults.chart, type: 'bar', height: 260 },
      series: [{ name: 'Net P&L', data: values }],
      plotOptions: {
        bar: {
          borderRadius: 4,
          columnWidth: '50%',
          distributed: true
        }
      },
      colors: colors,
      xaxis: {
        categories: categories,
        labels: { style: { colors: '#8b8ba3', fontSize: '11px', fontWeight: 500 } },
        axisBorder: { show: false }
      },
      yaxis: {
        labels: {
          style: { colors: '#5a5a72', fontSize: '11px' },
          formatter: v => formatINRShort(v)
        }
      },
      dataLabels: { enabled: false },
      legend: { show: false },
      tooltip: { theme: 'dark', y: { formatter: v => formatINR(v) } }
    };

    charts.timeofday = new ApexCharts(document.getElementById('chart-timeofday'), opts);
    charts.timeofday.render();
  }


  // ──────────────────────────────────────────────
  // CALENDAR — Real monthly calendar with navigation
  // ──────────────────────────────────────────────
  function renderCalendar(result) {
    const daily = result.dailyPnl;
    if (daily.length === 0) return;

    // Build lookup map: dateKey -> daily entry
    calPnlMap = {};
    for (const d of daily) {
      calPnlMap[d.dateStr] = d;
    }

    // Default to the last month with data
    const lastDay = daily[daily.length - 1].date;
    calCurrentMonth = { year: lastDay.getFullYear(), month: lastDay.getMonth() };

    // Set date filter defaults
    const firstDay = daily[0].date;
    document.getElementById('cal-date-from').value = firstDay.toISOString().split('T')[0];
    document.getElementById('cal-date-to').value = lastDay.toISOString().split('T')[0];

    renderCalendarMonth();
  }

  function renderCalendarMonth() {
    const { year, month } = calCurrentMonth;
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'];

    // Update title
    document.getElementById('cal-month-title').textContent = monthNames[month] + ' ' + year;

    // First day of month (0=Sun, convert so Mon=0)
    const firstDayOfMonth = new Date(year, month, 1);
    let startDow = firstDayOfMonth.getDay(); // 0=Sun
    startDow = (startDow + 6) % 7; // Convert to Mon=0

    const daysInMonth = new Date(year, month + 1, 0).getDate();

    const grid = document.getElementById('cal-grid');
    grid.innerHTML = '';

    // Compute max absolute PnL this month for intensity scaling
    let maxAbs = 0;
    for (let d = 1; d <= daysInMonth; d++) {
      const dk = year + '-' + String(month + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0');
      if (calPnlMap[dk]) {
        maxAbs = Math.max(maxAbs, Math.abs(calPnlMap[dk].netPnl));
      }
    }

    // Empty cells before the 1st
    for (let i = 0; i < startDow; i++) {
      const empty = document.createElement('div');
      empty.className = 'cal-day empty';
      grid.appendChild(empty);
    }

    // Day cells
    for (let d = 1; d <= daysInMonth; d++) {
      const dk = year + '-' + String(month + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0');
      const dayData = calPnlMap[dk];
      const cell = document.createElement('div');
      cell.className = 'cal-day';

      const numSpan = document.createElement('span');
      numSpan.className = 'cal-day-num';
      numSpan.textContent = d;
      cell.appendChild(numSpan);

      if (dayData) {
        const pnl = dayData.netPnl;
        const pnlSpan = document.createElement('span');
        pnlSpan.className = 'cal-day-pnl';

        // Short format for the cell
        const abs = Math.abs(pnl);
        if (abs >= 1e5) pnlSpan.textContent = (pnl >= 0 ? '+' : '-') + (abs / 1e3).toFixed(0) + 'K';
        else if (abs >= 1e3) pnlSpan.textContent = (pnl >= 0 ? '+' : '-') + (abs / 1e3).toFixed(1) + 'K';
        else pnlSpan.textContent = (pnl >= 0 ? '+' : '') + pnl.toFixed(0);
        cell.appendChild(pnlSpan);

        // Intensity level (1-3 based on magnitude relative to max)
        const intensity = maxAbs > 0 ? Math.min(3, Math.ceil((Math.abs(pnl) / maxAbs) * 3)) : 1;

        if (pnl > 0) {
          cell.classList.add('profit', 'profit-' + intensity);
        } else if (pnl < 0) {
          cell.classList.add('loss', 'loss-' + intensity);
        } else {
          cell.classList.add('flat');
        }

        // Tooltip events
        cell.addEventListener('mouseenter', (e) => showCalTooltip(e, dayData));
        cell.addEventListener('mousemove', (e) => moveCalTooltip(e));
        cell.addEventListener('mouseleave', hideCalTooltip);
      } else {
        cell.classList.add('no-trade');
      }

      grid.appendChild(cell);
    }

    // Update monthly stats sidebar
    updateCalMonthStats(year, month);
  }

  function showCalTooltip(e, dayData) {
    const tooltip = document.getElementById('cal-tooltip');
    const dateEl = document.getElementById('cal-tooltip-date');
    const pnlEl = document.getElementById('cal-tooltip-pnl');
    const detailEl = document.getElementById('cal-tooltip-detail');

    const dayName = dayOfWeekName(dayData.date);
    dateEl.textContent = dayName + ', ' + formatDate(dayData.date);

    pnlEl.textContent = 'Net P&L: ' + formatINR(dayData.netPnl);
    pnlEl.className = 'cal-tooltip-pnl ' + (dayData.netPnl >= 0 ? 'profit' : 'loss');

    detailEl.innerHTML =
      'Gross: ' + formatINR(dayData.grossPnl) + '<br>' +
      'Charges: ' + formatINR(dayData.totalCharges) + '<br>' +
      'Equity: ' + formatINR(dayData.equity) + '<br>' +
      'Drawdown: ' + formatPct(dayData.drawdownPct);

    tooltip.classList.add('visible');
    moveCalTooltip(e);
  }

  function moveCalTooltip(e) {
    const tooltip = document.getElementById('cal-tooltip');
    const card = document.querySelector('.cal-card-main');
    const cardRect = card.getBoundingClientRect();
    let x = e.clientX - cardRect.left + 16;
    let y = e.clientY - cardRect.top + 16;

    // Keep tooltip within the card
    const tw = tooltip.offsetWidth;
    const th = tooltip.offsetHeight;
    if (x + tw > cardRect.width - 20) x = x - tw - 32;
    if (y + th > cardRect.height - 20) y = y - th - 32;

    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
  }

  function hideCalTooltip() {
    document.getElementById('cal-tooltip').classList.remove('visible');
  }

  function updateCalMonthStats(year, month) {
    // Filter daily PnL for this month
    const monthDays = (analysisResult ? analysisResult.dailyPnl : []).filter(d => {
      return d.date.getFullYear() === year && d.date.getMonth() === month;
    });

    const setStatVal = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    const setStatCls = (id, val, cls) => {
      const el = document.getElementById(id);
      if (el) { el.textContent = val; el.className = 'cal-stat-val' + (cls ? ' ' + cls : ''); }
    };

    if (monthDays.length === 0) {
      ['cal-m-pnl','cal-m-gross','cal-m-charges','cal-m-wins','cal-m-losses',
       'cal-m-winrate','cal-m-best','cal-m-worst','cal-m-avg','cal-m-days']
        .forEach(id => setStatVal(id, '—'));
      return;
    }

    const netPnl = monthDays.reduce((s, d) => s + d.netPnl, 0);
    const grossPnl = monthDays.reduce((s, d) => s + d.grossPnl, 0);
    const charges = monthDays.reduce((s, d) => s + d.totalCharges, 0);
    const wins = monthDays.filter(d => d.netPnl > 0).length;
    const losses = monthDays.filter(d => d.netPnl < 0).length;
    const winRate = monthDays.length > 0 ? (wins / monthDays.length * 100) : 0;
    const best = Math.max(...monthDays.map(d => d.netPnl));
    const worst = Math.min(...monthDays.map(d => d.netPnl));
    const avg = netPnl / monthDays.length;

    setStatCls('cal-m-pnl', formatINR(netPnl), netPnl >= 0 ? 'profit' : 'loss');
    setStatCls('cal-m-gross', formatINR(grossPnl), grossPnl >= 0 ? 'profit' : 'loss');
    setStatVal('cal-m-charges', formatINR(charges));
    setStatCls('cal-m-wins', wins, 'profit');
    setStatCls('cal-m-losses', losses, 'loss');
    setStatVal('cal-m-winrate', winRate.toFixed(1) + '%');
    setStatCls('cal-m-best', formatINR(best), 'profit');
    setStatCls('cal-m-worst', formatINR(worst), 'loss');
    setStatCls('cal-m-avg', formatINR(avg), avg >= 0 ? 'profit' : 'loss');
    setStatVal('cal-m-days', monthDays.length);
  }

  function applyCalFilter() {
    const fromVal = document.getElementById('cal-date-from').value;
    const toVal = document.getElementById('cal-date-to').value;
    if (!fromVal || !toVal || !analysisResult) return;

    const from = new Date(fromVal);
    const to = new Date(toVal);
    to.setHours(23, 59, 59); // Include the end date

    const filtered = analysisResult.dailyPnl.filter(d => d.date >= from && d.date <= to);

    const statsSection = document.getElementById('cal-range-stats');
    if (filtered.length === 0) {
      statsSection.classList.add('hidden');
      return;
    }

    statsSection.classList.remove('hidden');

    const netPnl = filtered.reduce((s, d) => s + d.netPnl, 0);
    const wins = filtered.filter(d => d.netPnl > 0).length;
    const winRate = (wins / filtered.length * 100);
    const avg = netPnl / filtered.length;

    const pnlEl = document.getElementById('cal-range-pnl');
    pnlEl.textContent = formatINR(netPnl);
    pnlEl.className = 'stat-value ' + (netPnl >= 0 ? 'profit' : 'loss');

    document.getElementById('cal-range-winrate').textContent = winRate.toFixed(1) + '%';
    document.getElementById('cal-range-days').textContent = filtered.length + ' days';

    const avgEl = document.getElementById('cal-range-avg');
    avgEl.textContent = formatINR(avg);
    avgEl.className = 'stat-value ' + (avg >= 0 ? 'profit' : 'loss');
  }

  function renderWeeklyChart(result) {
    destroyChart('weekly');
    const weekly = result.weeklyPnl;
    if (weekly.length === 0) return;

    const categories = weekly.map(w => w.week);
    const values = weekly.map(w => w.netPnl);
    const colors = values.map(v => v >= 0 ? '#22c55e' : '#ef4444');

    const opts = {
      ...chartDefaults,
      chart: { ...chartDefaults.chart, type: 'bar', height: 300 },
      series: [{ name: 'Net P&L', data: values }],
      plotOptions: {
        bar: { distributed: true, borderRadius: 4, columnWidth: '70%' }
      },
      colors: colors,
      xaxis: {
        categories,
        labels: { style: { colors: '#5a5a72', fontSize: '10px' }, rotate: -45, trim: true },
        axisBorder: { show: false },
        axisTicks: { show: false }
      },
      yaxis: {
        labels: {
          style: { colors: '#5a5a72', fontSize: '11px' },
          formatter: v => formatINRShort(v)
        }
      },
      dataLabels: { enabled: false },
      legend: { show: false },
      tooltip: { theme: 'dark', y: { formatter: v => formatINR(v) } }
    };

    charts.weekly = new ApexCharts(document.getElementById('chart-weekly'), opts);
    charts.weekly.render();
  }

  function renderChargesDonut(result) {
    destroyChart('chargesDonut');
    const cb = result.totalChargesBreakdown;
    const labels = ['Brokerage', 'STT', 'Exchange Txn', 'GST', 'SEBI Fee', 'Stamp Duty'];
    const values = [cb.brokerage, cb.stt, cb.exchangeCharge, cb.gst, cb.sebi, cb.stamp];

    const opts = {
      chart: {
        type: 'donut',
        height: 300,
        background: 'transparent',
        foreColor: '#8b8ba3',
        fontFamily: "'Inter', sans-serif"
      },
      series: values,
      labels,
      colors: ['#6366f1', '#ef4444', '#f59e0b', '#06b6d4', '#8b5cf6', '#22c55e'],
      stroke: { width: 0 },
      plotOptions: {
        pie: {
          donut: {
            size: '60%',
            labels: {
              show: true,
              name: { show: true, color: '#8b8ba3' },
              value: {
                show: true,
                color: '#f0f0f5',
                formatter: v => formatINR(parseFloat(v))
              },
              total: {
                show: true,
                label: 'Total',
                color: '#8b8ba3',
                formatter: w => formatINR(w.globals.seriesTotals.reduce((a, b) => a + b, 0))
              }
            }
          }
        }
      },
      legend: {
        position: 'bottom',
        fontSize: '12px',
        labels: { colors: '#8b8ba3' },
        markers: { size: 6, offsetX: -3 }
      },
      dataLabels: { enabled: false },
      tooltip: {
        theme: 'dark',
        y: { formatter: v => formatINR(v) }
      }
    };

    charts.chargesDonut = new ApexCharts(document.getElementById('chart-charges-donut'), opts);
    charts.chargesDonut.render();
  }

  function renderChargesVsPnl(result) {
    destroyChart('chargesVsPnl');
    const daily = result.dailyPnl;
    if (daily.length === 0) return;

    const opts = {
      ...chartDefaults,
      chart: { ...chartDefaults.chart, type: 'bar', height: 300, stacked: false },
      series: [
        { name: 'Gross P&L', data: daily.map(d => ({ x: d.date.getTime(), y: Math.round(d.grossPnl) })) },
        { name: 'Charges', data: daily.map(d => ({ x: d.date.getTime(), y: -Math.round(d.totalCharges) })) }
      ],
      colors: ['#6366f1', '#ef4444'],
      plotOptions: { bar: { borderRadius: 2, columnWidth: '60%' } },
      xaxis: {
        type: 'datetime',
        labels: { style: { colors: '#5a5a72', fontSize: '10px' } },
        axisBorder: { show: false },
        axisTicks: { show: false }
      },
      yaxis: {
        labels: {
          style: { colors: '#5a5a72', fontSize: '11px' },
          formatter: v => formatINRShort(v)
        }
      },
      dataLabels: { enabled: false },
      tooltip: {
        theme: 'dark',
        x: { format: 'dd MMM yyyy' },
        y: { formatter: v => formatINR(v) }
      }
    };

    charts.chargesVsPnl = new ApexCharts(document.getElementById('chart-charges-vs-pnl'), opts);
    charts.chargesVsPnl.render();
  }

  // ──────────────────────────────────────────────
  // 7. DATA TABLES
  // ──────────────────────────────────────────────
  let tradeTablePage = 0;
  const TRADES_PER_PAGE = 50;

  function renderTradeTable(result) {
    const thead = document.getElementById('trade-table-head');
    const cols = ['Date', 'Time', 'Symbol', 'Expiry', 'Strike', 'Type', 'Instr', 'Side', 'Qty', 'Rate', 'Brokerage', 'STT', 'Exch Chg', 'GST', 'SEBI', 'Stamp', 'Total Chg'];
    thead.innerHTML = cols.map(c => '<th>' + c + '</th>').join('');
    renderTradeTablePage(result.trades, 0);
  }

  function renderTradeTablePage(trades, page, filter) {
    let filtered = trades;
    if (filter) {
      const f = filter.toLowerCase();
      filtered = trades.filter(t =>
        t.symbol.toLowerCase().includes(f) ||
        formatDate(t.tradeDate).toLowerCase().includes(f) ||
        t.instrument.toLowerCase().includes(f)
      );
    }

    const start = page * TRADES_PER_PAGE;
    const pageData = filtered.slice(start, start + TRADES_PER_PAGE);
    const tbody = document.getElementById('trade-table-body');

    tbody.innerHTML = pageData.map(t => {
      const sideCls = t.side === 'Buy' ? 'profit' : 'loss';
      return '<tr>' +
        '<td>' + formatDate(t.tradeDate) + '</td>' +
        '<td>' + (t.tradeTime || '') + '</td>' +
        '<td>' + t.symbol + '</td>' +
        '<td>' + formatDate(t.expiryDate) + '</td>' +
        '<td>' + t.strikePrice + '</td>' +
        '<td>' + t.optionType + '</td>' +
        '<td>' + t.instrument + '</td>' +
        '<td class="' + sideCls + '">' + t.side + '</td>' +
        '<td>' + t.absQty + '</td>' +
        '<td>' + t.rate.toFixed(2) + '</td>' +
        '<td>' + t.brokerage.toFixed(2) + '</td>' +
        '<td>' + t.stt.toFixed(2) + '</td>' +
        '<td>' + t.exchangeCharge.toFixed(2) + '</td>' +
        '<td>' + t.gst.toFixed(2) + '</td>' +
        '<td>' + t.sebi.toFixed(4) + '</td>' +
        '<td>' + t.stamp.toFixed(4) + '</td>' +
        '<td>' + t.totalCharges.toFixed(2) + '</td>' +
        '</tr>';
    }).join('');

    // Pagination
    const totalPages = Math.ceil(filtered.length / TRADES_PER_PAGE);
    const pag = document.getElementById('trade-pagination');
    let pagHtml = '';
    const maxButtons = 10;
    const startPage = Math.max(0, page - Math.floor(maxButtons / 2));
    const endPage = Math.min(totalPages, startPage + maxButtons);

    if (page > 0) pagHtml += '<button data-page="' + (page - 1) + '">‹</button>';
    for (let i = startPage; i < endPage; i++) {
      pagHtml += '<button data-page="' + i + '"' + (i === page ? ' class="active"' : '') + '>' + (i + 1) + '</button>';
    }
    if (page < totalPages - 1) pagHtml += '<button data-page="' + (page + 1) + '">›</button>';
    pag.innerHTML = pagHtml;

    pag.querySelectorAll('button').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page);
        tradeTablePage = p;
        renderTradeTablePage(analysisResult.trades, p, document.getElementById('trade-search').value);
      });
    });
  }

  function renderDailyPnlTable(result) {
    const thead = document.getElementById('daily-pnl-head');
    const cols = ['Date', 'Day', 'Gross P&L', 'Charges', 'Net P&L', 'Cumulative', 'Equity', 'Drawdown', 'DD %'];
    thead.innerHTML = cols.map(c => '<th>' + c + '</th>').join('');

    const tbody = document.getElementById('daily-pnl-body');
    tbody.innerHTML = result.dailyPnl.map(d => {
      const netCls = d.netPnl >= 0 ? 'profit' : 'loss';
      return '<tr>' +
        '<td>' + formatDate(d.date) + '</td>' +
        '<td>' + dayOfWeekName(d.date).slice(0, 3) + '</td>' +
        '<td class="' + (d.grossPnl >= 0 ? 'profit' : 'loss') + '">' + formatINR(d.grossPnl) + '</td>' +
        '<td>' + formatINR(d.totalCharges) + '</td>' +
        '<td class="' + netCls + '">' + formatINR(d.netPnl) + '</td>' +
        '<td class="' + (d.cumulativeNetPnl >= 0 ? 'profit' : 'loss') + '">' + formatINR(d.cumulativeNetPnl) + '</td>' +
        '<td>' + formatINR(d.equity) + '</td>' +
        '<td class="loss">' + formatINR(d.drawdown) + '</td>' +
        '<td class="loss">' + formatPct(d.drawdownPct) + '</td>' +
        '</tr>';
    }).join('');
  }

  // ──────────────────────────────────────────────
  // 8. EVENT HANDLERS
  // ──────────────────────────────────────────────
  function initUploadHandlers() {
    const zones = [
      { zoneId: 'zone-trade', inputId: 'file-trade', statusId: 'status-trade', type: 'trade' },
      { zoneId: 'zone-pl', inputId: 'file-pl', statusId: 'status-pl', type: 'pl' },
      { zoneId: 'zone-fifo', inputId: 'file-fifo', statusId: 'status-fifo', type: 'fifo' }
    ];

    zones.forEach(z => {
      const zone = document.getElementById(z.zoneId);
      const input = document.getElementById(z.inputId);
      const status = document.getElementById(z.statusId);

      // Drag events
      zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
      zone.addEventListener('drop', e => {
        e.preventDefault();
        zone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
          handleFile(e.dataTransfer.files[0], z.type, zone, status);
        }
      });

      // Click / file select
      input.addEventListener('change', () => {
        if (input.files.length > 0) {
          handleFile(input.files[0], z.type, zone, status);
        }
      });
    });

    // Analyze button
    document.getElementById('btn-analyze').addEventListener('click', () => {
      if (!uploadedFiles.trade) return;
      const btn = document.getElementById('btn-analyze');
      btn.classList.add('loading');
      btn.disabled = true;

      // Use setTimeout to let the UI update before heavy processing
      setTimeout(() => {
        try {
          const workbook = XLSX.read(uploadedFiles.trade, { type: 'array' });
          rawTrades = parseTradeReport(workbook);
          applyGlobalFilterAndRender();
        } catch (err) {
          console.error('Analysis error:', err);
          alert('Error analyzing report: ' + err.message);
        } finally {
          btn.classList.remove('loading');
          btn.disabled = false;
        }
      }, 50);
    });
  }

  function handleFile(file, type, zoneEl, statusEl) {
    const reader = new FileReader();
    reader.onload = (e) => {
      uploadedFiles[type] = new Uint8Array(e.target.result);
      zoneEl.classList.add('uploaded');
      statusEl.textContent = '✓ ' + file.name;
      updateAnalyzeButton();
    };
    reader.readAsArrayBuffer(file);
  }

  function updateAnalyzeButton() {
    const btn = document.getElementById('btn-analyze');
    btn.disabled = !uploadedFiles.trade;
  }

  function initTabHandlers() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const tab = btn.dataset.tab;
        // Deactivate all
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        // Activate
        btn.classList.add('active');
        document.getElementById('tab-' + tab).classList.add('active');

        // Re-render charts for the newly visible tab (ApexCharts needs visible container)
        if (analysisResult) {
          setTimeout(() => {
            if (tab === 'overview') {
              if (charts.equity) charts.equity.render();
            }
            if (tab === 'calendar') {
              if (analysisResult) renderCalendarMonth();
              if (charts.weekly) charts.weekly.render();
            }
            if (tab === 'charges') {
              if (charts.chargesDonut) charts.chargesDonut.render();
              if (charts.chargesVsPnl) charts.chargesVsPnl.render();
            }
          }, 50);
        }
      });
    });
  }

  function initEquityToggle() {
    document.getElementById('toggle-equity').addEventListener('click', () => {
      currentEquityView = 'equity';
      document.getElementById('toggle-equity').classList.add('active');
      document.getElementById('toggle-drawdown').classList.remove('active');
      if (analysisResult) renderEquityChart(analysisResult);
    });
    document.getElementById('toggle-drawdown').addEventListener('click', () => {
      currentEquityView = 'drawdown';
      document.getElementById('toggle-drawdown').classList.add('active');
      document.getElementById('toggle-equity').classList.remove('active');
      if (analysisResult) renderEquityChart(analysisResult);
    });
  }

  function initResetButton() {
    document.getElementById('btn-reset').addEventListener('click', () => {
      // Destroy all charts
      Object.keys(charts).forEach(k => destroyChart(k));
      // Reset state
      uploadedFiles = { trade: null, pl: null, fifo: null };
      analysisResult = null;
      // Reset upload zones
      document.querySelectorAll('.upload-zone').forEach(z => z.classList.remove('uploaded'));
      document.getElementById('status-trade').textContent = 'Drop .xlsx or click to upload';
      document.getElementById('status-pl').textContent = 'Optional — Drop .xlsx or click';
      document.getElementById('status-fifo').textContent = 'Optional — Drop .xlsx or click';
      document.querySelectorAll('.zone-file-input').forEach(inp => { inp.value = ''; });
      updateAnalyzeButton();
      // Switch screens
      document.getElementById('dashboard-screen').classList.add('hidden');
      document.getElementById('upload-screen').classList.remove('hidden');
    });
  }

  function initSearchHandler() {
    document.getElementById('trade-search').addEventListener('input', (e) => {
      if (analysisResult) {
        tradeTablePage = 0;
        renderTradeTablePage(analysisResult.trades, 0, e.target.value);
      }
    });
  }

  function initCsvExportHandler() {
    document.getElementById('btn-export-csv').addEventListener('click', () => {
      if (!analysisResult || analysisResult.realizedEvents.length === 0) return;
      
      const csvData = [
        ['Date', 'Time', 'Symbol', 'Option Type', 'Strike', 'Qty', 'Entry Price', 'Exit Price', 'Realized P&L']
      ];
      
      analysisResult.realizedEvents.forEach(ev => {
        csvData.push([
          formatDate(ev.date),
          ev.tradeTime || '',
          ev.symbol,
          ev.optionType,
          ev.strikePrice > 0 ? ev.strikePrice : '',
          ev.qty,
          ev.entryPrice.toFixed(2),
          ev.exitPrice.toFixed(2),
          ev.pnl.toFixed(2)
        ]);
      });
      
      const ws = XLSX.utils.aoa_to_sheet(csvData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Realized_Trades");
      XLSX.writeFile(wb, "vyro_backoffice_trades.csv");
    });
  }

  // ──────────────────────────────────────────────
  // 9. INIT
  // ──────────────────────────────────────────────
  function initCalendarHandlers() {
    document.getElementById('cal-prev').addEventListener('click', () => {
      calCurrentMonth.month--;
      if (calCurrentMonth.month < 0) { calCurrentMonth.month = 11; calCurrentMonth.year--; }
      renderCalendarMonth();
    });
    document.getElementById('cal-next').addEventListener('click', () => {
      calCurrentMonth.month++;
      if (calCurrentMonth.month > 11) { calCurrentMonth.month = 0; calCurrentMonth.year++; }
      renderCalendarMonth();
    });
    document.getElementById('cal-filter-apply').addEventListener('click', applyCalFilter);
    document.getElementById('cal-filter-reset').addEventListener('click', () => {
      document.getElementById('cal-range-stats').classList.add('hidden');
      if (analysisResult && analysisResult.dailyPnl.length > 0) {
        const first = analysisResult.dailyPnl[0].date;
        const last = analysisResult.dailyPnl[analysisResult.dailyPnl.length - 1].date;
        document.getElementById('cal-date-from').value = first.toISOString().split('T')[0];
        document.getElementById('cal-date-to').value = last.toISOString().split('T')[0];
      }
    });
    document.getElementById('global-filter-apply').addEventListener('click', () => {
      applyGlobalFilterAndRender();
    });
  }
  function init() {
    initUploadHandlers();
    initTabHandlers();
    initEquityToggle();
    initResetButton();
    initSearchHandler();
    initCsvExportHandler();
    initCalendarHandlers();
  }

  // Run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
