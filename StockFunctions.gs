/**
 * ===============================================
 * STOCK FUNCTIONS & FINANCIAL INDICATORS
 * ===============================================
 */

// 1) Mapping timeframe → resolution cho TCBS API
const INTERVAL_MAP = {
  'D1': 'D',
  'W1': 'W',
  'M1': 'M',
  'H4': '240'
};

/**
 * TCBS_BARS: Lấy N bars (D1/W1/M1/H4) từ API TCBS
 *
 * @param {string} ticker    Mã (ví dụ "ACB", "^VNINDEX.VN")
 * @param {string} type      "stock" hoặc "index"
 * @param {string} timeframe "D1","W1","M1","H4"
 * @param {number} count     Số bars muốn lấy
 * @return Mảng 2D: [Date,Open,High,Low,Close,Volume]
 */
function TCBS_BARS(ticker, type, timeframe, count) {
  const tz     = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const key    = timeframe.toUpperCase().trim();
  const res    = INTERVAL_MAP[key];
  if (!res) throw new Error('Không hỗ trợ timeframe: ' + timeframe);

  const now = Math.floor(Date.now() / 1000);
  const url = [
    'https://apipubaws.tcbs.com.vn/stock-insight/v2/stock/bars?',
    'ticker=',     encodeURIComponent(ticker),
    '&type=',      encodeURIComponent(type),
    '&resolution=',encodeURIComponent(res),
    '&to=',        now,
    '&countBack=', count
  ].join('');

  const resp = UrlFetchApp.fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });
  const js   = JSON.parse(resp.getContentText());
  const data = js.data;
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('Không có dữ liệu trả về cho ' + ticker);
  }

  const out = [['Date','Open','High','Low','Close','Volume']];
  data.slice(-count).forEach(pt => {
    // chỉ lấy phần YYYY-MM-DD trước ký tự 'T'
    const dateStr = pt.tradingDate.split('T')[0];
    out.push([
      dateStr,
      pt.open,
      pt.high,
      pt.low,
      pt.close,
      pt.volume
    ]);
  });

  return out;
}
/**
 * HODLDATA: Lấy N bars lịch sử giá (daily/weekly/monthly)
 *  - Với D1: trả trực tiếp daily bars
 *  - Với W1: gom daily thành weekly bars (Monday as week start)
 *  - Với M1: gom daily thành monthly bars
 *
 * @param {string} symbol    Mã hoặc ô tham chiếu (ví dụ "VCB", D5, "VNINDEX")
 * @param {string} timeframe Khung: "D1","W1","M1"
 * @param {number} bars      Số thanh muốn lấy (ví dụ 600)
 * @return Mảng 2D: [Date,Open,High,Low,Close,Volume]
 */
function HODLDATA(symbol, timeframe, bars) {
  const sym = normalizeSymbol(symbol);
  const tz  = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const ua  = { "User-Agent": "Mozilla/5.0" };

  // 1) Fetch daily data via Spark v8 → fallback Chart v8
  let daily = [];
  try {
    const sparkUrl = [
      'https://query2.finance.yahoo.com/v8/finance/spark?symbols=', encodeURIComponent(sym),
      '&range=max&interval=1d',
      '&indicators=close,open,high,low,volume',
      '&includeTimestamps=true'
    ].join('');
    const r = UrlFetchApp.fetch(sparkUrl, { headers: ua });
    const j = JSON.parse(r.getContentText());
    const info = j.spark?.result?.[0] || j.result?.[0];
    if (info) {
      const blk = info.response?.[0] || info;
      const ts  = blk.timestamp  || blk.timestamps;
      const q   = blk.indicators?.quote?.[0];
      if (ts && q) {
        for (let i = 0; i < ts.length; i++) {
          const o = q.open[i], c = q.close[i];
          if (o==null||c==null) continue;
          daily.push([
            new Date(ts[i]*1000),
            o, q.high[i], q.low[i], c, q.volume[i]
          ]);
        }
      }
    }
  } catch(e){/* ignore spark errors */ }

  if (daily.length === 0) {
    // fallback Chart v8
    const chartUrl = [
      'https://query1.finance.yahoo.com/v8/finance/chart/', encodeURIComponent(sym),
      '?range=max&interval=1d&includePrePost=false&events=div%2Csplit'
    ].join('');
    const r2 = UrlFetchApp.fetch(chartUrl, { headers: ua });
    const j2 = JSON.parse(r2.getContentText());
    const res = j2.chart?.result?.[0];
    if (!res) throw new Error('Không fetch được daily data cho ' + sym);
    const ts2 = res.timestamp;
    const q2  = res.indicators.quote[0];
    for (let i = 0; i < ts2.length; i++) {
      const o = q2.open[i], c = q2.close[i];
      if (o==null||c==null) continue;
      daily.push([
        new Date(ts2[i]*1000),
        o, q2.high[i], q2.low[i], c, q2.volume[i]
      ]);
    }
  }

  // 2) Build output depending on timeframe
  let rows;
  if (timeframe.toUpperCase() === 'D1') {
    rows = daily;
  } else if (timeframe.toUpperCase() === 'W1') {
    rows = aggregateWeekly(daily, tz);
  } else if (timeframe.toUpperCase() === 'M1') {
    rows = aggregateMonthly(daily, tz);
  } else {
    throw new Error('Timeframe không hợp lệ: ' + timeframe);
  }

  // 3) Take last N bars and format Date
  const out = [['Date','Open','High','Low','Close','Volume']];
  const slice = rows.slice(-bars);
  slice.forEach(r => {
    out.push([
      Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'),
      r[1], r[2], r[3], r[4], r[5]
    ]);
  });
  if (slice.length === 0) throw new Error('Không có data để trả về cho ' + sym);
  return out;
}

/** Gom daily → weekly (week start = Monday) */
function aggregateWeekly(daily, tz) {
  const groups = {};
  daily.forEach(r => {
    const d = r[0];
    // compute Monday of that week
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    const m = new Date(d);
    m.setDate(diff);
    const key = m.getFullYear()+'-'+m.getMonth()+'-'+m.getDate();
    if (!groups[key]) {
      groups[key] = { date:m, open:r[1], high:r[2], low:r[3], close:r[4], volume:r[5] };
    } else {
      const g = groups[key];
      g.high   = Math.max(g.high,  r[2]);
      g.low    = Math.min(g.low,   r[3]);
      g.close  = r[4];
      g.volume = g.volume + r[5];
    }
  });
  const arr = Object.values(groups);
  arr.sort((a,b) => a.date - b.date);
  return arr.map(g => [ g.date, g.open, g.high, g.low, g.close, g.volume ]);
}

/** Gom daily → monthly */
function aggregateMonthly(daily, tz) {
  const groups = {};
  daily.forEach(r => {
    const d = r[0];
    const key = d.getFullYear()+'-'+d.getMonth();
    if (!groups[key]) {
      const firstOfMonth = new Date(d.getFullYear(), d.getMonth(), 1);
      groups[key] = { date:firstOfMonth, open:r[1], high:r[2], low:r[3], close:r[4], volume:r[5] };
    } else {
      const g = groups[key];
      g.high   = Math.max(g.high,  r[2]);
      g.low    = Math.min(g.low,   r[3]);
      g.close  = r[4];
      g.volume = g.volume + r[5];
    }
  });
  const arr = Object.values(groups);
  arr.sort((a,b) => a.date - b.date);
  return arr.map(g => [ g.date, g.open, g.high, g.low, g.close, g.volume ]);
}
/** map khung thời gian sang interval của Yahoo */
function mapTimeframeToInterval(tf) {
  const t = String(tf||'').toUpperCase().trim();
  const m = { 'D1':'1d', 'W1':'1wk', 'M1':'1mo' };
  if (m[t]) return m[t];
  throw new Error('Timeframe không hợp lệ: ' + tf);
}

/** seconds tương ứng với mỗi bar của timeframe */
function secondsPerBar(tf) {
  const t = String(tf||'').toUpperCase().trim();
  if (t === 'D1') return 24 * 3600;
  if (t === 'W1') return 7 * 24 * 3600;
  if (t === 'M1') return 30 * 24 * 3600;  // quy ước 30 ngày
  throw new Error('Timeframe không hợp lệ: ' + tf);
}

/** chuẩn hóa symbol Việt Nam */
function normalizeSymbol(input) {
  const s = String(input||'').toUpperCase().trim();
  const map = { 'VNINDEX':'^VNINDEX.VN', 'VN30':'FIVNM30.FGI' };
  if (map[s])              return map[s];
  if (s.startsWith('^')     || s.includes('.')) return s;
  return s + '.VN';
}
/**
 * HODLVALUE: Lấy metric fundamental từ Yahoo Finance qua v7/finance/quote
 *
 * @param {string} symbol Mã bạn nhập: "VCB", "VNINDEX", v.v.
 * @param {string} metric Tên metric: "bvps","pb","pe","eps"
 * @return Giá trị metric
 */
function HODLVALUE(symbol, metric) {
  const sym = normalizeSymbol(symbol);
  const url = 'https://query1.finance.yahoo.com/v7/finance/quote?symbols='
            + encodeURIComponent(sym);

  const resp = UrlFetchApp.fetch(url);
  const js   = JSON.parse(resp.getContentText());
  const q    = js.quoteResponse && js.quoteResponse.result && js.quoteResponse.result[0];
  if (!q) throw new Error('Không lấy được dữ liệu cho ' + sym);

  const m = String(metric||'').toLowerCase().trim();
  let val;
  switch(m) {
    case 'bvps':
      val = q.bookValue;
      break;
    case 'pb':
    case 'p/b':
      val = q.priceToBook;
      break;
    case 'pe':
    case 'p/e':
      val = q.trailingPE;
      break;
    case 'eps':
      val = q.epsTrailingTwelveMonths;
      break;
    default:
      throw new Error('Metric "'+metric+'" chưa được hỗ trợ!');
  }
  return val;
}

/**
 * PIVOTFIB – Tính Fibonacci Pivot động trên một vùng dữ liệu High–Low–Close,
 * tự reload khi dữ liệu thay đổi, thêm logic chọn row theo thứ trong tuần.
 *
 * @param {number[][]} dataRange  Vùng dữ liệu 3 cột: High, Low, Close (ví dụ: 'FXD1'!C2:E)
 * @return {number[][]}           7 dòng × 1 cột: [R3],[R2],[R1],[PP],[S1],[S2],[S3]
 * @customfunction
 */
function PIVOTFIB(dataRange) {
  // 1) Tách dữ liệu High, Low, Close và lọc số
  var highs = [], lows = [], closes = [];
  for (var i = 0; i < dataRange.length; i++) {
    var row = dataRange[i];
    if (row.length >= 3) {
      var h = row[0], l = row[1], c = row[2];
      if (typeof h === 'number' && typeof l === 'number' && typeof c === 'number') {
        highs.push(h);
        lows.push(l);
        closes.push(c);
      }
    }
  }

  var n = highs.length;
  if (n === 0) {
    // không có data thì thôi
    return [[null]];
  }

  // 2) Check ngày trong tuần
  var today = new Date();
  var d = today.getDay(); // 0=Chủ nhật, 6=Thứ bảy
  // cuối tuần → index = n-1; ngày thường → index = n-2
  var idx;
  if (d === 0 || d === 6) {
    // CN hoặc T7: lấy dòng cuối cùng
    idx = n - 1;
  } else {
    // T2–T6: cần ít nhất 2 bản ghi
    if (n < 2) return [[null]];
    idx = n - 2;
  }

  // 3) Lấy High, Low, Close theo idx vừa chọn
  var high  = highs[idx];
  var low   = lows[idx];
  var close = closes[idx];

  // 4) Tính PP và range
  var pp    = (high + low + close) / 3;
  var range = high - low;

  // 5) Tính các mức Fibonacci
  var R3 = pp + 1.000 * range;
  var R2 = pp + 0.618 * range;
  var R1 = pp + 0.382 * range;
  var S1 = pp - 0.382 * range;
  var S2 = pp - 0.618 * range;
  var S3 = pp - 1.000 * range;

  // 6) Trả về mảng 7×1
  return [
    [R3],
    [R2],
    [R1],
    [pp],
    [S1],
    [S2],
    [S3]
  ];
}
/**
 * ATR – Tính Average True Range (Wilder) động từ một phạm vi 3 cột [High, Low, Close], tự reload khi dữ liệu thay đổi.
 *
 * @param {number[][]} dataRange  Dãy 3 cột: High, Low, Close (ví dụ: GOLD-D1!C2:E)
 * @param {number} [period=14]    Số chu kỳ ATR (mặc định 14)
 * @return {number[][]} 1×1 cell chứa ATR cuối cùng
 * @customfunction
 */
function ATR(dataRange, period) {
  // 1) Xác định period mặc định
  period = (typeof period === 'number' && period > 0) ? Math.floor(period) : 14;

  // 2) Tách 3 cột High, Low, Close
  var highs = [], lows = [], closes = [];
  dataRange.forEach(function(row) {
    if (row.length >= 3) {
      var h = row[0], l = row[1], c = row[2];
      if ([h, l, c].every(function(v) { return typeof v === 'number'; })) {
        highs.push(h);
        lows.push(l);
        closes.push(c);
      }
    }
  });

  var n = highs.length;
  if (n < period + 1) {
    return [[null]];
  }

  // 3) Tính True Range
  var tr = [];
  for (var i = 1; i < n; i++) {
    var h = highs[i], l = lows[i], pc = closes[i - 1];
    tr.push(Math.max(h - l, Math.abs(h - pc), Math.abs(l - pc)));
  }

  // 4) Wilder smoothing
  var sum = 0;
  for (var i = 0; i < period; i++) sum += tr[i];
  var atr = sum / period;
  for (var i = period; i < tr.length; i++) {
    atr = (atr * (period - 1) + tr[i]) / period;
  }

  // 5) Trả về ATR cuối cùng
  return [[atr]];
}
/**
 * STOCHASTIC – Tính Slow Stochastic %K và %D động, tự reload khi dữ liệu thay đổi.
 * Hỗ trợ:
 *  - Range 4 cột [Date, High, Low, Close] (sẽ tự sort theo Date)
 *  - Range 3 cột [High, Low, Close]
 *
 * @param {Array} dataRange    Dãy 3 hoặc 4 cột.
 * @param {number} [kPeriod=14] Số chu kỳ %K (mặc định 14).
 * @param {number} [dPeriod=3]  Số chu kỳ trung bình cho %K và %D (mặc định 3).
 * @return 2 dòng × 1 cột: [%K cuối], [%D cuối]
 * @customfunction
 */
function STOCHASTIC(dataRange, kPeriod, dPeriod) {
  // Thiết lập default
  kPeriod = (typeof kPeriod === 'number' && kPeriod > 0) ? Math.floor(kPeriod) : 14;
  dPeriod = (typeof dPeriod === 'number' && dPeriod > 0) ? Math.floor(dPeriod) : 3;

  // Lọc rows chỉ chứa số (lọc blank và text)
  var rows = dataRange.filter(function(r) {
    var hasDate = r.length >= 4 && r[0] instanceof Date;
    var idx = hasDate ? 1 : 0;
    return r.length >= idx + 3
        && typeof r[idx]   === 'number'
        && typeof r[idx+1] === 'number'
        && typeof r[idx+2] === 'number';
  });
  if (rows.length < kPeriod) return [[null],[null]];

  // Sort theo Date nếu có cột Date
  var hasDate = rows[0].length >= 4 && rows[0][0] instanceof Date;
  if (hasDate) {
    rows.sort(function(a, b) {
      return new Date(a[0]) - new Date(b[0]);
    });
  }

  // Chuẩn hóa thành mảng High, Low, Close
  var highs = [], lows = [], closes = [];
  rows.forEach(function(r) {
    var idx = hasDate ? 1 : 0;
    highs.push(r[idx]);
    lows.push(r[idx+1]);
    closes.push(r[idx+2]);
  });
  var n = closes.length;

  // 1) Tính raw %K
  var rawK = Array(n).fill(null);
  for (var i = kPeriod - 1; i < n; i++) {
    var windowHigh = Math.max.apply(null, highs.slice(i - kPeriod + 1, i + 1));
    var windowLow  = Math.min.apply(null, lows.slice(i - kPeriod + 1, i + 1));
    rawK[i] = windowHigh > windowLow
      ? ((closes[i] - windowLow) / (windowHigh - windowLow) * 100)
      : 0;
  }

  // 2) Tính Slow %K = SMA(rawK, dPeriod)
  var slowK = Array(n).fill(null);
  for (var j = (kPeriod - 1) + (dPeriod - 1); j < n; j++) {
    var segment = rawK.slice(j - dPeriod + 1, j + 1);
    if (segment.every(function(v) { return v != null; })) {
      slowK[j] = segment.reduce(function(a, b) { return a + b; }, 0) / dPeriod;
    }
  }

  // 3) Tính Slow %D = SMA(slowK, dPeriod)
  var slowD = Array(n).fill(null);
  for (var k = (kPeriod - 1) + 2 * (dPeriod - 1); k < n; k++) {
    var seg2 = slowK.slice(k - dPeriod + 1, k + 1);
    if (seg2.every(function(v) { return v != null; })) {
      slowD[k] = seg2.reduce(function(a, b) { return a + b; }, 0) / dPeriod;
    }
  }

  // Trả về 2 giá trị cuối cùng
  var idxLast = n - 1;
  return [ [ slowK[idxLast] ], [ slowD[idxLast] ] ];
}
function STOCH_TREND(k, d) {
  var kVal = parseFloat(k), dVal = parseFloat(d);
  if (isNaN(kVal) || isNaN(dVal)) return 'Giá trị không hợp lệ';
  
  // 1) Quá mua / Quá bán
  if (kVal > 80 && dVal > 80) return 'Quá mua';
  if (kVal < 20 && dVal < 20) return 'Quá bán';
  
  // 2) Mua mạnh / Mua
  if (kVal > dVal && kVal > 70) return 'Mua mạnh';
  if (kVal > dVal && kVal > 50) return 'Mua';
  
  // 3) Bán mạnh / Bán
  if (kVal < dVal && kVal < 30) return 'Bán mạnh';
  if (kVal < dVal && kVal < 50) return 'Bán';
  
  // 4) Chỉnh giảm (pullback in downtrend)
  if (kVal > dVal && kVal < 50) return 'Chỉnh giảm';
  
  // 5) Chỉnh tăng (pullback in uptrend)
  if (kVal < dVal && kVal > 50) return 'Chỉnh tăng';
  
  // 6) Còn lại: tích luỹ
  return 'Tích luỹ';
}
/**
 * @customfunction
 * Tính Wilder RSI(14) cho mảng giá và trả về giá trị RSI cuối cùng.
 *
 * @param {number[][]} priceRange Mảng giá Close (ví dụ B2:B).
 * @return {number[][]} 1 dòng × 1 cột: [RSI14 cuối cùng].
 */
function RSI(priceRange) {
  const prices = priceRange.flat().filter(v => typeof v === 'number');
  const period = 14, n = prices.length;
  if (n <= period) return [[null]];
  const gains = [], losses = [];
  for (let i = 1; i < n; i++) {
    const d = prices[i] - prices[i-1];
    gains.push(d > 0 ? d : 0);
    losses.push(d < 0 ? -d : 0);
  }
  let avgG = gains.slice(0, period).reduce((a,b) => a + b, 0) / period;
  let avgL = losses.slice(0, period).reduce((a,b) => a + b, 0) / period;
  for (let i = period; i < gains.length; i++) {
    avgG = (avgG*(period-1) + gains[i]) / period;
    avgL = (avgL*(period-1) + losses[i]) / period;
  }
  const rsi = 100 - 100/(1 + avgG/avgL);
  return [[ rsi ]];
}
/**
 * @customfunction
 * RSI_TREND – Đánh giá xu hướng từ giá trị RSI.
 *
 *   >68        → "Quá mua"
 *   62–68      → "Mua mạnh"
 *   55–<62     → "Mua"
 *   45–<55     → "Tích lũy"
 *   38–<45     → "Bán"
 *   32–<38     → "Bán mạnh"
 *   ≤32        → "Quá bán"
 *
 * @param {number} rsiValue
 * @return {string}
 */
function RSI_TREND(rsiValue) {
  if (typeof rsiValue !== 'number' || isNaN(rsiValue)) return '';
  if (rsiValue > 68) {
    return 'Quá mua';
  }
  if (rsiValue >= 62) {        // 62 ≤ x ≤ 68
    return 'Mua mạnh';
  }
  if (rsiValue >= 55) {        // 55 ≤ x < 62
    return 'Mua';
  }
  if (rsiValue >= 45) {        // 45 ≤ x < 55
    return 'Tích lũy';
  }
  if (rsiValue >= 38) {        // 38 ≤ x < 45
    return 'Bán';
  }
  if (rsiValue > 32) {         // 32 < x < 38
    return 'Bán mạnh';
  }
  // còn lại x ≤ 32
  return 'Quá bán';
}
/**
 * EMA – Tính giá trị Exponential Moving Average cho chu kỳ tuỳ biến
 * và trả về EMA cuối cùng.
 *
 * @param {number[][]} priceRange  Vùng dữ liệu 1 cột chứa giá (ví dụ 'FXD1'!E2:E).
 * @param {number} period          Số chu kỳ EMA (ví dụ 10 cho EMA10).
 * @return {number}               Giá trị EMA cuối cùng hoặc null nếu không đủ dữ liệu.
 * @customfunction
 */
function EMA(priceRange, period) {
  // 1) Validate period
  if (typeof period !== 'number' || period < 1) {
    throw new Error('EMA: period phải là số >= 1');
  }
  period = Math.floor(period);

  // 2) Flatten và lọc giá numeric
  var flat = priceRange.reduce(function(acc, row) {
    return acc.concat(row);
  }, []);
  var prices = flat.filter(function(v) {
    return typeof v === 'number' && !isNaN(v);
  });

  // 3) Kiểm tra đủ dữ liệu
  if (prices.length < period) {
    return null;
  }

  // 4) Tính EMA
  var alpha = 2 / (period + 1);
  // Seed bằng SMA của period đầu
  var sum = 0;
  for (var i = 0; i < period; i++) {
    sum += prices[i];
  }
  var ema = sum / period;
  // Smoothing
  for (var j = period; j < prices.length; j++) {
    ema = alpha * prices[j] + (1 - alpha) * ema;
  }

  // 5) Trả về giá trị cuối cùng
  return ema;
}
/**
 * @customfunction
 * Tính MACD(19,39,9) cho mảng giá và trả về 2 giá trị cuối cùng dọc: [MACD line], [Signal line].
 *
 * @param {number[][]} priceRange Mảng giá Close (ví dụ B2:B).
 * @return {number[][]} 2 dòng × 1 cột: [MACD line], [Signal line].
 */
function MACD(priceRange) {
  const prices = priceRange.flat().filter(v => typeof v === 'number');
  const fastP = 19, slowP = 39, sigP = 9;
  const αf = 2/(fastP+1), αs = 2/(slowP+1), ασ = 2/(sigP+1);
  const n = prices.length;
  if (n < slowP) return [[null],[null]];
  const fastEMA = Array(n).fill(null), slowEMA = Array(n).fill(null);
  let sumF = 0;
  for (let i = 0; i < fastP; i++) sumF += prices[i];
  fastEMA[fastP-1] = sumF/fastP;
  for (let i = fastP; i < n; i++) {
    fastEMA[i] = αf*prices[i] + (1-αf)*fastEMA[i-1];
  }
  let sumS = 0;
  for (let i = 0; i < slowP; i++) sumS += prices[i];
  slowEMA[slowP-1] = sumS/slowP;
  for (let i = slowP; i < n; i++) {
    slowEMA[i] = αs*prices[i] + (1-αs)*slowEMA[i-1];
  }
  const macd = fastEMA.map((v,i) => (v!=null && slowEMA[i]!=null) ? v - slowEMA[i] : null);
  const signal = Array(n).fill(null);
  const startSig = slowP-1 + (sigP-1);
  let sumSig = 0, cnt = 0;
  for (let i = slowP-1; i <= startSig; i++) {
    if (macd[i] != null) { sumSig += macd[i]; cnt++; }
  }
  if (cnt) signal[startSig] = sumSig/cnt;
  for (let i = startSig+1; i < n; i++) {
    signal[i] = macd[i]!=null ? ασ*macd[i] + (1-ασ)*signal[i-1] : null;
  }
  const last = n-1;
  return [
    [ macd[last] ],
    [ signal[last] ]
  ];
}
/**
 * Trả về xu hướng dựa trên MACD line và Signal line.
 *
 * @customfunction
 * @param {number} macd Giá trị MACD line (ô K12).
 * @param {number} signal Giá trị Signal line (ô K13).
 * @return {string} "Mua mạnh", "Bán mạnh", "Chỉnh giảm", "Chỉnh tăng" hoặc "" nếu không thỏa điều kiện.
 */
function MACD_TREND(macd, signal) {
  // nếu trống hoặc không phải số thì trả về blank
  if (macd === '' || signal === '' || macd == null || signal == null) {
    return '';
  }
  macd = parseFloat(macd);
  signal = parseFloat(signal);
  if (isNaN(macd) || isNaN(signal)) {
    return '';
  }

  // 1) Mua mạnh: MACD > Signal và SIGNAL > 0
  if (macd > signal && signal > 0) {
    return 'Mua mạnh';
  }
  // 2) Bán mạnh: MACD < Signal và SIGNAL < 0
  if (macd < signal && signal < 0) {
    return 'Bán mạnh';
  }
  // 3) Chỉnh giảm: MACD > Signal nhưng SIGNAL <= 0
  if (macd > signal && macd <= 0) {
    return 'Chỉnh giảm';
  }
  // 4) Chỉnh tăng: MACD < Signal nhưng SIGNAL >= 0
  if (macd < signal && macd >= 0) {
    return 'Chỉnh tăng';
  }
  // 5) Các trường hợp khác (MACD == Signal hoặc không rõ xu hướng)
  return '';
}
/**
 * MPRICE: Lấy thị giá hiện tại từ TCBS API Bars (1-phút last-close)
 *
 * @param {string|Range} input Mã chứng khoán hoặc ô chứa mã (ví dụ "VCB" hoặc "VNINDEX")
 * @return {number} Giá đóng của bar 1-phút gần nhất
 * @customfunction
 */
function MPRICE(input) {
  // 1) Lấy raw input,uppercase
  const raw = (input && input.getValue) ? input.getValue() : input;
  if (!raw) return 0;
  const symbol = String(raw).trim().toUpperCase();

  // 2) Danh sách chỉ số
  const indexList = ['VNINDEX', 'VN30', 'VNMIDCAP', 'UPCOM'];

  // 3) Xác định type & ticker
  let type, ticker;
  if (indexList.includes(symbol)) {
    type = 'index';
    ticker = symbol;            // chỉ VNINDEX, VN30,...
  } else {
    type = 'stock';
    ticker = symbol;            // mã cổ phiếu bình thường
  }

  // 4) Build URL
  const now = Math.floor(Date.now() / 1000);
  const url = [
    'https://apipubaws.tcbs.com.vn/stock-insight/v2/stock/bars?',
    'ticker=',     encodeURIComponent(ticker),
    '&type=',      encodeURIComponent(type),
    '&resolution=1',
    '&to=',        now,
    '&countBack=1'
  ].join('');

  // 5) Fetch & parse
  try {
    const resp = UrlFetchApp.fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });
    const js   = JSON.parse(resp.getContentText());
    const data = js.data;
    if (!Array.isArray(data) || data.length === 0) {
      return 0; // Return 0 instead of error to avoid #ERROR! in sheet
    }
    // 6) Trả về close
    return data[0].close;
  } catch (e) {
    return 0;
  }
}
/**
 * AVERAGE_DOWN: Tính số cổ phiếu cần mua thêm và vốn cần để bình quân giá
 *
 * @param {number} existingQty   Số cổ phiếu đang nắm (ví dụ 6500)
 * @param {number} existingPrice Giá vốn hiện tại (ví dụ 20000)
 * @param {number} newBuyPrice   Giá bạn sẽ mua thêm (ví dụ 15000)
 * @param {number} targetPrice   Giá trung bình mong muốn (ví dụ 17650)
 * @return Hai giá trị: [số cổ phiếu cần mua, vốn cần thiết]
 * @customfunction
 */
function AVERAGE_DOWN(existingQty, existingPrice, newBuyPrice, targetPrice) {
  // Điều kiện hợp lệ
  if (targetPrice >= existingPrice) throw new Error('targetPrice phải nhỏ hơn existingPrice');
  if (newBuyPrice >= targetPrice)   throw new Error('newBuyPrice phải nhỏ hơn targetPrice');
  // Q1 = Q0*(P0 - Pt)/(Pt - P1)
  const addQty = existingQty * (existingPrice - targetPrice) / (targetPrice - newBuyPrice);
  const capital = addQty * newBuyPrice;
  return [ addQty, capital ];
}
