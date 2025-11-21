/**
 * MPRICE: Lấy thị giá hiện tại từ TCBS API Bars (1-phút last-close)
 *
 * @param {string|Range} input Mã chứng khoán hoặc ô chứa mã (ví dụ "VCB" hoặc "VNINDEX")
 * @return {number} Giá đóng của bar 1-phút gần nhất
 * @customfunction
 */
function MPRICE(input) {
  // 1) Lấy raw input,uppercase
  const raw = (input.getValue) ? input.getValue() : input;
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
  const resp = UrlFetchApp.fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });
  const js   = JSON.parse(resp.getContentText());
  const data = js.data;
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error(`Không lấy được thị giá TCBS cho ${ticker}`);
  }

  // 6) Trả về close
  return data[0].close;
}
/**
 * CPRICE – Lấy giá hiện tại từ Yahoo Chart API v8 cho Forex hoặc Crypto.
 * Sử dụng Chart API v8:
 *   https://query2.finance.yahoo.com/v8/finance/chart/{symbol}
 *
 * @param {string|Range} input   Chuỗi symbol ("EURUSD", "BTCUSD", "DOGEUSD", …) hoặc tham chiếu ô.
 * @return {number}              Giá hiện tại (regularMarketPrice).
 * @customfunction
 */
function CPRICE(input) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1) Xác định symbol từ input
  var sym = input;
  if (typeof input === 'string' && input.includes('!')) {
    var parts = input.split('!');
    var sheet = ss.getSheetByName(parts[0]);
    if (!sheet) throw 'CPRICE: Không tìm thấy sheet "' + parts[0] + '"';
    sym = sheet.getRange(parts[1]).getValue();
  } else if (typeof input !== 'string') {
    sym = input.getValue();
  }
  sym = String(sym).trim().toUpperCase();
  if (!sym) throw 'CPRICE: Symbol trống';

  // 2) Chuẩn hóa symbol
  // Nếu kết thúc bằng "USD", thử dạng crypto (base-USD)
  var cryptoSym = null;
  if (sym.endsWith('USD')) {
    var base = sym.slice(0, sym.length - 3);
    cryptoSym = base + '-USD';
  }
  // Đồng thời chuẩn bị forex dạng XXXYYY=X
  var forexSym = sym + '=X';

  // 3) Danh sách ưu tiên lấy dữ liệu: crypto nếu có, sau đó forex
  var symbolList = cryptoSym ? [cryptoSym, forexSym] : [forexSym];

  // 4) Gọi Yahoo Chart API v8
  for (var i = 0; i < symbolList.length; i++) {
    var s = encodeURIComponent(symbolList[i]);
    var url = 'https://query2.finance.yahoo.com/v8/finance/chart/' + s +
              '?range=1d&interval=1d';
    try {
      var resp = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      var js = JSON.parse(resp.getContentText());
      var meta = js.chart && js.chart.result && js.chart.result[0] && js.chart.result[0].meta;
      if (meta && typeof meta.regularMarketPrice === 'number') {
        return meta.regularMarketPrice;
      }
    } catch (e) {
      // Bỏ qua lỗi, thử symbol tiếp theo
    }
  }
  throw 'CPRICE: Không lấy được giá cho ' + sym;
}
/**
 * Hàm lấy giá vàng mua vào (pb) từ BTMC.
 * Tự động tìm kiếm trong tất cả các dòng dữ liệu.
 * @param {string} tu_khoa Tên loại vàng (Ví dụ: "SJC", "Nhẫn", "Gold").
 * @return Giá mua vào.
 * @customfunction
 */
function GPRICE(tu_khoa) {
  if (!tu_khoa) return "Chưa nhập tên";
  
  var keyword = tu_khoa.toString().toLowerCase();

  // 1. Quy đổi từ khóa chung sang tên cụ thể của BTMC
  if (keyword === "gold" || keyword === "xauusd" || keyword === "vàng") {
    keyword = "vàng rồng thăng long"; 
  }
  
  // 2. Lấy dữ liệu (có lưu cache 60s)
  var cache = CacheService.getScriptCache();
  var xmlContent = cache.get("btmc_data_cache_v2"); // Đổi tên cache để làm mới
  
  if (!xmlContent) {
    var url = "http://api.btmc.vn/api/BTMCAPI/getpricebtmc?key=3kd8ub1llcg9t45hnoh8hmn7t5kc2v";
    try {
      var response = UrlFetchApp.fetch(url);
      xmlContent = response.getContentText();
      cache.put("btmc_data_cache_v2", xmlContent, 60); 
    } catch (e) {
      return "Lỗi API";
    }
  }

  // 3. Xử lý tìm kiếm thông minh (Bỏ qua số thứ tự dòng n_1, n_2...)
  var rawRows = xmlContent.split('<Data');

  // Duyệt từ trên xuống dưới (Dữ liệu mới nhất thường ở trên cùng)
  for (var i = 1; i < rawRows.length; i++) {
    var line = rawRows[i];
    
    // Tìm tên vàng: Bắt pattern n_số="giá_trị"
    var nameMatch = line.match(/n_\d+="([^"]*)"/);
    var nameInApi = nameMatch ? nameMatch[1].toLowerCase() : "";

    // Nếu tên trong API có chứa từ khóa bạn nhập
    if (nameInApi.indexOf(keyword) > -1) {
      
      // Tìm giá mua vào (pb): Bắt pattern pb_số="giá_trị"
      var priceMatch = line.match(/pb_\d+="([^"]*)"/);
      var price = priceMatch ? priceMatch[1] : "0";
      
      return parseFloat(price);
    }
  }

  return "Không tìm thấy";
}