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