/**
 * ===============================================
 * ID GENERATOR
 * ===============================================
 * 
 * Helper để tạo Semantic ID (ID có ý nghĩa)
 * Format: [KEYWORD][YYYYMMDD]
 */

var IDGenerator = {
  
  /**
   * Tạo ID từ tên và ngày
   * @param {string} name - Tên (VD: Vay mua xe)
   * @param {Date} date - Ngày giao dịch
   * @return {string} ID (VD: VAYMUAXE20251122)
   */
  generate: function(name, date) {
    if (!name) return Utilities.getUuid();
    
    // 1. Chuẩn hóa tên: Uppercase, bỏ dấu, bỏ khoảng trắng
    let keyword = this.normalizeString(name);
    
    // Giới hạn độ dài keyword (VD: 20 ký tự) để ID không quá dài
    if (keyword.length > 20) {
      keyword = keyword.substring(0, 20);
    }
    
    // 2. Chuẩn hóa ngày: YYYYMMDD
    const d = new Date(date);
    const yyyy = d.getFullYear();
    const mm = (d.getMonth() + 1).toString().padStart(2, '0');
    const dd = d.getDate().toString().padStart(2, '0');
    
    return `${keyword}${yyyy}${mm}${dd}`;
  },
  
  /**
   * Tạo ID cho kỳ trả nợ (Suffix)
   * @param {string} parentId - ID của khoản nợ gốc
   * @param {number} count - Số lần đã trả (để tính suffix tiếp theo)
   * @return {string} ID (VD: VAYMUAXE20251122-01)
   */
  generateSuffix: function(parentId, count) {
    const suffix = (count + 1).toString().padStart(2, '0');
    return `${parentId}-${suffix}`;
  },
  
  /**
   * Helper: Bỏ dấu tiếng Việt và ký tự đặc biệt
   */
  normalizeString: function(str) {
    str = str.toString().toUpperCase();
    str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
    str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
    str = str.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I");
    str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
    str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
    str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
    str = str.replace(/Đ/g, "D");
    // Bỏ ký tự đặc biệt, chỉ giữ chữ và số
    str = str.replace(/[^A-Z0-9]/g, "");
    return str;
  }
};
