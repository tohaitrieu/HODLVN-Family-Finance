/**
 * ===============================================
 * CHANGELOG_MANAGER.GS
 * ===============================================
 * 
 * Quản lý việc hiển thị và cập nhật Changelog
 */

const ChangelogManager = {
  
  /**
   * Kiểm tra phiên bản và cập nhật Changelog nếu cần
   */
  checkVersionAndUpdate() {
    const props = PropertiesService.getScriptProperties();
    const currentVersion = APP_CONFIG.VERSION;
    const lastVersion = props.getProperty('LAST_VERSION');
    
    // Nếu phiên bản thay đổi hoặc chưa có changelog
    if (currentVersion !== lastVersion) {
      this.updateChangelogSheet();
      props.setProperty('LAST_VERSION', currentVersion);
      
      // Thông báo cho người dùng
      SpreadsheetApp.getUi().alert(
        'Cập nhật mới!',
        `Hệ thống đã được cập nhật lên phiên bản v${currentVersion}.\n` +
        'Vui lòng xem sheet "LỊCH SỬ CẬP NHẬT" để biết chi tiết thay đổi.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  },
  
  /**
   * Cập nhật sheet Changelog
   */
  updateChangelogSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CHANGELOG);
    
    // Nếu chưa có sheet thì tạo mới
    if (!sheet) {
      sheet = SheetInitializer.initializeChangelogSheet();
    }
    
    // Xóa dữ liệu cũ (trừ header)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    // Ghi dữ liệu mới
    let currentRow = 2;
    
    CHANGELOG_HISTORY.forEach(entry => {
      // 1. Dòng Header của Version
      const versionHeader = `v${entry.version} - ${entry.title} (${entry.date})`;
      sheet.getRange(currentRow, 1, 1, 3).merge().setValue(versionHeader)
        .setFontWeight('bold')
        .setBackground('#E2EFDA') // Xanh nhạt
        .setFontColor('#38761D'); // Xanh đậm
      
      currentRow++;
      
      // 2. Danh sách thay đổi
      entry.changes.forEach(change => {
        sheet.getRange(currentRow, 1).setValue(''); // Cột A trống
        sheet.getRange(currentRow, 2).setValue(change)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        
        // Nếu có actions tương ứng (logic đơn giản: map theo index hoặc gom chung)
        // Ở đây ta gom chung actions vào dòng cuối hoặc hiển thị riêng
        sheet.getRange(currentRow, 3).setValue('');
        
        currentRow++;
      });
      
      // 3. Hành động khuyến nghị (nếu có)
      if (entry.actions && entry.actions.length > 0) {
        sheet.getRange(currentRow, 1).setValue('');
        sheet.getRange(currentRow, 2).setValue('⚡ HÀNH ĐỘNG CẦN THIẾT:')
          .setFontWeight('bold')
          .setFontColor('#C00000');
          
        sheet.getRange(currentRow, 3).setValue(entry.actions.join('\n'))
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setFontColor('#C00000');
          
        currentRow++;
      }
      
      // Dòng trống ngăn cách
      currentRow++;
    });
    
    // Format lại border cho đẹp
    const dataRange = sheet.getRange(2, 1, currentRow - 2, 3);
    dataRange.setVerticalAlignment('top');
  }
};
