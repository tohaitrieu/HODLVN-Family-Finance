/**
 * ===============================================
 * CHANGELOG_MANAGER.GS
 * ===============================================
 * 
 * Quản lý việc hiển thị và cập nhật Changelog
 */

// Changelog data with bilingual content
// Changelog data is now loaded from ChangelogData.gs

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
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CHANGELOG);
    
    // Nếu chưa có sheet thì tạo mới
    if (!sheet) {
      sheet = SheetInitializer.initializeChangelogSheet();
    }
    
    // Xóa dữ liệu cũ (trừ header)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    // Apply sheet formatting with new structure
    SheetUtils.applySheetFormat(sheet, 'CHANGELOG');
    
    // Set column widths for better readability
    sheet.setColumnWidth(1, 150); // Version
    sheet.setColumnWidth(2, 250); // Feature EN
    sheet.setColumnWidth(3, 250); // Feature VI
    sheet.setColumnWidth(4, 300); // Detail EN
    sheet.setColumnWidth(5, 300); // Detail VI
    sheet.setColumnWidth(6, 250); // Action EN
    sheet.setColumnWidth(7, 250); // Action VI
    
    // Ghi dữ liệu mới
    let currentRow = 2;
    
    CHANGELOG_HISTORY.forEach(entry => {
      // For each version entry
      if (entry.changesEn && entry.changesVi) {
        // New bilingual format
        const maxChanges = Math.max(entry.changesEn.length, entry.changesVi.length);
        
        for (let i = 0; i < maxChanges; i++) {
          const rowData = [
            i === 0 ? `v${entry.version}` : '', // Version only on first row
            entry.changesEn[i] || '', // Feature EN
            entry.changesVi[i] || '', // Feature VI
            entry.detailsEn && entry.detailsEn[i] || '', // Detail EN
            entry.detailsVi && entry.detailsVi[i] || '', // Detail VI
            entry.actionsEn && entry.actionsEn[i] || '', // Action EN
            entry.actionsVi && entry.actionsVi[i] || '' // Action VI
          ];
          
          sheet.getRange(currentRow, 1, 1, 7).setValues([rowData]);
          
          // Style the version cell
          if (i === 0) {
            sheet.getRange(currentRow, 1).setFontWeight('bold')
              .setBackground('#E2EFDA')
              .setFontColor('#38761D');
          }
          
          currentRow++;
        }
        
      } else {
        // Legacy format - convert to bilingual
        entry.changes.forEach((change, index) => {
          const rowData = [
            index === 0 ? `v${entry.version}` : '', // Version
            change, // Feature EN (use existing as EN)
            change, // Feature VI (duplicate for now)
            '', // Detail EN
            '', // Detail VI
            entry.actions && entry.actions[index] || '', // Action EN
            entry.actions && entry.actions[index] || '' // Action VI
          ];
          
          sheet.getRange(currentRow, 1, 1, 7).setValues([rowData]);
          
          if (index === 0) {
            sheet.getRange(currentRow, 1).setFontWeight('bold')
              .setBackground('#E2EFDA')
              .setFontColor('#38761D');
          }
          
          currentRow++;
        });
      }
      
      // Add empty row between versions
      currentRow++;
    });
    
    // Format with wrap text
    const dataRange = sheet.getRange(2, 1, currentRow - 2, 7);
    dataRange.setVerticalAlignment('top')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    // Apply borders
    dataRange.setBorder(true, true, true, true, true, true, 
      GLOBAL_SHEET_CONFIG.BORDER_COLOR, GLOBAL_SHEET_CONFIG.BORDER_STYLE);
    
    // Set row heights for better readability
    for (let row = 2; row < currentRow; row++) {
      sheet.setRowHeight(row, 35);
    }
  }
};
