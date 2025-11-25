/**
 * ===============================================
 * CHANGELOG_MANAGER.GS
 * ===============================================
 * 
 * Quản lý việc hiển thị và cập nhật Changelog
 */

// Changelog data with bilingual content
const CHANGELOG_HISTORY = [
  {
    version: '5.0',
    date: '2024-12',
    changesEn: [
      'Dashboard Redesign',
      'Custom Function Auto-refresh',
      'Form UI Enhancement',
      'Budget Planning Module',
      'Bilingual Changelog'
    ],
    changesVi: [
      'Thiết kế lại Dashboard',
      'Tự động làm mới hàm tùy chỉnh',
      'Cải thiện giao diện Form',
      'Module lập ngân sách',
      'Lịch sử cập nhật song ngữ'
    ],
    detailsEn: [
      'New 2x2 grid layout with financial summary, charts, and event tables',
      'Functions update automatically when filters change or data is added',
      '2-column forms with HODL.VN branding and rules explanation',
      'Added comprehensive budget tracking with category allocation',
      'Support for English/Vietnamese in changelog documentation'
    ],
    detailsVi: [
      'Bố cục lưới 2x2 mới với tóm tắt tài chính, biểu đồ và bảng sự kiện',
      'Các hàm tự động cập nhật khi thay đổi bộ lọc hoặc thêm dữ liệu',
      'Form 2 cột với thương hiệu HODL.VN và giải thích quy tắc',
      'Thêm theo dõi ngân sách toàn diện với phân bổ danh mục',
      'Hỗ trợ tiếng Anh/Việt trong tài liệu lịch sử cập nhật'
    ],
    actionsEn: [
      'Click Menu > Update Dashboard to apply new layout',
      'No action needed - automatic',
      'Forms will use new design automatically',
      'Check BUDGET sheet for planning',
      'View this sheet for updates'
    ],
    actionsVi: [
      'Click Menu > Cập nhật Dashboard để áp dụng bố cục mới',
      'Không cần thao tác - tự động',
      'Form sẽ tự động dùng thiết kế mới',
      'Xem sheet BUDGET để lập kế hoạch',
      'Xem sheet này để cập nhật'
    ]
  },
  {
    version: '4.0',
    date: '2024-11',
    changesEn: [
      'Debt Management Enhancement',
      'Investment Tracking',
      'Performance Optimization'
    ],
    changesVi: [
      'Cải thiện quản lý nợ',
      'Theo dõi đầu tư',
      'Tối ưu hiệu suất'
    ],
    detailsEn: [
      'Added loan types, interest calculation, and payment scheduling',
      'Stock, Gold, Crypto portfolio with real-time pricing',
      'Faster data processing and reduced memory usage'
    ],
    detailsVi: [
      'Thêm loại vay, tính lãi và lịch thanh toán',
      'Danh mục CK, Vàng, Crypto với giá thời gian thực',
      'Xử lý dữ liệu nhanh hơn và giảm bộ nhớ'
    ],
    actionsEn: [
      'Review debt entries in DEBT MANAGEMENT sheet',
      'Add investments via respective forms',
      'No action required'
    ],
    actionsVi: [
      'Xem lại các khoản nợ trong sheet QUẢN LÝ NỢ',
      'Thêm đầu tư qua các form tương ứng',
      'Không cần thao tác'
    ]
  }
];

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
