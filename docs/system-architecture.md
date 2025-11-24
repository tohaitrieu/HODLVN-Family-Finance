# System Architecture

## Overview

HODLVN-Family-Finance employs a modular, cloud-native architecture built entirely on Google's infrastructure. This document details the system design, component interactions, data flow, and technical implementation patterns.

## 🏗️ Architectural Overview

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        USER INTERFACE LAYER                     │
├─────────────────┬─────────────────┬─────────────────┬───────────┤
│   Web Browser   │  Mobile Apps    │  Google Sheets  │  Sharing  │
│   (Desktop)     │  (iOS/Android)  │   Native UI     │  Features │
└─────────────────┴─────────────────┴─────────────────┴───────────┘
                                │
┌─────────────────────────────────────────────────────────────────┐
│                     PRESENTATION LAYER                          │
├─────────────────┬─────────────────┬─────────────────┬───────────┤
│  HTML Forms     │   Dashboard     │  Menu System    │  Reports  │
│  (Input/Edit)   │  (Analytics)    │  (Navigation)   │  (Output) │
└─────────────────┴─────────────────┴─────────────────┴───────────┘
                                │
┌─────────────────────────────────────────────────────────────────┐
│                      APPLICATION LAYER                          │
├──────────────────┬──────────────────┬──────────────────┬────────┤
│  Main Controller │  Data Processors │  Business Logic  │ Utils  │
│  (Routing)       │  (Transactions)  │  (Validation)    │ (Libs) │
└──────────────────┴──────────────────┴──────────────────┴────────┘
                                │
┌─────────────────────────────────────────────────────────────────┐
│                      INTEGRATION LAYER                          │
├─────────────────┬─────────────────┬─────────────────┬───────────┤
│   Google APIs   │  External APIs  │    Triggers     │  Security │
│   (Sheets/Drive)│  (Stock/Gold)   │  (Automation)   │  (Auth)   │
└─────────────────┴─────────────────┴─────────────────┴───────────┘
                                │
┌─────────────────────────────────────────────────────────────────┐
│                         DATA LAYER                              │
├─────────────────┬─────────────────┬─────────────────┬───────────┤
│ Google Sheets   │    Formulas     │   Properties    │  Cache    │
│ (Primary Store) │ (Calculations)  │   (Settings)    │ (Temp)    │
└─────────────────┴─────────────────┴─────────────────┴───────────┘
```

### Architectural Principles

1. **Serverless-First**: Leverages Google Apps Script serverless runtime
2. **Data-Centric**: Google Sheets as the primary data store and UI
3. **Event-Driven**: Trigger-based automation and real-time updates
4. **Modular Design**: Clear separation of concerns across modules
5. **Stateless Operations**: Each function is independent and reusable
6. **Mobile-Responsive**: Works across all device types

## 📦 Component Architecture

### Layer 1: User Interface Layer

#### Google Sheets Native UI
```javascript
// Sheet-based interface components
const UI_COMPONENTS = {
  DASHBOARD: 'Real-time analytics and summary views',
  DATA_SHEETS: 'Transaction tables with built-in filtering',
  CHARTS: 'Visual representations of financial data',
  CONDITIONAL_FORMATTING: 'Color-coded alerts and status indicators'
};
```

#### HTML Service Forms
```javascript
// Dynamic form components
const FORM_COMPONENTS = {
  SETUP_WIZARD: 'Multi-step user onboarding',
  TRANSACTION_FORMS: 'Input forms for each transaction type',
  BUDGET_MANAGER: 'Budget configuration interface',
  SETTINGS_PANEL: 'System configuration options'
};
```

### Layer 2: Presentation Layer

#### Menu System (Main.gs)
```javascript
// Hierarchical menu structure
const MENU_ARCHITECTURE = {
  ROOT_MENU: 'HODLVN Finance',
  SUBMENUS: {
    TRANSACTIONS: ['Income', 'Expense', 'Investment', 'Debt'],
    BUDGETS: ['Setup', 'Reports', 'Analysis'],
    DASHBOARD: ['Overview', 'Refresh', 'Settings']
  }
};

// Menu dispatcher pattern
function menuDispatcher(action, parameters) {
  const handlers = {
    'show-income-form': () => showIncomeForm(),
    'show-expense-form': () => showExpenseForm(),
    'refresh-dashboard': () => updateDashboard(),
    // ... other handlers
  };
  
  return handlers[action]?.(parameters);
}
```

#### Dashboard System (DashboardManager.gs)
```javascript
// Dashboard component architecture
const DASHBOARD_COMPONENTS = {
  SUMMARY_CARDS: {
    TOTAL_INCOME: 'Sum of all income this month',
    TOTAL_EXPENSE: 'Sum of all expenses this month', 
    NET_CASH_FLOW: 'Income minus expenses minus investments',
    BUDGET_STATUS: 'Overall budget utilization percentage'
  },
  
  CHARTS: {
    INCOME_VS_EXPENSE: 'Monthly trend comparison',
    CATEGORY_BREAKDOWN: 'Expense distribution by category',
    INVESTMENT_PERFORMANCE: 'Portfolio value over time',
    BUDGET_PROGRESS: 'Category-wise budget usage'
  },
  
  ALERTS: {
    BUDGET_WARNINGS: 'Categories exceeding thresholds',
    UPCOMING_DEBT_PAYMENTS: 'Payment reminders',
    INVESTMENT_ALERTS: 'Significant portfolio changes'
  }
};
```

### Layer 3: Application Layer

#### Core Modules

##### Main Controller (Main.gs)
```javascript
// Central application controller
class ApplicationController {
  constructor() {
    this.version = APP_CONFIG.VERSION;
    this.modules = {
      dataProcessor: new DataProcessor(),
      budgetManager: new BudgetManager(),
      dashboardManager: new DashboardManager()
    };
  }
  
  // Request routing
  handleRequest(request) {
    const { module, action, data } = request;
    return this.modules[module][action](data);
  }
  
  // System initialization
  initialize() {
    this.setupMenu();
    this.validateSheets();
    this.checkVersion();
  }
}
```

##### Data Processing Engine (DataProcessor.gs)
```javascript
// Transaction processing pipeline
class DataProcessor {
  
  // Generic transaction processor
  processTransaction(type, data) {
    try {
      // 1. Validate input data
      this.validateInput(type, data);
      
      // 2. Business logic processing
      const processedData = this.applyBusinessRules(type, data);
      
      // 3. Data persistence
      const result = this.persistData(type, processedData);
      
      // 4. Post-processing
      this.triggerPostProcessing(type, result);
      
      return { success: true, data: result };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
  
  // Specific processors
  addIncome(data) { return this.processTransaction('INCOME', data); }
  addExpense(data) { return this.processTransaction('EXPENSE', data); }
  addInvestment(data) { return this.processTransaction('INVESTMENT', data); }
}
```

##### Budget Management (BudgetManager.gs)
```javascript
// Budget monitoring and alerting
class BudgetManager {
  
  // Budget validation pipeline
  validateExpense(category, amount) {
    const budget = this.getCategoryBudget(category);
    const spent = this.getCategorySpent(category);
    const utilization = (spent + amount) / budget;
    
    return {
      approved: utilization <= 1.0,
      warning: utilization >= 0.7,
      critical: utilization >= 0.9,
      utilization: utilization
    };
  }
  
  // Alert system
  generateBudgetAlerts() {
    const categories = this.getAllCategories();
    const alerts = [];
    
    categories.forEach(category => {
      const status = this.getCategoryStatus(category);
      if (status.needsAlert) {
        alerts.push(this.createAlert(category, status));
      }
    });
    
    return alerts;
  }
}
```

### Layer 4: Integration Layer

#### External API Integration
```javascript
// API client pattern
class ExternalAPIManager {
  
  // Stock price API (TCBS)
  async getStockPrice(symbol) {
    const url = `https://apipubaws.tcbs.com.vn/stock-insight/v1/stock/${symbol}/overview`;
    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());
      return this.normalizeStockData(data);
    } catch (error) {
      Logger.log(`Stock API error for ${symbol}: ${error.message}`);
      return null;
    }
  }
  
  // Gold price API (Doji)
  async getGoldPrice() {
    const url = 'https://api.doji.vn/api/gold-price/current';
    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());
      return this.normalizeGoldData(data);
    } catch (error) {
      Logger.log(`Gold API error: ${error.message}`);
      return this.getFallbackGoldPrice();
    }
  }
}
```

#### Trigger System
```javascript
// Event-driven automation
class TriggerManager {
  
  // Setup automated triggers
  setupTriggers() {
    this.createDailyTrigger();
    this.createWeeklyTrigger();
    this.createMonthlyTrigger();
  }
  
  createDailyTrigger() {
    ScriptApp.newTrigger('dailyMaintenance')
      .timeBased()
      .everyDays(1)
      .atHour(6) // 6 AM
      .create();
  }
  
  // Daily maintenance tasks
  dailyMaintenance() {
    this.updateMarketPrices();
    this.checkBudgetAlerts();
    this.cleanupTempData();
    this.generateDailyReport();
  }
}
```

### Layer 5: Data Layer

#### Sheet Architecture
```javascript
// Data schema definition
const SHEET_SCHEMAS = {
  INCOME: {
    columns: ['STT', 'Ngày', 'Loại thu', 'Mô tả', 'Số tiền', 'Ghi chú'],
    types: ['number', 'date', 'string', 'string', 'currency', 'string'],
    required: ['Ngày', 'Loại thu', 'Số tiền']
  },
  
  EXPENSE: {
    columns: ['STT', 'Ngày', 'Danh mục', 'Mô tả', 'Số tiền', 'Phương thức'],
    types: ['number', 'date', 'string', 'string', 'currency', 'string'],
    required: ['Ngày', 'Danh mục', 'Số tiền'],
    budgetTracking: true
  },
  
  STOCK: {
    columns: ['STT', 'Ngày', 'Loại GD', 'Mã CK', 'Số lượng', 'Giá', 'Phí'],
    types: ['number', 'date', 'string', 'string', 'number', 'currency', 'currency'],
    required: ['Ngày', 'Loại GD', 'Mã CK', 'Số lượng', 'Giá'],
    calculations: ['Tổng tiền', 'P&L']
  }
};
```

#### Data Access Layer
```javascript
// Data access abstraction
class DataAccessLayer {
  
  constructor() {
    this.sheets = this.initializeSheets();
    this.cache = new Map();
  }
  
  // Generic CRUD operations
  create(sheetName, data) {
    const sheet = this.getSheet(sheetName);
    const row = this.findEmptyRow(sheet);
    const values = this.prepareRowData(sheetName, data);
    
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    this.invalidateCache(sheetName);
    
    return { row, data: values };
  }
  
  read(sheetName, filters = {}) {
    const cacheKey = `${sheetName}_${JSON.stringify(filters)}`;
    
    if (this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey);
    }
    
    const data = this.querySheet(sheetName, filters);
    this.cache.set(cacheKey, data);
    
    return data;
  }
  
  update(sheetName, rowIndex, data) {
    const sheet = this.getSheet(sheetName);
    const values = this.prepareRowData(sheetName, data);
    
    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
    this.invalidateCache(sheetName);
  }
  
  delete(sheetName, rowIndex) {
    const sheet = this.getSheet(sheetName);
    sheet.deleteRow(rowIndex);
    this.invalidateCache(sheetName);
  }
}
```

## 🔄 Data Flow Architecture

### Transaction Processing Flow

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   User Input    │───▶│   HTML Form      │───▶│   Validation    │
│   (Manual)      │    │   (Client-side)  │    │   (Client)      │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                                                         │
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│  Sheet Update   │◀───│  Data Processing │◀───│   Server Call   │
│  (Persistence)  │    │  (Business Logic)│    │  (google.script)│
└─────────────────┘    └──────────────────┘    └─────────────────┘
          │                       │
          ▼                       ▼
┌─────────────────┐    ┌──────────────────┐
│ Dashboard Refresh│    │  Budget Check    │
│ (Analytics)     │    │  (Alert System)  │
└─────────────────┘    └──────────────────┘
```

### Budget Monitoring Flow

```
┌─────────────────┐
│ Expense Entry   │
└─────────┬───────┘
          ▼
┌─────────────────┐
│ Budget Checker  │
├─────────────────┤
│ • Get category  │
│ • Check limit   │
│ • Calculate %   │
└─────────┬───────┘
          ▼
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   < 70%         │    │   70% - 90%      │    │     > 90%       │
│   Green         │    │   Yellow         │    │     Red         │
│   Continue      │    │   Show Warning   │    │   Show Alert    │
└─────────────────┘    └──────────────────┘    └─────────────────┘
          │                       │                       │
          └───────────────────────┼───────────────────────┘
                                  ▼
                    ┌──────────────────┐
                    │ Update Dashboard │
                    │ Apply Formatting │
                    │ Log Transaction  │
                    └──────────────────┘
```

### Dashboard Update Flow

```
┌─────────────────┐
│   Data Change   │
│   (Any Sheet)   │
└─────────┬───────┘
          ▼
┌─────────────────┐
│   Auto Trigger  │
│   (onChange)    │
└─────────┬───────┘
          ▼
┌─────────────────┐    ┌──────────────────┐
│ Aggregate Data  │───▶│  Calculate KPIs  │
│ • Sum income    │    │  • Cash flow     │
│ • Sum expense   │    │  • Budget usage  │
│ • Sum investments│    │  • ROI           │
└─────────────────┘    └─────────┬────────┘
                                ▼
┌─────────────────┐    ┌──────────────────┐
│ Update Charts   │◀───│  Update Summary  │
│ • Trends        │    │  • Cards         │
│ • Categories    │    │  • Alerts        │
│ • Performance   │    │  • Status        │
└─────────────────┘    └──────────────────┘
```

## ⚡ Performance Architecture

### Optimization Strategies

#### 1. Formula-Based Calculations
```javascript
// Use spreadsheet formulas instead of script calculations
const FORMULA_PATTERNS = {
  SUM_INCOME: '=SUMIF(THU!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),THU!E:E)',
  SUM_EXPENSE: '=SUMIF(CHI!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),CHI!E:E)',
  CASH_FLOW: '=Income_Total-Expense_Total-Investment_Total',
  BUDGET_USAGE: '=Expense_Category/Budget_Category'
};
```

#### 2. Batch Operations
```javascript
// Batch data operations for performance
class BatchProcessor {
  
  constructor(batchSize = 100) {
    this.batchSize = batchSize;
    this.pendingOperations = [];
  }
  
  addOperation(operation) {
    this.pendingOperations.push(operation);
    
    if (this.pendingOperations.length >= this.batchSize) {
      this.processBatch();
    }
  }
  
  processBatch() {
    const batch = this.pendingOperations.splice(0, this.batchSize);
    
    // Execute all operations in a single API call
    const values = batch.map(op => op.values);
    const range = `A${batch[0].row}:G${batch[0].row + batch.length - 1}`;
    
    this.sheet.getRange(range).setValues(values);
  }
}
```

#### 3. Caching Strategy
```javascript
// Multi-layer caching system
class CacheManager {
  
  constructor() {
    this.scriptCache = CacheService.getScriptCache();
    this.memoryCache = new Map();
    this.propertyCache = PropertiesService.getScriptProperties();
  }
  
  get(key, level = 'memory') {
    switch(level) {
      case 'memory':
        return this.memoryCache.get(key);
      case 'script':
        return JSON.parse(this.scriptCache.get(key) || 'null');
      case 'property':
        return JSON.parse(this.propertyCache.getProperty(key) || 'null');
    }
  }
  
  set(key, value, ttl = 600, level = 'memory') {
    const serialized = JSON.stringify(value);
    
    switch(level) {
      case 'memory':
        this.memoryCache.set(key, value);
        break;
      case 'script':
        this.scriptCache.put(key, serialized, ttl);
        break;
      case 'property':
        this.propertyCache.setProperty(key, serialized);
        break;
    }
  }
}
```

### Scalability Considerations

#### Data Volume Limits
```javascript
// Monitor data growth and implement archiving
const SCALABILITY_LIMITS = {
  MAX_ROWS_PER_SHEET: 40000,        // Google Sheets limit
  MAX_EXECUTION_TIME: 360000,       // 6 minutes (Apps Script limit)
  MAX_DAILY_TRIGGERS: 20,           // Apps Script quota
  MAX_SCRIPT_RUNTIME: 6 * 60 * 1000, // 6 minutes in ms
  
  // Performance thresholds
  PERFORMANCE_TARGETS: {
    FORM_SUBMISSION: 3000,          // 3 seconds max
    DASHBOARD_REFRESH: 5000,        // 5 seconds max
    BULK_IMPORT: 30000              // 30 seconds max
  }
};
```

#### Auto-archiving System
```javascript
// Automatic data archiving for performance
function setupAutoArchiving() {
  // Archive data older than 2 years
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 2);
  
  const sheets = ['THU', 'CHI', 'CHỨNG KHOÁN', 'VÀNG', 'CRYPTO'];
  
  sheets.forEach(sheetName => {
    archiveOldData(sheetName, cutoffDate);
  });
}

function archiveOldData(sheetName, cutoffDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  // Split data into current and archive
  const currentData = data.filter(row => new Date(row[1]) >= cutoffDate);
  const archiveData = data.filter(row => new Date(row[1]) < cutoffDate);
  
  if (archiveData.length > 0) {
    // Create archive sheet
    const archiveName = `${sheetName}_ARCHIVE_${cutoffDate.getFullYear()}`;
    createArchiveSheet(archiveName, archiveData);
    
    // Update current sheet with recent data only
    sheet.clear();
    sheet.getRange(1, 1, currentData.length, currentData[0].length)
      .setValues(currentData);
  }
}
```

## 🔐 Security Architecture

### Authentication & Authorization
```javascript
// User access control
class SecurityManager {
  
  validateUser() {
    const user = Session.getActiveUser();
    
    if (!user.getEmail()) {
      throw new Error('User not authenticated');
    }
    
    // Check if user has access to this spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const editors = spreadsheet.getEditors();
    
    const hasAccess = editors.some(editor => 
      editor.getEmail() === user.getEmail()
    );
    
    if (!hasAccess) {
      throw new Error('User not authorized');
    }
    
    return user;
  }
  
  validatePermissions(requiredPermission) {
    const user = this.validateUser();
    const userPermissions = this.getUserPermissions(user.getEmail());
    
    if (!userPermissions.includes(requiredPermission)) {
      throw new Error(`Permission denied: ${requiredPermission}`);
    }
  }
}
```

### Data Protection
```javascript
// Input sanitization and validation
class InputSanitizer {
  
  static sanitizeInput(input, type) {
    if (input === null || input === undefined) {
      return input;
    }
    
    switch(type) {
      case 'string':
        return String(input)
          .replace(/[<>]/g, '')
          .replace(/javascript:/gi, '')
          .trim();
          
      case 'number':
        const num = parseFloat(input);
        return isNaN(num) ? 0 : num;
        
      case 'date':
        const date = new Date(input);
        return isNaN(date.getTime()) ? new Date() : date;
        
      default:
        return input;
    }
  }
  
  static validateFinancialData(data) {
    if (data.amount <= 0) {
      throw new Error('Amount must be positive');
    }
    
    if (data.amount > 1e12) {
      throw new Error('Amount exceeds maximum limit');
    }
    
    if (!data.category || data.category.trim() === '') {
      throw new Error('Category is required');
    }
    
    const transactionDate = new Date(data.date);
    const maxFutureDate = new Date();
    maxFutureDate.setDate(maxFutureDate.getDate() + 30);
    
    if (transactionDate > maxFutureDate) {
      throw new Error('Transaction date cannot be more than 30 days in the future');
    }
  }
}
```

## 🔧 Error Handling Architecture

### Global Error Handler
```javascript
// Centralized error handling
class ErrorManager {
  
  static handleError(error, context = {}) {
    const errorInfo = {
      message: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString(),
      user: Session.getActiveUser().getEmail(),
      context: context,
      version: APP_CONFIG.VERSION
    };
    
    // Log error
    Logger.log('Error: ' + JSON.stringify(errorInfo));
    
    // User-friendly error message
    const userMessage = this.getUserFriendlyMessage(error);
    
    // Optional: Send to external monitoring
    if (APP_CONFIG.ERROR_REPORTING_ENABLED) {
      this.reportToExternalService(errorInfo);
    }
    
    return {
      success: false,
      message: userMessage,
      errorId: this.generateErrorId()
    };
  }
  
  static getUserFriendlyMessage(error) {
    const errorMap = {
      'Permission denied': 'Bạn không có quyền thực hiện thao tác này',
      'Invalid input': 'Dữ liệu nhập vào không hợp lệ',
      'Network error': 'Lỗi kết nối mạng, vui lòng thử lại',
      'Quota exceeded': 'Đã vượt quá giới hạn sử dụng, vui lòng thử lại sau'
    };
    
    for (const [key, message] of Object.entries(errorMap)) {
      if (error.message.includes(key)) {
        return message;
      }
    }
    
    return 'Có lỗi xảy ra, vui lòng thử lại hoặc liên hệ hỗ trợ';
  }
}
```

This system architecture provides a robust, scalable, and maintainable foundation for the HODLVN-Family-Finance application, ensuring reliable performance while maintaining simplicity for end users.