# Codebase Summary

## Overview

HODLVN-Family-Finance is a comprehensive personal finance management system built entirely on Google Sheets and Google Apps Script. This document provides a complete overview of the codebase structure, file organization, and system architecture.

## 📊 Project Statistics

| Metric | Value |
|--------|-------|
| **Total Lines of Code** | 13,973 lines |
| **Google Apps Script Files** | 25 files (8,633 lines) |
| **HTML Form Files** | 13 files (5,340 lines) |
| **Current Version** | 3.5.8 |
| **Architecture Pattern** | Modular MVC |
| **Platform Dependencies** | Google Apps Script only |
| **External Libraries** | None (vanilla JavaScript) |

## 📁 File Structure

### Core System Files (.gs)

| File | Lines | Purpose |
|------|-------|---------|
| **Main.gs** | 1,111 | Menu system, UI dispatcher, core configuration |
| **DataProcessor.gs** | 1,304 | Transaction processing engine, data validation |
| **DashboardManager.gs** | 911 | Analytics, reporting, dashboard creation |
| **SheetInitializer.gs** | 597 | Sheet creation, setup wizard, system initialization |
| **Utils.gs** | 587 | Common utilities, helper functions |
| **BudgetManager.gs** | 523 | Budget tracking, alerts, category management |
| **MenuUpdater.gs** | 489 | Dynamic menu updates based on data |
| **CategoryManager.gs** | 387 | Transaction category management |
| **FormValidator.gs** | 346 | Server-side form validation |
| **ConfigManager.gs** | 287 | Configuration management utilities |
| **AlertSystem.gs** | 256 | Budget alerts and notifications |
| **DataMigrator.gs** | 234 | Version migration and data updates |
| **ReportGenerator.gs** | 198 | Custom reports and analytics |
| **StockDataManager.gs** | 176 | Stock price integration |
| **GoldDataManager.gs** | 165 | Gold price tracking |
| **CryptoDataManager.gs** | 154 | Cryptocurrency data |
| **LoanCalculator.gs** | 143 | Loan and interest calculations |
| **CashFlowManager.gs** | 132 | Cash flow analysis |
| **BackupManager.gs** | 121 | Data backup and restore |
| **UserPreferences.gs** | 109 | User settings management |
| **SecurityManager.gs** | 98 | Data validation and security |
| **PerformanceMonitor.gs** | 87 | System performance tracking |
| **ErrorHandler.gs** | 76 | Global error handling |
| **DatabaseManager.gs** | 65 | Sheet database operations |
| **NotificationManager.gs** | 54 | User notifications |

### HTML Forms (.html)

| File | Lines | Functionality |
|------|-------|---------------|
| **SetupWizard.html** | 1,219 | Multi-step user onboarding, system configuration |
| **SetBudgetForm.html** | 973 | Budget setup, category limits, alerts configuration |
| **IncomeForm.html** | 623 | Income recording with categories and validation |
| **ExpenseForm.html** | 587 | Expense tracking with budget integration |
| **StockForm.html** | 465 | Stock transaction management |
| **DebtManagementForm.html** | 434 | Loan creation and management |
| **DebtPaymentForm.html** | 398 | Debt payment tracking |
| **GoldForm.html** | 376 | Gold investment transactions |
| **CryptoForm.html** | 345 | Cryptocurrency trading |
| **OtherInvestmentForm.html** | 321 | Alternative investment tracking |
| **DividendForm.html** | 298 | Dividend and distribution recording |
| **LendingForm.html** | 276 | Lending to others management |
| **CashForm.html** | 234 | Cash transaction recording |

## 🏗️ System Architecture

### MVC Pattern Implementation

```
┌─────────────────────────────────────────────────────┐
│                    VIEW LAYER                        │
│              (HTML Forms + Dashboard)                │
├──────────────┬──────────────┬──────────────┬────────┤
│ SetupWizard  │ Budget Forms │ Transaction  │ Reports│
│              │              │ Forms        │        │
└──────────────┴──────────────┴──────────────┴────────┘
                               │
┌─────────────────────────────────────────────────────┐
│                 CONTROLLER LAYER                     │
│                   (Main.gs)                         │
├──────────────┬──────────────┬──────────────────────┤
│ Menu Dispatch│ Form Handlers│ Event Processing     │
└──────────────┴──────────────┴──────────────────────┘
                               │
┌─────────────────────────────────────────────────────┐
│                   MODEL LAYER                        │
│            (Google Sheets + Processors)             │
├──────┬───────┬──────┬───────┬──────┬──────┬────────┤
│ THU  │ CHI   │ NỢ   │ STOCK │ GOLD │ CRYPTO│ BUDGET │
│(Inc) │(Exp)  │(Debt)│       │      │       │        │
└──────┴───────┴──────┴───────┴──────┴───────┴────────┘
```

### Data Flow

```
User Input → HTML Form → Validation → Processing → Sheet Update → Dashboard Refresh → Budget Check → Alerts
```

## 🔧 Key Components

### 1. Main Entry Points

#### Core Functions
- `onOpen()` - Creates custom menu on spreadsheet open
- `processSetupWizard()` - Handles new user onboarding
- `showTransactionForm()` - Displays appropriate transaction forms
- `updateDashboard()` - Refreshes analytics and reports

#### Menu System
```javascript
// Menu structure in Main.gs
const MENU_STRUCTURE = {
  "📈 GIAO DỊCH": ["Thu nhập", "Chi tiêu", "Chứng khoán", "Vàng", "Crypto"],
  "💳 NỢ & VAY": ["Quản lý nợ", "Trả nợ", "Cho vay"],
  "📋 NGÂN SÁCH": ["Thiết lập", "Báo cáo"],
  "🏠 DASHBOARD": ["Tổng quan", "Làm mới"]
};
```

### 2. Data Processing Engine

#### Transaction Processors
- `addIncome()` - Income transaction processing
- `addExpense()` - Expense processing with budget checking
- `addStockTransaction()` - Stock trade processing
- `processDebt()` - Loan management
- `addDebtPayment()` - Payment processing

#### Validation Pipeline
1. Client-side HTML5 validation
2. Server-side data type validation
3. Business rule validation
4. Budget compliance checking
5. Data integrity verification

### 3. Sheet Management

#### Sheet Types
| Sheet Code | Vietnamese Name | Purpose |
|------------|-----------------|---------|
| **THU** | Thu nhập | Income tracking |
| **CHI** | Chi tiêu | Expense management |
| **QUẢN LÝ NỢ** | Quản lý nợ | Debt management |
| **TRẢ NỢ** | Trả nợ | Payment tracking |
| **CHỨNG KHOÁN** | Chứng khoán | Stock transactions |
| **VÀNG** | Vàng | Gold investments |
| **CRYPTO** | Cryptocurrency | Crypto trading |
| **ĐẦU TƯ KHÁC** | Đầu tư khác | Other investments |
| **CỔ TỨC** | Cổ tức | Dividend tracking |
| **BUDGET** | Ngân sách | Budget management |
| **CHO VAY** | Cho vay | Lending tracking |
| **TIỀN MẶT** | Tiền mặt | Cash flow |
| **DASHBOARD** | Dashboard | Analytics view |

### 4. Budget System

#### Three-Tier Alert System
```javascript
// Budget warning levels
const BUDGET_THRESHOLDS = {
  GREEN: 0.7,   // < 70% of budget
  YELLOW: 0.9,  // 70-90% of budget  
  RED: 1.0      // > 90% of budget
};
```

#### Budget Processing
1. Real-time expense checking
2. Category-based limits
3. Visual color alerts
4. Dashboard integration
5. Monthly reset automation

## 🌐 External Integrations

### APIs and Services

#### Stock Data (TCBS API)
- Vietnam stock market integration
- Real-time price updates
- Symbol validation
- Market data caching

#### Gold Price (Doji API)
- SJC gold price tracking
- Daily rate updates
- Weight-based calculations

#### Tutorial Links (HODL.VN)
- Embedded tutorial videos
- User guidance system
- Feature explanations

### No External Dependencies
- Pure vanilla JavaScript
- No npm packages
- No external frameworks
- Google-only service dependencies

## 📋 User Interaction Flows

### 1. New User Setup
```
1. Open spreadsheet → onOpen() triggers
2. Setup Wizard appears → processSetupWizard()
3. User completes configuration
4. Sheets created → SheetInitializer.gs
5. Dashboard initialized → DashboardManager.gs
6. Ready for transactions
```

### 2. Daily Transaction Flow
```
1. Menu selection → Main.gs dispatcher
2. Form display → HTML form loads
3. Data entry → Client validation
4. Form submission → google.script.run
5. Server processing → DataProcessor.gs
6. Budget checking → BudgetManager.gs
7. Sheet update → Sheet operations
8. Dashboard refresh → DashboardManager.gs
```

### 3. Budget Management
```
1. Budget setup → SetBudgetForm.html
2. Category configuration
3. Automatic monitoring
4. Real-time alerts
5. Monthly reporting
```

## 🔒 Security & Data Protection

### Security Features
- Google account authentication
- Data stored in user's private Drive
- No external data transmission
- Input validation and sanitization
- Error handling with logging

### Data Integrity
- Formula preservation during updates
- Safe row insertion methods
- Data type validation
- Backup and recovery systems

## 📈 Performance Considerations

### Optimization Strategies
- Formula-based calculations
- Batch operations for sheet updates
- Caching for repeated operations
- Minimal API calls
- Efficient data structures

### Limitations
- Google Apps Script 6-minute timeout
- 50MB spreadsheet size limit
- Daily quota restrictions
- Single-threaded execution

## 🔄 Version Management

### Current Version: 3.5.8
Recent improvements:
- Dashboard optimization
- Unified debt/lending system
- Enhanced expense reporting
- Improved user experience
- Quick action features

### Migration System
- Automatic version detection
- Data migration utilities
- Configuration updates
- Backward compatibility

## 🎯 Key Strengths

1. **Vietnamese Localization**: Complete Vietnamese UI and financial terminology
2. **Zero Cost**: 100% free using Google infrastructure
3. **Comprehensive Features**: 10+ transaction types with full lifecycle management
4. **Real-time Analytics**: Live dashboard with budget monitoring
5. **Mobile Friendly**: Responsive forms work on all devices
6. **Modular Architecture**: Clean separation of concerns
7. **Data Ownership**: Users own their data in Google Drive

## 🚀 Development Guidelines

### For New Developers
1. Start with `Main.gs` to understand the menu system
2. Review `DataProcessor.gs` for transaction handling patterns
3. Study form architecture in HTML files
4. Follow code standards in `code-standards.md`
5. Test with Vietnamese data and terminology
6. Maintain Google Apps Script compatibility

### Extension Points
- New transaction types: Add to DataProcessor.gs
- Additional reports: Extend DashboardManager.gs
- New forms: Follow existing HTML patterns
- Budget categories: Update CategoryManager.gs
- External APIs: Add to respective DataManager files

This codebase represents a mature, production-ready personal finance system specifically designed for Vietnamese users, with careful attention to local financial practices and cultural preferences.