# Deployment Guide

## Overview

This guide covers the deployment process for HODLVN-Family-Finance, including installation, configuration, updates, and maintenance procedures. The system is deployed on Google Apps Script and Google Sheets infrastructure.

## 🚀 Deployment Architecture

### Platform Components

```
┌─────────────────────────────────────────────────────┐
│                 USER ACCESS LAYER                    │
├─────────────────────────────────────────────────────┤
│ Google Sheets UI (Forms & Dashboard)                 │
│ - Web browser access                                 │
│ - Mobile app access                                  │
│ - Collaborative sharing                              │
└─────────────────────────────────────────────────────┘
                            │
┌─────────────────────────────────────────────────────┐
│               APPLICATION LAYER                      │
├─────────────────────────────────────────────────────┤
│ Google Apps Script Runtime                           │
│ - Server-side JavaScript execution                  │
│ - HTML service for forms                            │
│ - Trigger-based automation                          │
└─────────────────────────────────────────────────────┘
                            │
┌─────────────────────────────────────────────────────┐
│                 DATA LAYER                          │
├─────────────────────────────────────────────────────┤
│ Google Sheets (Spreadsheet Database)                │
│ - 13 specialized sheets                             │
│ - Real-time collaboration                           │
│ - Formula-based calculations                        │
│ - Google Drive storage                              │
└─────────────────────────────────────────────────────┘
```

## 📋 Prerequisites

### Requirements

#### For End Users
- Google Account (Gmail)
- Modern web browser (Chrome, Firefox, Edge, Safari)
- Internet connection
- Basic Google Sheets familiarity

#### For Developers/Administrators
- Google Account with Apps Script access
- GitHub account (for source code)
- Text editor or IDE
- Understanding of JavaScript and Google Apps Script

#### Permissions Needed
```
Google Apps Script permissions:
- View and manage spreadsheets in Google Drive
- Display and run third-party web content
- Connect to external services
- Allow script to send email (for notifications)
```

## 🔧 Installation Methods

### Method 1: Template Deployment (Recommended for End Users)

#### Step 1: Copy Template Spreadsheet
```bash
# Template URL (when available)
https://docs.google.com/spreadsheets/d/[TEMPLATE_ID]/copy

# Manual steps:
1. Open template link
2. File → Make a copy
3. Name: "HODLVN-Family-Finance - [Your Name]"
4. Choose Google Drive location
5. Click "Make a copy"
```

#### Step 2: Enable Apps Script
```bash
1. Open copied spreadsheet
2. Extensions → Apps Script
3. Script editor opens (already configured)
4. First-time permissions dialog appears
5. Click "Review permissions"
6. Allow required access
```

#### Step 3: Run Initial Setup
```bash
1. In Apps Script editor, select "setup" function
2. Click "Run" button (▶️)
3. Grant permissions when prompted
4. Wait for completion (green checkmark)
5. Return to spreadsheet
```

#### Step 4: Launch Setup Wizard
```bash
1. Refresh spreadsheet (F5)
2. "HODLVN Finance" menu appears
3. Click "🚀 Setup Wizard"
4. Complete user information form
5. Click "Khởi tạo hệ thống"
6. System creates all required sheets
```

### Method 2: Source Code Deployment (For Developers)

#### Step 1: Clone Repository
```bash
git clone https://github.com/tohaitrieu/HODLVN-Family-Finance.git
cd HODLVN-Family-Finance
```

#### Step 2: Create New Google Sheets Project
```bash
1. Create new Google Spreadsheet
2. Extensions → Apps Script
3. Delete default Code.gs content
4. Create new files matching project structure
```

#### Step 3: Deploy Source Files
```bash
# Upload .gs files to Apps Script editor
# Main.gs - Core menu system
# DataProcessor.gs - Transaction handling
# DashboardManager.gs - Analytics
# SheetInitializer.gs - Setup system
# Utils.gs - Utilities
# [Additional modules as needed]

# Upload .html files as HTML files
# SetupWizard.html
# IncomeForm.html
# ExpenseForm.html
# [Additional forms as needed]
```

#### Step 4: Configure Project
```bash
1. Set project name: "HODLVN-Family-Finance"
2. Save all files (Ctrl+S)
3. Run initial authorization
4. Test basic functionality
```

## ⚙️ Configuration

### System Configuration

#### App Configuration (Main.gs)
```javascript
const APP_CONFIG = {
  VERSION: '3.5.8',
  APP_NAME: '💰 Quản lý Tài chính',
  
  // Sheet names (modify if needed)
  SHEETS: {
    INCOME: 'THU',
    EXPENSE: 'CHI',
    DEBT: 'QUẢN LÝ NỢ',
    // ... other sheets
  },
  
  // Customization options
  COLORS: {
    PRIMARY: '#4472C4',
    SUCCESS: '#70AD47',
    ERROR: '#E74C3C'
  },
  
  // Feature flags
  FEATURES: {
    ENABLE_STOCK_API: true,
    ENABLE_GOLD_API: true,
    ENABLE_EMAIL_ALERTS: false
  }
};
```

#### Budget Thresholds
```javascript
const BUDGET_CONFIG = {
  WARNING_THRESHOLD: 0.7,  // 70% - Yellow warning
  DANGER_THRESHOLD: 0.9,   // 90% - Red alert
  
  // Default categories
  DEFAULT_CATEGORIES: [
    'Ăn uống',
    'Đi lại', 
    'Học tập',
    'Giải trí',
    'Y tế'
  ]
};
```

### User Customization

#### Category Management
```javascript
// Add custom expense categories
function addCustomCategory(categoryName) {
  const categories = getExpenseCategories();
  categories.push(categoryName);
  updateCategoriesInSheet(categories);
}

// Modify budget limits
function setBudgetLimit(category, monthlyLimit) {
  const budgetSheet = getSheet('BUDGET');
  updateCategoryBudget(budgetSheet, category, monthlyLimit);
}
```

#### Notification Settings
```javascript
// Configure email notifications
const NOTIFICATION_CONFIG = {
  ENABLE_EMAIL: false,              // Set to true to enable
  EMAIL_ADDRESS: 'user@email.com',  // User's email
  BUDGET_ALERTS: true,              // Budget warning emails
  DEBT_REMINDERS: true              // Debt payment reminders
};
```

## 🔄 Update Procedures

### Version Updates

#### Automatic Update Detection
```javascript
// System checks for updates on menu load
function checkForUpdates() {
  const currentVersion = APP_CONFIG.VERSION;
  const latestVersion = getLatestVersionFromGitHub();
  
  if (compareVersions(latestVersion, currentVersion) > 0) {
    showUpdateNotification();
  }
}
```

#### Manual Update Process
```bash
1. Backup current data (Extensions → Apps Script → Manage deployments)
2. Download latest source code from GitHub
3. Replace .gs files in Apps Script editor
4. Update .html files
5. Run migration function if needed
6. Test functionality
7. Update version number in APP_CONFIG
```

#### Data Migration
```javascript
// Version migration example
function migrateToVersion358() {
  try {
    // Add new columns to existing sheets
    addNewBudgetColumns();
    
    // Update formulas
    updateDashboardFormulas();
    
    // Migrate user preferences
    migrateUserSettings();
    
    Logger.log('✅ Migration to v3.5.8 completed');
    return { success: true };
  } catch (error) {
    Logger.log('❌ Migration failed: ' + error.message);
    return { success: false, error: error.message };
  }
}
```

### Content Updates

#### Adding New Transaction Categories
```bash
1. Open CategoryManager.gs
2. Add category to appropriate array
3. Update form dropdowns in HTML files
4. Test category selection
5. Update documentation
```

#### Modifying Sheet Structure
```bash
1. Backup existing data
2. Update SheetInitializer.gs
3. Modify DataProcessor.gs for new fields
4. Update form HTML
5. Run migration script
6. Test data integrity
```

## 🔒 Security Deployment

### Access Control

#### User Permissions
```bash
# Sheet sharing settings
1. Share spreadsheet with specific users
2. Set permissions (Editor/Viewer)
3. Disable link sharing if needed
4. Set expiration dates for temporary access
```

#### Apps Script Security
```javascript
// Validate user access
function validateUserAccess() {
  const user = Session.getActiveUser().getEmail();
  const allowedUsers = PropertiesService
    .getScriptProperties()
    .getProperty('ALLOWED_USERS')
    .split(',');
    
  if (!allowedUsers.includes(user)) {
    throw new Error('Unauthorized access');
  }
}
```

### Data Protection

#### Backup Strategy
```bash
# Automated backup (weekly trigger)
function scheduleBackups() {
  ScriptApp.newTrigger('createBackup')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .create();
}

function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = `HODLVN-Backup-${Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd')}`;
  const backup = ss.copy(backupName);
  
  // Move to backup folder
  const backupFolder = DriveApp.getFoldersByName('HODLVN-Backups').next();
  DriveApp.getFileById(backup.getId()).moveTo(backupFolder);
}
```

#### Data Validation
```javascript
// Input sanitization
function sanitizeInput(input) {
  if (typeof input !== 'string') return input;
  
  return input
    .replace(/[<>]/g, '')      // Remove angle brackets
    .replace(/javascript:/gi, '') // Remove JavaScript protocol
    .trim();
}

// Financial data validation
function validateFinancialData(amount, category, date) {
  if (amount <= 0) throw new Error('Amount must be positive');
  if (!category) throw new Error('Category is required');
  if (new Date(date) > new Date()) throw new Error('Date cannot be in future');
}
```

## 📊 Monitoring & Maintenance

### Performance Monitoring

#### Execution Time Tracking
```javascript
function trackPerformance(functionName, startTime) {
  const executionTime = Date.now() - startTime;
  
  Logger.log(`📊 ${functionName} executed in ${executionTime}ms`);
  
  // Log to sheet for analysis
  const perfSheet = getSheet('PERFORMANCE');
  if (perfSheet) {
    perfSheet.appendRow([
      new Date(),
      functionName,
      executionTime,
      Session.getActiveUser().getEmail()
    ]);
  }
}
```

#### Error Logging
```javascript
function logError(error, context = '') {
  const errorLog = {
    timestamp: new Date().toISOString(),
    error: error.toString(),
    context: context,
    user: Session.getActiveUser().getEmail(),
    version: APP_CONFIG.VERSION
  };
  
  Logger.log('❌ Error: ' + JSON.stringify(errorLog));
  
  // Optional: Send to external monitoring service
  // sendErrorToMonitoring(errorLog);
}
```

### Health Checks

#### System Health Dashboard
```javascript
function getSystemHealth() {
  const health = {
    version: APP_CONFIG.VERSION,
    lastUpdate: new Date(),
    sheetsCount: getSheetCount(),
    dataIntegrity: checkDataIntegrity(),
    performance: getPerformanceMetrics(),
    errors: getRecentErrors()
  };
  
  return health;
}

// Run daily health check
function scheduleDailyHealthCheck() {
  ScriptApp.newTrigger('runHealthCheck')
    .timeBased()
    .everyDays(1)
    .atHour(6) // 6 AM
    .create();
}
```

### Maintenance Tasks

#### Weekly Maintenance Checklist
```bash
□ Check system health dashboard
□ Review error logs
□ Verify backup completion
□ Monitor performance metrics
□ Update external API keys (if needed)
□ Check user feedback
□ Review security access logs
```

#### Monthly Tasks
```bash
□ Full data integrity check
□ Performance optimization review
□ User access audit
□ Version update check
□ Documentation updates
□ Feature usage analytics
```

#### Quarterly Tasks
```bash
□ Security review
□ Disaster recovery testing
□ User satisfaction survey
□ Infrastructure cost review
□ Feature roadmap planning
□ Training material updates
```

## 🔧 Troubleshooting

### Common Deployment Issues

#### Permission Errors
```bash
Issue: "Script does not have permission to perform that action"
Solution:
1. Run authorization again: Run → Review permissions
2. Clear browser cache
3. Try incognito mode
4. Check Google account admin settings
```

#### Script Timeout Errors
```bash
Issue: "Script execution exceeded timeout"
Solution:
1. Optimize slow functions
2. Implement batch processing
3. Use triggers for long operations
4. Split large operations into smaller chunks
```

#### Missing Menu Items
```bash
Issue: "HODLVN Finance menu doesn't appear"
Solution:
1. Refresh spreadsheet (F5)
2. Check onOpen() function execution
3. Run onOpen() manually in script editor
4. Check for JavaScript errors in console
```

### Diagnostic Tools

#### Debug Mode
```javascript
const DEBUG_MODE = true;

function debugLog(message, data = null) {
  if (!DEBUG_MODE) return;
  
  console.log(`🐛 DEBUG: ${message}`);
  if (data) console.log(data);
  
  Logger.log(`DEBUG: ${message} | Data: ${JSON.stringify(data)}`);
}
```

#### System Diagnostic Function
```javascript
function runDiagnostics() {
  const results = {
    version: APP_CONFIG.VERSION,
    permissions: checkPermissions(),
    sheets: validateSheetStructure(),
    triggers: getTriggerStatus(),
    performance: getPerformanceReport()
  };
  
  Logger.log('📋 Diagnostic Results:', JSON.stringify(results, null, 2));
  return results;
}
```

## 📚 Resources

### Support Documentation
- [Installation Guide](INSTALLATION.md) - Detailed setup instructions
- [User Guide](USER_GUIDE.md) - Feature documentation
- [Troubleshooting](TROUBLESHOOTING.md) - Common issues and solutions
- [Technical Documentation](TECHNICAL_DOCUMENTATION.md) - Developer reference

### External Resources
- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Google Sheets API Reference](https://developers.google.com/sheets/api)
- [GitHub Repository](https://github.com/tohaitrieu/HODLVN-Family-Finance)
- [Community Support](https://facebook.com/groups/hodl.vn)

### Contact Information
- **Technical Support**: GitHub Issues
- **General Questions**: Facebook Group
- **Security Issues**: Direct email to project maintainer

This deployment guide ensures reliable, secure, and maintainable installations of HODLVN-Family-Finance across different user environments and use cases.