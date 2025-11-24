# Code Standards & Conventions

## Overview

This document defines the coding standards and conventions for the HODLVN-Family-Finance project. All contributors must follow these guidelines to maintain code quality, consistency, and maintainability.

## 1. Google Apps Script Standards

### 1.1 File Organization

#### File Header Format
Every .gs file must start with a comprehensive header block:

```javascript
/**
 * ===============================================
 * FILENAME.GS - MODULE DESCRIPTION
 * ===============================================
 * 
 * Chức năng:
 * - Feature 1 description
 * - Feature 2 description
 * 
 * VERSION: x.x.x - Brief version description
 * CHANGELOG vx.x.x (DD/MM/YYYY):
 * ✅ NEW: New features added
 * ✅ FIX: Bug fixes implemented
 * ✅ UPDATE: Updates and improvements
 */
```

#### Section Dividers
Use clear section dividers to organize code:

```javascript
// ==================== CONFIGURATION ====================
// Configuration objects and constants

// ==================== CORE FUNCTIONS ====================
// Main functionality

// ==================== HELPER FUNCTIONS ====================
// Utility and helper functions
```

### 1.2 Naming Conventions

| Type | Convention | Example |
|------|-----------|---------|
| **Files** | PascalCase | `DataProcessor.gs`, `SheetInitializer.gs` |
| **Functions** | camelCase | `addIncome()`, `showExpenseForm()` |
| **Constants** | UPPER_SNAKE_CASE | `APP_CONFIG`, `LOAN_TYPES` |
| **Variables** | camelCase | `emptyRow`, `totalAmount` |
| **Sheet Names** | Vietnamese | `"Thu nhập"`, `"Chi tiêu"` |

### 1.3 Function Documentation

All functions must have JSDoc-style comments:

```javascript
/**
 * Brief description in Vietnamese
 * @param {string} sheetName - Tên sheet
 * @param {Object} data - Dữ liệu giao dịch
 * @param {number} data.amount - Số tiền
 * @param {string} data.date - Ngày giao dịch
 * @return {Object} {success: boolean, message: string}
 */
function processTransaction(sheetName, data) {
  // Implementation
}
```

### 1.4 Error Handling Pattern

Use consistent try-catch blocks with structured responses:

```javascript
function functionName(data) {
  try {
    // Validation
    if (!data.requiredField) {
      throw new Error('Missing required field');
    }
    
    // Main logic
    
    return {
      success: true,
      message: '✅ Thêm giao dịch thành công!'
    };
  } catch (error) {
    Logger.log('Error in functionName: ' + error.message);
    return {
      success: false,
      message: '❌ Lỗi: ' + error.message
    };
  }
}
```

### 1.5 Configuration Structure

Centralize configuration in `APP_CONFIG`:

```javascript
const APP_CONFIG = {
  VERSION: '3.5.8',
  APP_NAME: '💰 Quản lý Tài chính',
  SHEETS: {
    INCOME: 'THU',
    EXPENSE: 'CHI',
    DEBT: 'QUẢN LÝ NỢ',
    // etc...
  },
  COLORS: {
    PRIMARY: '#4472C4',
    SUCCESS: '#70AD47',
    ERROR: '#E74C3C'
  },
  FORMATS: {
    CURRENCY: '#,##0',
    PERCENTAGE: '0.00%'
  }
};
```

## 2. HTML/CSS Standards

### 2.1 HTML Structure

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Form Title</title>
  <style>
    /* CSS styles here */
  </style>
</head>
<body>
  <div class="container">
    <h2>Form Title</h2>
    <form id="formName" onsubmit="handleSubmit(event)">
      <!-- Form content -->
    </form>
    <div id="message" class="message"></div>
  </div>
  <script>
    // JavaScript here
  </script>
</body>
</html>
```

### 2.2 CSS Conventions

#### Color Palette
```css
:root {
  --primary-color: #4472C4;
  --success-color: #70AD47;
  --error-color: #E74C3C;
  --border-color: #ddd;
  --background-color: #f0f0f0;
}
```

#### Form Styling
```css
.form-group {
  margin-bottom: 15px;
}

label {
  display: block;
  margin-bottom: 5px;
  font-weight: bold;
}

input, select, textarea {
  width: 100%;
  padding: 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  box-sizing: border-box;
}

.required {
  color: red;
}
```

### 2.3 Form Patterns

#### Required Field Markup
```html
<label for="fieldName">Field Label <span class="required">*</span></label>
<input type="text" id="fieldName" name="fieldName" required>
```

#### Form Submission Pattern
```javascript
function handleSubmit(event) {
  event.preventDefault();
  
  // Show loading state
  const submitButton = event.target.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = 'Đang xử lý...';
  
  // Collect form data
  const formData = {
    field1: document.getElementById('field1').value,
    field2: document.getElementById('field2').value
  };
  
  // Submit to server
  google.script.run
    .withSuccessHandler(handleSuccess)
    .withFailureHandler(handleError)
    .processFormData(formData);
}
```

## 3. Data Processing Standards

### 3.1 Sheet Operations

#### Finding Empty Row
Always use `findEmptyRow()` instead of `getLastRow()`:

```javascript
function findEmptyRow(sheet, startColumn = 'A') {
  const data = sheet.getRange(startColumn + ':' + startColumn).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === '') {
      return i + 1;
    }
  }
  return data.length + 1;
}
```

#### Auto-increment STT
```javascript
function getNextSTT(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  
  const lastSTT = sheet.getRange(lastRow, 1).getValue();
  return (parseInt(lastSTT) || 0) + 1;
}
```

### 3.2 Data Validation

#### Required Field Validation
```javascript
function validateRequiredFields(data, requiredFields) {
  for (const field of requiredFields) {
    if (!data[field]) {
      throw new Error(`${field} là bắt buộc`);
    }
  }
}
```

#### Numeric Validation
```javascript
function validateNumericField(value, fieldName) {
  const numValue = parseFloat(value);
  if (isNaN(numValue) || numValue <= 0) {
    throw new Error(`${fieldName} phải là số dương`);
  }
  return numValue;
}
```

## 4. Language & Localization

### 4.1 Language Usage

| Context | Language | Example |
|---------|----------|---------|
| **User Messages** | Vietnamese | `"✅ Thêm thu nhập thành công!"` |
| **Variable Names** | English | `totalAmount`, `transactionDate` |
| **Code Comments** | Vietnamese/English | Based on context |
| **Sheet Names** | Vietnamese | `"THU"`, `"CHI"`, `"QUẢN LÝ NỢ"` |
| **Log Messages** | Vietnamese (important) / English (debug) | Mixed usage |

### 4.2 Message Format

#### Success Messages
```javascript
return {
  success: true,
  message: '✅ ' + actionDescription + ' thành công!'
};
```

#### Error Messages
```javascript
return {
  success: false,
  message: '❌ Lỗi: ' + errorDescription
};
```

## 5. Best Practices

### 5.1 Code Quality

1. **DRY (Don't Repeat Yourself)**: Extract common logic into reusable functions
2. **Single Responsibility**: Each function should do one thing well
3. **Error Handling**: Always use try-catch for operations that might fail
4. **Logging**: Log important operations and errors for debugging

### 5.2 Performance

1. **Batch Operations**: Use `setValues()` instead of multiple `setValue()` calls
2. **Cache Sheet References**: Store sheet references instead of repeated `getSheetByName()` calls
3. **Minimize API Calls**: Group Google Sheets API calls when possible

### 5.3 Security

1. **Input Validation**: Always validate user input on the server side
2. **No Sensitive Data**: Never hardcode sensitive information
3. **Access Control**: Use Google's built-in authentication

### 5.4 Maintenance

1. **Version Control**: Update version numbers and changelog for significant changes
2. **Documentation**: Keep code comments and documentation up to date
3. **Testing**: Test all changes thoroughly before deployment

## 6. Git Commit Standards

### 6.1 Commit Message Format

```
<type>: <subject>

<body>

<footer>
```

#### Types
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting, etc.)
- `refactor`: Code refactoring
- `test`: Test additions or modifications
- `chore`: Build process or auxiliary tool changes

#### Example
```
feat: Add budget warning system

- Implement 3-level warning (70%, 90%, 100%)
- Add color coding for visual alerts
- Update dashboard to show warnings

Closes #123
```

## 7. Review Checklist

Before submitting code:

- [ ] Code follows naming conventions
- [ ] Functions have proper JSDoc comments
- [ ] Error handling is implemented
- [ ] User messages are in Vietnamese
- [ ] Code is tested with various inputs
- [ ] No hardcoded values (use configuration)
- [ ] Changelog is updated (if applicable)
- [ ] Code is formatted consistently

## 8. Resources

- [Google Apps Script Best Practices](https://developers.google.com/apps-script/guides/best-practices)
- [JSDoc Documentation](https://jsdoc.app/)
- [Google Style Guides](https://google.github.io/styleguide/)