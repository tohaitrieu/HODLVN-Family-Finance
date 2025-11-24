# Design Guidelines

## Overview

This document outlines the design principles, UI/UX guidelines, and visual standards for the HODLVN-Family-Finance project. These guidelines ensure a consistent, accessible, and user-friendly experience across all components.

## 🎯 Design Philosophy

### Core Principles

1. **Vietnamese-First Design**: All UI elements, terminology, and user flows are designed specifically for Vietnamese users
2. **Simplicity & Clarity**: Clean, uncluttered interfaces that prioritize essential information
3. **Accessibility**: Usable by people of all technical skill levels
4. **Mobile-Responsive**: Works seamlessly across desktop, tablet, and mobile devices
5. **Trust & Security**: Visual design conveys reliability and data security

### User-Centric Approach

- **Familiar Metaphors**: Use concepts familiar to Vietnamese financial culture
- **Error Prevention**: Proactive validation and clear feedback
- **Quick Actions**: Common tasks should be completed in 3 clicks or less
- **Visual Hierarchy**: Most important information is immediately visible

## 🎨 Visual Identity

### Color Palette

#### Primary Colors
```css
:root {
  --primary-blue: #4472C4;        /* Main brand color */
  --secondary-blue: #5B9BD5;      /* Lighter accent */
  --dark-blue: #2F5597;           /* Dark variant */
}
```

#### Status Colors
```css
:root {
  --success-green: #70AD47;       /* Success states, positive values */
  --warning-yellow: #FFC000;      /* Warning states, 70-90% budget */
  --error-red: #E74C3C;           /* Errors, critical alerts */
  --info-blue: #5DADE2;           /* Informational messages */
}
```

#### Budget Alert Colors
```css
:root {
  --budget-safe: #70AD47;         /* < 70% budget usage */
  --budget-warning: #F39C12;      /* 70-90% budget usage */
  --budget-danger: #E74C3C;       /* > 90% budget usage */
}
```

#### Neutral Colors
```css
:root {
  --background-light: #F8F9FA;    /* Page background */
  --background-white: #FFFFFF;    /* Card/form background */
  --border-light: #E0E0E0;        /* Light borders */
  --text-primary: #2C3E50;        /* Main text */
  --text-secondary: #7F8C8D;      /* Secondary text */
  --text-muted: #BDC3C7;          /* Placeholder text */
}
```

### Typography

#### Font Family
```css
body {
  font-family: 'Arial', 'Helvetica', sans-serif;
  /* Fallback: 'Segoe UI', 'Roboto', system fonts */
}
```

#### Font Hierarchy
```css
h1 { font-size: 24px; font-weight: 700; color: var(--primary-blue); }
h2 { font-size: 20px; font-weight: 600; color: var(--text-primary); }
h3 { font-size: 18px; font-weight: 600; color: var(--text-primary); }
h4 { font-size: 16px; font-weight: 600; color: var(--text-primary); }

.body-large { font-size: 16px; line-height: 1.5; }
.body-normal { font-size: 14px; line-height: 1.4; }
.body-small { font-size: 12px; line-height: 1.3; }
.caption { font-size: 11px; color: var(--text-secondary); }
```

#### Vietnamese Text Considerations
- Ensure proper rendering of Vietnamese diacritics
- Use clear, professional terminology
- Maintain consistency with Vietnamese financial terms

### Iconography

#### Emoji Usage
Use emojis consistently for visual recognition:

```css
/* Transaction Types */
📥 Thu nhập (Income)
📤 Chi tiêu (Expense)
💳 Nợ vay (Debt)
📈 Đầu tư (Investment)
💰 Ngân sách (Budget)
🏠 Dashboard (Overview)

/* Status Indicators */
✅ Success
❌ Error
⚠️ Warning
ℹ️ Information
🔄 Processing
```

#### Standard Icons
- Use simple, recognizable icons
- Maintain consistent style across all icons
- Ensure icons work well at small sizes

## 📱 Layout & Spacing

### Grid System

#### Container Widths
```css
.container {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0 20px;
}

/* Responsive breakpoints */
@media (max-width: 768px) {
  .container { padding: 0 16px; }
}
```

#### Spacing Scale
```css
:root {
  --space-xs: 4px;    /* Tight spacing */
  --space-sm: 8px;    /* Small spacing */
  --space-md: 16px;   /* Standard spacing */
  --space-lg: 24px;   /* Large spacing */
  --space-xl: 32px;   /* Extra large spacing */
}
```

### Component Spacing

#### Form Elements
```css
.form-group {
  margin-bottom: var(--space-md);
}

.form-row {
  margin-bottom: var(--space-sm);
}

label {
  margin-bottom: var(--space-xs);
  display: block;
}
```

#### Cards and Containers
```css
.card {
  padding: var(--space-lg);
  margin-bottom: var(--space-md);
  border-radius: 8px;
  border: 1px solid var(--border-light);
}
```

## 🔄 Component Design

### Forms

#### Form Structure
```html
<div class="form-container">
  <h2 class="form-title">📥 Thêm Thu nhập</h2>
  
  <form class="form">
    <div class="form-group">
      <label for="amount">Số tiền <span class="required">*</span></label>
      <input type="number" id="amount" class="form-input" required>
    </div>
    
    <div class="form-actions">
      <button type="submit" class="btn btn-primary">Thêm giao dịch</button>
      <button type="button" class="btn btn-secondary">Hủy</button>
    </div>
  </form>
</div>
```

#### Input Styles
```css
.form-input {
  width: 100%;
  padding: 12px;
  border: 1px solid var(--border-light);
  border-radius: 4px;
  font-size: 14px;
  transition: border-color 0.2s ease;
}

.form-input:focus {
  outline: none;
  border-color: var(--primary-blue);
  box-shadow: 0 0 0 2px rgba(68, 114, 196, 0.1);
}

.form-input:invalid {
  border-color: var(--error-red);
}
```

#### Button Styles
```css
.btn {
  padding: 12px 24px;
  border: none;
  border-radius: 4px;
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s ease;
  display: inline-block;
  text-align: center;
  text-decoration: none;
}

.btn-primary {
  background-color: var(--primary-blue);
  color: white;
}

.btn-primary:hover {
  background-color: var(--dark-blue);
}

.btn-secondary {
  background-color: var(--background-white);
  color: var(--text-primary);
  border: 1px solid var(--border-light);
}
```

### Tables & Data Display

#### Table Styling
```css
.data-table {
  width: 100%;
  border-collapse: collapse;
  margin-bottom: var(--space-lg);
}

.data-table th {
  background-color: var(--background-light);
  font-weight: 600;
  text-align: left;
  padding: 12px;
  border-bottom: 2px solid var(--border-light);
}

.data-table td {
  padding: 12px;
  border-bottom: 1px solid var(--border-light);
}
```

#### Number Formatting
```css
.currency {
  text-align: right;
  font-weight: 600;
  font-family: monospace;
}

.positive { color: var(--success-green); }
.negative { color: var(--error-red); }
.neutral { color: var(--text-primary); }
```

### Dashboard Components

#### Metric Cards
```css
.metric-card {
  background: white;
  border-radius: 8px;
  padding: var(--space-lg);
  border-left: 4px solid var(--primary-blue);
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.metric-value {
  font-size: 28px;
  font-weight: 700;
  color: var(--text-primary);
  margin-bottom: var(--space-xs);
}

.metric-label {
  font-size: 14px;
  color: var(--text-secondary);
  text-transform: uppercase;
  letter-spacing: 0.5px;
}
```

#### Chart Containers
```css
.chart-container {
  background: white;
  border-radius: 8px;
  padding: var(--space-lg);
  margin-bottom: var(--space-md);
  min-height: 300px;
}

.chart-title {
  font-size: 18px;
  font-weight: 600;
  margin-bottom: var(--space-md);
  color: var(--text-primary);
}
```

## 📋 User Experience Guidelines

### Navigation

#### Menu Hierarchy
```
Main Menu
├── 📈 GIAO DỊCH
│   ├── Thu nhập
│   ├── Chi tiêu
│   ├── Chứng khoán
│   ├── Vàng
│   └── Cryptocurrency
├── 💳 NỢ & VAY
│   ├── Quản lý nợ
│   ├── Trả nợ
│   └── Cho vay
├── 📋 NGÂN SÁCH
│   ├── Thiết lập ngân sách
│   └── Báo cáo ngân sách
└── 🏠 DASHBOARD
    ├── Tổng quan
    └── Làm mới dữ liệu
```

#### Breadcrumb Pattern
```html
<div class="breadcrumb">
  <span>🏠 Dashboard</span>
  <span class="separator">›</span>
  <span>📈 Giao dịch</span>
  <span class="separator">›</span>
  <span class="current">Thu nhập</span>
</div>
```

### Form UX Patterns

#### Progressive Disclosure
- Show basic fields first
- Reveal advanced options only when needed
- Use accordion or toggle patterns for optional fields

#### Validation Feedback
```html
<!-- Real-time validation -->
<div class="form-group">
  <label for="amount">Số tiền</label>
  <input type="number" id="amount" class="form-input" required>
  <div class="field-error" id="amount-error" style="display: none;">
    Vui lòng nhập số tiền hợp lệ
  </div>
</div>
```

#### Success/Error States
```css
.message-success {
  background-color: var(--success-green);
  color: white;
  padding: 12px;
  border-radius: 4px;
  margin-bottom: var(--space-md);
}

.message-error {
  background-color: var(--error-red);
  color: white;
  padding: 12px;
  border-radius: 4px;
  margin-bottom: var(--space-md);
}
```

### Loading States

#### Button Loading
```css
.btn-loading {
  position: relative;
  pointer-events: none;
  opacity: 0.7;
}

.btn-loading::after {
  content: "";
  position: absolute;
  width: 16px;
  height: 16px;
  margin: auto;
  border: 2px solid transparent;
  border-top-color: currentColor;
  border-radius: 50%;
  animation: button-loading-spinner 1s ease infinite;
}
```

## 📱 Responsive Design

### Mobile-First Approach

#### Breakpoints
```css
/* Mobile first */
.component { /* Mobile styles */ }

/* Tablet */
@media (min-width: 768px) {
  .component { /* Tablet styles */ }
}

/* Desktop */
@media (min-width: 1024px) {
  .component { /* Desktop styles */ }
}
```

#### Mobile Optimization
- Touch-friendly button sizes (minimum 44px)
- Simplified navigation for small screens
- Larger text for readability
- Optimized form layouts for portrait orientation

### Touch Interactions
```css
.touch-target {
  min-height: 44px;
  min-width: 44px;
  display: flex;
  align-items: center;
  justify-content: center;
}

/* Hover states only on non-touch devices */
@media (hover: hover) {
  .btn:hover {
    background-color: var(--dark-blue);
  }
}
```

## 🌍 Localization Design

### Vietnamese Text Considerations

#### Text Length Variations
- Account for longer Vietnamese text in UI layouts
- Ensure buttons and labels can accommodate text expansion
- Use flexible layouts that adapt to content length

#### Cultural Adaptations
- Use appropriate currency formatting (₫)
- Follow Vietnamese date formats (DD/MM/YYYY)
- Use familiar Vietnamese financial terminology

#### Reading Patterns
- Left-to-right reading flow
- Clear visual hierarchy for scanning
- Important information positioned prominently

## ♿ Accessibility Standards

### WCAG 2.1 Compliance

#### Color Contrast
- Minimum 4.5:1 for normal text
- Minimum 3:1 for large text (18px+)
- Never rely on color alone to convey information

#### Keyboard Navigation
```css
.focusable:focus {
  outline: 2px solid var(--primary-blue);
  outline-offset: 2px;
}

/* Skip to main content link */
.skip-link {
  position: absolute;
  top: -40px;
  left: 6px;
  background: var(--primary-blue);
  color: white;
  padding: 8px;
  text-decoration: none;
  z-index: 1000;
}

.skip-link:focus {
  top: 6px;
}
```

#### Screen Reader Support
```html
<!-- Descriptive labels -->
<label for="income-amount" id="income-label">
  Số tiền thu nhập (VND)
</label>
<input 
  type="number" 
  id="income-amount"
  aria-labelledby="income-label"
  aria-describedby="income-help"
  required
>
<div id="income-help" class="help-text">
  Nhập số tiền thu nhập không bao gồm dấu phẩy
</div>
```

## 🎨 Design Tokens

### CSS Custom Properties
```css
:root {
  /* Colors */
  --color-primary: #4472C4;
  --color-success: #70AD47;
  --color-warning: #FFC000;
  --color-error: #E74C3C;
  
  /* Typography */
  --font-family-base: 'Arial', sans-serif;
  --font-size-base: 14px;
  --line-height-base: 1.4;
  
  /* Spacing */
  --spacing-unit: 8px;
  --border-radius: 4px;
  --border-width: 1px;
  
  /* Shadows */
  --shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.1);
  --shadow-md: 0 2px 6px rgba(0, 0, 0, 0.1);
  --shadow-lg: 0 4px 12px rgba(0, 0, 0, 0.1);
  
  /* Transitions */
  --transition-fast: 0.15s ease;
  --transition-normal: 0.2s ease;
  --transition-slow: 0.3s ease;
}
```

## 📏 Quality Checklist

### Design Review Criteria

#### Visual Consistency
- [ ] Colors match design system
- [ ] Typography follows hierarchy
- [ ] Spacing follows 8px grid
- [ ] Icons are consistent style

#### Usability
- [ ] Forms are easy to complete
- [ ] Error messages are helpful
- [ ] Navigation is intuitive
- [ ] Touch targets are appropriate size

#### Accessibility
- [ ] Color contrast meets WCAG standards
- [ ] Keyboard navigation works
- [ ] Screen reader friendly markup
- [ ] Focus states are visible

#### Responsiveness
- [ ] Works on mobile devices
- [ ] Content adapts to screen size
- [ ] Touch interactions are smooth
- [ ] Performance is acceptable

#### Vietnamese Localization
- [ ] All text is in Vietnamese
- [ ] Cultural context is appropriate
- [ ] Financial terms are accurate
- [ ] Date/number formats are correct

## 🔄 Design System Evolution

### Maintenance Process
1. **Regular Review**: Quarterly design system audits
2. **User Feedback**: Incorporate user testing insights
3. **Technology Updates**: Adapt to new platform features
4. **Documentation**: Keep guidelines updated with changes

### Contributing to Design System
- Document new patterns before implementation
- Share design decisions with the team
- Test changes across multiple devices
- Consider accessibility impact of changes

This design system ensures that HODLVN-Family-Finance maintains a consistent, professional, and user-friendly experience that serves Vietnamese users effectively across all touchpoints.