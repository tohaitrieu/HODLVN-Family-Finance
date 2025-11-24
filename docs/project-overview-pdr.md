# HODLVN-Family-Finance: Project Overview (PDR)

## Problem Statement

Vietnamese families and individuals face significant challenges in managing their personal finances:

- **Lack of localized tools**: Most financial management applications are designed for Western markets with different financial contexts
- **Language barriers**: International tools use English terminology that doesn't translate well to Vietnamese financial concepts  
- **Cost barriers**: Premium financial tools are expensive for average Vietnamese users
- **Complex financial ecosystem**: Need to track multiple asset classes (stocks, gold, crypto) popular in Vietnam
- **Debt management complexity**: Vietnamese lending practices (informal loans, family loans) need special handling
- **Manual tracking burden**: Without automation, maintaining financial records is time-consuming

## Domain Context

### Personal Finance Management in Vietnam

The Vietnamese personal finance landscape has unique characteristics:

1. **Multi-asset portfolios**: Vietnamese investors commonly hold diverse assets:
   - Traditional stocks (VN-Index)
   - Physical gold (SJC gold bars)
   - Cryptocurrency (growing adoption)
   - Real estate investments
   - Informal lending networks

2. **Cash-based society transitioning to digital**: While cash remains dominant, digital payments are growing rapidly

3. **Family-centric financial planning**: Extended family financial obligations are common

4. **High savings rate**: Vietnamese households save 20-30% of income on average

### Technology Platform

Google Sheets was chosen as the platform because:
- **Free access**: Available to anyone with a Google account
- **Familiar interface**: Many Vietnamese users already use Google Sheets
- **Cloud-based**: Accessible from any device
- **Collaborative**: Supports family financial planning
- **Extensible**: Google Apps Script enables automation

## Business Requirements

### Core Requirements

1. **Transaction Management**
   - Track all income sources (salary, business, investments)
   - Categorize expenses with Vietnamese spending patterns
   - Support multiple transaction types

2. **Debt & Lending**
   - Track formal and informal loans
   - Calculate interest (simple and compound)
   - Payment scheduling and reminders
   - Support for lending to others

3. **Investment Portfolio**
   - Track stocks with VN market specifics
   - Gold investment tracking (by weight/value)
   - Cryptocurrency portfolio management
   - Other investments (real estate, bonds)

4. **Budgeting & Planning**
   - Monthly/yearly budget setting
   - Real-time budget tracking
   - Visual alerts for overspending
   - Cash flow projections

5. **Analytics & Reporting**
   - Real-time dashboard
   - Income vs expense analysis
   - Investment performance
   - Net worth tracking

## Functional Requirements

### 1. Income Management (THU)
- Record multiple income types
- Support recurring income
- Categorize by source
- Auto-calculate totals

### 2. Expense Tracking (CHI)
- Vietnamese expense categories
- Receipt attachment capability
- Recurring expense support
- Payment method tracking

### 3. Debt Management (QUẢN LÝ NỢ)
- Create loan records
- Interest calculation
- Payment scheduling
- Auto-generate income entries for loans

### 4. Debt Payment (TRẢ NỢ)
- Track principal and interest
- Partial payment support
- Payment history
- Outstanding balance calculation

### 5. Stock Investment (CHỨNG KHOÁN)
- Buy/sell transactions
- Portfolio valuation
- Profit/loss tracking
- Margin trading support
- Dividend tracking

### 6. Gold Investment (VÀNG)
- Track by weight (chỉ, lượng)
- Current price integration
- Buy/sell history
- P&L calculation

### 7. Cryptocurrency (CRYPTO)
- Multi-coin support
- Transaction history
- Portfolio valuation
- Exchange tracking

### 8. Other Investments (ĐẦU TƯ KHÁC)
- Real estate tracking
- Bond investments
- Mutual funds
- Custom investment types

### 9. Budget Management (BUDGET)
- Category-based budgets
- Visual progress indicators
- Alert thresholds (70%, 90%)
- Budget vs actual reporting

### 10. Dashboard & Analytics
- Real-time overview
- Cash flow visualization
- Trend analysis
- Period comparisons

## Non-Functional Requirements

### Performance
- Form submission < 3 seconds
- Dashboard refresh < 5 seconds
- Support 10,000+ transactions

### Usability
- Vietnamese UI/UX
- Mobile-responsive forms
- Intuitive navigation
- Minimal training required

### Security
- Data stored in user's Google Drive
- No external data transmission
- Google account authentication

### Reliability
- 99.9% uptime (Google infrastructure)
- Automatic formula preservation
- Data validation

### Maintainability
- Modular code architecture
- Clear documentation
- Version control via GitHub

## Technical Architecture

### Technology Stack
- **Platform**: Google Sheets + Google Apps Script
- **Backend**: Google Apps Script (ES6)
- **Frontend**: HTML5, CSS3, JavaScript
- **UI Framework**: Custom responsive forms
- **Database**: Google Sheets as data store

### Architecture Pattern
Modular Architecture v3.5.8 with clear separation of concerns:

```
┌─────────────────────────────────────────────────┐
│              GOOGLE SHEETS UI                    │
│            (10 Sheets + Dashboard)               │
└──────────────────┬──────────────────────────────┘
                   │
┌──────────────────▼──────────────────────────────┐
│             MAIN MENU (Main.gs)                  │
│          Menu Dispatcher & Router                │
└─┬────────┬────────┬────────┬────────┬──────────┘
  │        │        │        │        │
┌─▼──┐ ┌──▼──┐ ┌──▼──┐ ┌───▼───┐ ┌──▼──┐
│Sheet│ │Data │ │Budget│ │Dashboard│ │Utils│
│Init │ │Proc │ │Mgr   │ │Manager  │ │     │
└────┘ └─────┘ └──────┘ └─────────┘ └─────┘
```

### Key Modules
1. **Main.gs**: Menu system and routing
2. **SheetInitializer.gs**: Sheet creation and setup
3. **DataProcessor.gs**: Transaction processing
4. **BudgetManager.gs**: Budget tracking and alerts
5. **DashboardManager.gs**: Analytics and reporting
6. **Utils.gs**: Shared utilities

## User Personas

### 1. The Family Financial Manager
- **Name**: Nguyễn Thị Hương
- **Age**: 35
- **Role**: Manages finances for family of 4
- **Needs**: Track household expenses, plan for children's education, manage family investments
- **Tech Level**: Basic spreadsheet user

### 2. The Young Professional
- **Name**: Trần Văn Nam
- **Age**: 28
- **Role**: IT professional with growing income
- **Needs**: Track spending, invest in stocks/crypto, plan for apartment purchase
- **Tech Level**: Advanced user

### 3. The Small Business Owner
- **Name**: Lê Minh Tuấn
- **Age**: 42
- **Role**: Owns small trading business
- **Needs**: Separate personal/business finances, track loans, manage cash flow
- **Tech Level**: Intermediate user

## Use Cases

### UC1: Record Monthly Salary
1. User opens Income form
2. Selects "Lương" category
3. Enters amount and date
4. System records transaction
5. Dashboard updates automatically

### UC2: Track Stock Investment
1. User opens Stock form
2. Enters buy transaction
3. System calculates fees
4. Updates portfolio value
5. Shows P&L on dashboard

### UC3: Manage Loan Payment
1. User creates loan record
2. System generates payment schedule
3. User records payment
4. System updates outstanding balance
5. Sends reminder for next payment

### UC4: Budget Monitoring
1. User sets monthly budget
2. Records daily expenses
3. System tracks progress
4. Shows color-coded alerts
5. Provides spending insights

## Constraints & Assumptions

### Technical Constraints
- Limited by Google Apps Script quotas
- 6-minute execution timeout
- 50MB spreadsheet size limit
- No real-time external data feeds

### Business Constraints
- Must remain 100% free
- Vietnamese language only initially
- Single-user focus (no multi-tenancy)
- Manual data entry (no bank integration)

### Assumptions
- Users have Google accounts
- Basic spreadsheet knowledge
- Internet connectivity
- Vietnamese financial context

## Success Metrics

### Adoption Metrics
- 1,000+ active users in year 1
- 50+ GitHub stars
- 20+ community contributors

### Usage Metrics
- Average 5+ transactions/day per user
- 80% users check dashboard weekly
- 60% use budget features

### Business Impact
- Users save 10%+ more income
- 90% user satisfaction rating
- 50% reduction in financial tracking time

### Quality Metrics
- Zero data loss incidents
- < 5 critical bugs per release
- 95% uptime availability