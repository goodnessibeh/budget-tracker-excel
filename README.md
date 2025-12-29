# Budget Tracker for Excel

A world-class expense management and budget tracking system built with Python and xlsxwriter. Generates a comprehensive Excel workbook with dynamic dashboards, charts, analytics, and multi-currency freelance income tracking.

## Features

### Dashboard
- **Real-time YTD Metrics**: Annual budget, spent, remaining, and budget utilization
- **Month Selector**: View data for ALL TIME or any specific month (January-December)
- **Currency Selector**: Set global currency (CZK, USD, EUR, NGN)
- **Financial Discipline Scoring**: Track opening/closing balances with automatic scoring
- **3 Embedded Charts**: Expense trend, budget allocation pie chart, surplus/deficit chart
- **Dynamic Category Breakdown**: Updates based on selected month view

### Freelance Income Tracking
- **Per-Month Currency Selection**: Each month's freelance income can be in different currencies
- **Freelance Earnings by Currency**: Dashboard summary showing totals for CZK, USD, EUR, NGN
- **Automatic Calculations**: Total earned, months active, average per month, percentage of total

### Monthly Expense Sheets (12 sheets)
- **Pre-populated Dates**: Every day from January 1 to December 31
- **Weekend Highlighting**: Saturdays and Sundays highlighted in light red
- **Day Names**: Mon, Tue, Wed, etc. shown for easy reference
- **Category Dropdowns**: Select from customizable budget categories
- **Payment Method**: Cash, Card, Transfer, Online options
- **Running Totals**: Automatic cumulative expense calculation
- **Category Breakdown**: Per-month summary with status indicators (OK/WARN/OVER)

### Example Budget Categories
The script comes with example categories that you can customize:

| Category | Example Monthly Budget |
|----------|------------------------|
| Housing | 1,500 |
| Utilities | 200 |
| Internet & Phone | 100 |
| Food & Groceries | 600 |
| Transportation | 300 |
| Entertainment | 200 |
| Savings | 500 |
| Emergency Fund | 200 |
| Personal Care | 100 |
| Investments | 300 |
| Miscellaneous | 200 |

*Note: These are example values. See [Customization](#customization) to set your own budgets.*

### Charts & Analytics
10 professional charts including:
- Budget Allocation Pie Chart
- Monthly Budget vs Actual Column Chart
- Spending Distribution Doughnut Chart
- Budget Utilization Horizontal Bar Chart
- Expense Trend Line Chart
- Monthly Variance Chart
- Cumulative Expense Area Chart
- Income vs Expenses Chart
- Surplus Tracker Chart

### Additional Features
- **Frozen Headers**: All sheets have frozen panes for easy scrolling
- **Modern UI Font**: Calibri font throughout
- **Color-coded Categories**: Each category has a unique color
- **Color-coded Months**: Each month has a distinct color theme
- **Conditional Formatting**: Automatic alerts for budget status
- **Data Validation**: Dropdown lists for categories and payment methods

## Requirements

- Python 3.8+
- xlsxwriter library

## Installation

1. Clone the repository:
```bash
git clone https://github.com/goodnessibeh/budget-tracker-excel.git
cd budget-tracker-excel
```

2. Create a virtual environment (recommended):
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install xlsxwriter
```

## Usage

Run the module to generate the Excel file:

```bash
python -m src.main
```

This creates `Budget_Tracker.xlsx` in the current directory.

## Project Structure

```
budget-tracker-excel/
├── README.md                    # This file
├── src/
│   ├── __init__.py
│   ├── config.py                # Configuration constants
│   ├── formats.py               # Excel format definitions
│   ├── main.py                  # Core generation logic
│   └── sheets/
│       ├── __init__.py
│       ├── dashboard.py         # Dashboard sheet
│       ├── monthly_budget.py    # Budget allocation sheet
│       ├── monthly_expense.py   # 12 monthly expense sheets
│       ├── budget_vs_actual.py  # Comparison sheet
│       ├── annual_summary.py    # Annual summary sheet
│       ├── charts.py            # Charts & analytics sheet
│       └── category_legend.py   # Legend sheet
└── Budget_Tracker.xlsx          # Generated Excel file
```

## Customization

The modular structure makes customization easy. Edit files in the `src/` directory:

### 1. Change Output File Name and Year (`src/config.py`)

```python
OUTPUT_FILE = "My_Budget_2025.xlsx"
YEAR = 2025
```

### 2. Modify Budget Categories (`src/config.py`)

```python
CATEGORIES = [
    'Housing',
    'Utilities',
    'Food & Groceries',
    # Add your categories...
]

# Must match the order of CATEGORIES
BUDGETS = [1500, 200, 600, ...]
```

### 3. Change Category Colors (`src/config.py`)

```python
CATEGORY_COLORS = {
    'Housing': '#E74C3C',      # Red
    'Utilities': '#3498DB',    # Blue
    # Add colors for each category...
}
```

### 4. Modify Currency Options (`src/config.py`)

```python
CURRENCY_OPTIONS = ['USD', 'EUR', 'GBP', 'JPY']
DEFAULT_CURRENCY = 'USD'
```

### 5. Adjust for Leap Years (`src/config.py`)

```python
DAYS_IN_MONTH = {
    'February': 29,  # Change to 29 for leap years
    # ...
}
```

### Key Configuration Files

| File | Purpose |
|------|---------|
| `src/config.py` | All constants: categories, budgets, colors, year |
| `src/formats.py` | Excel cell formats and styles |
| `src/sheets/dashboard.py` | Dashboard layout and logic |
| `src/sheets/monthly_expense.py` | Monthly sheet structure |

## Excel Workbook Structure

The generated workbook contains 19 sheets:

1. **Dashboard** - Main overview with metrics and charts
2. **Monthly Budget** - Budget allocation by category
3. **January - December** - 12 monthly expense sheets
4. **Budget vs Actual** - YTD comparison
5. **Annual Summary** - Monthly expense summary
6. **Charts & Analytics** - Visual analytics
7. **Category Legend** - Color codes and descriptions

## License

MIT License - feel free to use and modify for personal or commercial use.

## Author

Created by Goodness Ibeh

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -m 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Open a Pull Request
