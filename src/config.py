"""
Configuration constants for the Budget Tracker.

CUSTOMIZATION GUIDE:
- Modify CATEGORIES and BUDGETS to set your expense categories and amounts
- Modify CURRENCY_OPTIONS to change available currencies
- Modify YEAR to change the tracking year
- Modify OUTPUT_FILE to change the output file name
"""

# =============================================================================
# OUTPUT SETTINGS
# =============================================================================
OUTPUT_FILE = "Budget_Tracker.xlsx"
YEAR = 2026

# =============================================================================
# BUDGET CATEGORIES AND AMOUNTS
# Modify these lists to customize your budget
# The order of BUDGETS must match the order of CATEGORIES
# =============================================================================
CATEGORIES = [
    'Rent + Utilities',
    'Electricity Deposit',
    'Internet & Phone Bills',
    'Food & Groceries',
    'Investments',
    'Date Nights',
    'Annual Costs Savings',
    'Emergency Buffer',
    'Personal Care',
    'Personal Savings',
    'Miscellaneous'
]

# Monthly budget amounts (must have same length as CATEGORIES)
BUDGETS = [34300, 2884, 1200, 6000, 4000, 1500, 1816, 1000, 1000, 4300, 2000]

# =============================================================================
# CURRENCY OPTIONS
# =============================================================================
CURRENCY_OPTIONS = ['CZK', 'USD', 'EUR', 'NGN']
DEFAULT_CURRENCY = 'CZK'

# =============================================================================
# MONTHS AND DAYS
# =============================================================================
MONTHS = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
]

# Days in each month (update February to 29 for leap years)
DAYS_IN_MONTH = {
    'January': 31, 'February': 28, 'March': 31, 'April': 30,
    'May': 31, 'June': 30, 'July': 31, 'August': 31,
    'September': 30, 'October': 31, 'November': 30, 'December': 31
}

# =============================================================================
# COLOR SCHEMES
# =============================================================================

# Month colors (unique color for each month)
MONTH_COLORS = {
    'January': '#1E3A5F',    # Deep Blue
    'February': '#8E44AD',   # Purple
    'March': '#27AE60',      # Green
    'April': '#F39C12',      # Orange
    'May': '#E91E63',        # Pink
    'June': '#00BCD4',       # Cyan
    'July': '#FF5722',       # Deep Orange
    'August': '#4CAF50',     # Light Green
    'September': '#9C27B0',  # Deep Purple
    'October': '#FF9800',    # Amber
    'November': '#795548',   # Brown
    'December': '#607D8B'    # Blue Grey
}

# Category colors for expense coding
CATEGORY_COLORS = {
    'Rent + Utilities': '#E74C3C',           # Red
    'Electricity Deposit': '#C0392B',        # Dark Red
    'Internet & Phone Bills': '#3498DB',     # Blue
    'Food & Groceries': '#27AE60',           # Green
    'Investments': '#1ABC9C',                # Teal
    'Date Nights': '#FF6B6B',                # Coral
    'Annual Costs Savings': '#34495E',       # Dark Grey
    'Emergency Buffer': '#95A5A6',           # Grey
    'Personal Care': '#FF9800',              # Amber
    'Personal Savings': '#2ECC71',           # Emerald Green
    'Miscellaneous': '#607D8B'               # Blue Grey
}

# Currency display colors
CURRENCY_COLORS = {
    'CZK': '#1E3A5F',
    'USD': '#27AE60',
    'EUR': '#3498DB',
    'NGN': '#9B59B6'
}

# =============================================================================
# UI SETTINGS
# =============================================================================
UI_FONT = 'Calibri'

# Payment methods for expense entries
PAYMENT_METHODS = ['Cash', 'Card', 'Transfer', 'Online']

# =============================================================================
# INCOME SETTINGS
# =============================================================================
DEFAULT_MONTHLY_SALARY = 60000
TARGET_SAVINGS_MIN = 5000
TARGET_SAVINGS_MAX = 10000
