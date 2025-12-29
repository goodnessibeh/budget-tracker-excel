#!/usr/bin/env python3
"""
Budget Tracker for Excel
A world-class expense management and budget tracking system.

Usage:
    python -m src.main
    python src/main.py
"""

import os
import xlsxwriter

from .config import OUTPUT_FILE, YEAR, MONTHS, DAYS_IN_MONTH, CATEGORIES
from .formats import FormatManager
from .sheets import (
    create_dashboard,
    create_monthly_budget,
    create_monthly_expense_sheets,
    create_budget_vs_actual,
    create_annual_summary,
    create_charts_sheet,
    create_category_legend,
)


def create_budget_tracker(output_path=None):
    """
    Create the complete budget tracker Excel workbook.

    Args:
        output_path: Optional custom output path. Defaults to OUTPUT_FILE from config.

    Returns:
        str: Path to the created file.
    """
    if output_path is None:
        output_path = OUTPUT_FILE

    # Create workbook
    workbook = xlsxwriter.Workbook(output_path)

    # Initialize format manager
    formats = FormatManager(workbook)

    # Create all sheets
    dashboard = create_dashboard(workbook, formats)
    create_monthly_budget(workbook, formats)
    create_monthly_expense_sheets(workbook, formats)
    create_budget_vs_actual(workbook, formats)
    create_annual_summary(workbook, formats)
    create_charts_sheet(workbook, formats)
    create_category_legend(workbook, formats)

    # Set Dashboard as active sheet
    dashboard.activate()

    # Close workbook
    workbook.close()

    return output_path


def print_summary(output_path):
    """Print creation summary."""
    print("=" * 70)
    print("  BUDGET TRACKER CREATED SUCCESSFULLY!")
    print("=" * 70)
    print(f"\n  File: {output_path}")
    print(f"\n  TRACKING PERIOD: January 1, {YEAR} - December 31, {YEAR} (365 days)")
    print(f"\n  FEATURES INCLUDED:")
    print("  " + "-" * 50)
    print("  - Dashboard with real-time YTD metrics")
    print("  - Month selector with ALL TIME option")
    print("  - Currency selector (CZK, USD, EUR, NGN)")
    print("  - Dynamic category breakdown based on selection")
    print("  - 3 embedded dashboard charts")
    print("  - Account balance tracking (opening/closing)")
    print("  - Freelance income tracking with currency selection")
    print("  - Freelance earnings by currency summary")
    print("  - 12 monthly expense sheets with pre-populated dates")
    print("  - Weekend days highlighted")
    print("  - Day names shown (Mon, Tue, Wed, etc.)")
    print(f"  - Color-coded categories ({len(CATEGORIES)} categories)")
    print("  - Color-coded months (unique color per month)")
    print("  - Budget vs Actual comparison")
    print("  - Annual summary with monthly trends")
    print("  - 10 professional charts in Charts & Analytics")
    print("  - Conditional formatting alerts")
    print("  - Category dropdown validation")
    print("  - Financial discipline scoring")
    print("  - Frozen headers on all sheets")
    print("  - Modern UI font (Calibri)")
    print("  " + "-" * 50)
    print("\n  DASHBOARD FEATURES:")
    print("    - Select VIEW: ALL TIME or any month (Jan-Dec)")
    print("    - Select CURRENCY: CZK, USD, EUR, NGN (default: CZK)")
    print("    - Dynamic category breakdown updates automatically")
    print("\n  MONTHLY SHEETS:")
    for month in MONTHS:
        print(f"    - {month} {YEAR}: {DAYS_IN_MONTH[month]} days")
    print()


def main():
    """Main entry point."""
    output_path = create_budget_tracker()
    print_summary(output_path)


if __name__ == "__main__":
    main()
