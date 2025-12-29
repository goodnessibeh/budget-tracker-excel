"""Charts & Analytics sheet creation module."""

from ..config import CATEGORIES, UI_FONT


def create_charts_sheet(workbook, formats):
    """Create the charts and analytics sheet."""
    sheet = workbook.add_worksheet('Charts & Analytics')
    sheet.set_tab_color('#9B59B6')

    # Column widths
    sheet.set_column('A:A', 3)
    sheet.set_column('B:Z', 16)

    # Title
    sheet.merge_range('B1:N1', 'VISUAL ANALYTICS & CHARTS', formats.title)
    sheet.set_row(0, 45)

    # Budget Allocation Pie Chart
    sheet.merge_range('B2:H2', 'BUDGET ALLOCATION OVERVIEW', formats.sub_header)

    pie_chart = workbook.add_chart({'type': 'pie'})
    pie_chart.add_series({
        'name': 'Budget Allocation',
        'categories': f"='Monthly Budget'!$C$3:$C${len(CATEGORIES)+2}",
        'values': f"='Monthly Budget'!$D$3:$D${len(CATEGORIES)+2}",
        'data_labels': {'percentage': True, 'category': False, 'font': {'size': 9, 'name': UI_FONT}},
    })
    pie_chart.set_title({'name': 'Monthly Budget Allocation', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    pie_chart.set_style(10)
    pie_chart.set_size({'width': 520, 'height': 380})
    sheet.insert_chart('B3', pie_chart)

    # Monthly Budget vs Actual Bar Chart
    sheet.merge_range('J2:R2', 'MONTHLY BUDGET VS ACTUAL', formats.sub_header)

    bar_chart = workbook.add_chart({'type': 'column'})
    bar_chart.add_series({
        'name': 'Actual Expenses',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$E$4:$E$15",
        'fill': {'color': '#E74C3C'},
        'gap': 150
    })
    bar_chart.add_series({
        'name': 'Budget',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$F$4:$F$15",
        'fill': {'color': '#3498DB'},
    })
    bar_chart.set_title({'name': 'Monthly Budget vs Actual Expenses', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    bar_chart.set_x_axis({'name': 'Month', 'num_font': {'rotation': -45, 'size': 9, 'name': UI_FONT}})
    bar_chart.set_y_axis({'name': 'Amount', 'num_font': {'size': 9, 'name': UI_FONT}})
    bar_chart.set_style(10)
    bar_chart.set_size({'width': 720, 'height': 380})
    bar_chart.set_legend({'position': 'top', 'font': {'name': UI_FONT, 'size': 10}})
    sheet.insert_chart('J3', bar_chart)

    # Spending Distribution Doughnut
    sheet.merge_range('B23:H23', 'SPENDING DISTRIBUTION (YTD)', formats.sub_header)

    doughnut = workbook.add_chart({'type': 'doughnut'})
    doughnut.add_series({
        'name': 'Spending by Category',
        'categories': f"='Budget vs Actual'!$B$4:$B${len(CATEGORIES)+3}",
        'values': f"='Budget vs Actual'!$D$4:$D${len(CATEGORIES)+3}",
        'data_labels': {'percentage': True, 'font': {'size': 8, 'name': UI_FONT}},
    })
    doughnut.set_title({'name': 'YTD Spending by Category', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    doughnut.set_style(10)
    doughnut.set_size({'width': 520, 'height': 380})
    doughnut.set_hole_size(50)
    sheet.insert_chart('B24', doughnut)

    # Budget Utilization Bar
    sheet.merge_range('J23:R23', 'BUDGET UTILIZATION BY CATEGORY', formats.sub_header)

    util_chart = workbook.add_chart({'type': 'bar'})
    util_chart.add_series({
        'name': '% Used',
        'categories': f"='Budget vs Actual'!$B$4:$B${len(CATEGORIES)+3}",
        'values': f"='Budget vs Actual'!$F$4:$F${len(CATEGORIES)+3}",
        'fill': {'color': '#9B59B6'},
        'data_labels': {'value': True, 'font': {'size': 9, 'name': UI_FONT}, 'num_format': '0%'}
    })
    util_chart.set_title({'name': 'Budget Utilization by Category', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    util_chart.set_x_axis({'name': 'Percentage Used', 'min': 0, 'max': 1.5, 'num_font': {'size': 9, 'name': UI_FONT}})
    util_chart.set_y_axis({'reverse': True, 'num_font': {'size': 9, 'name': UI_FONT}})
    util_chart.set_style(10)
    util_chart.set_size({'width': 720, 'height': 380})
    util_chart.set_legend({'none': True})
    sheet.insert_chart('J24', util_chart)

    # Expense Trend Line Chart
    sheet.merge_range('B44:N44', 'EXPENSE TREND ANALYSIS', formats.sub_header)

    trend_chart = workbook.add_chart({'type': 'line'})
    trend_chart.add_series({
        'name': 'Expenses',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$E$4:$E$15",
        'line': {'color': '#E74C3C', 'width': 2.5},
        'marker': {'type': 'circle', 'size': 8, 'fill': {'color': '#E74C3C'}}
    })
    trend_chart.add_series({
        'name': 'Budget',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$F$4:$F$15",
        'line': {'color': '#3498DB', 'width': 2.5, 'dash_type': 'dash'},
        'marker': {'type': 'diamond', 'size': 6, 'fill': {'color': '#3498DB'}}
    })
    trend_chart.set_title({'name': 'Expense Trend', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    trend_chart.set_x_axis({'name': 'Month', 'num_font': {'size': 9, 'name': UI_FONT}})
    trend_chart.set_y_axis({'name': 'Amount', 'num_font': {'size': 9, 'name': UI_FONT}})
    trend_chart.set_style(10)
    trend_chart.set_size({'width': 950, 'height': 400})
    trend_chart.set_legend({'position': 'bottom', 'font': {'name': UI_FONT, 'size': 10}})
    sheet.insert_chart('B45', trend_chart)

    # Variance Chart
    sheet.merge_range('B66:H66', 'MONTHLY VARIANCE', formats.sub_header)

    variance_chart = workbook.add_chart({'type': 'column'})
    variance_chart.add_series({
        'name': 'Variance (+ is under budget)',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$G$4:$G$15",
        'fill': {'color': '#27AE60'},
        'data_labels': {'value': True, 'font': {'size': 8, 'name': UI_FONT}, 'num_format': '#,##0'}
    })
    variance_chart.set_title({'name': 'Monthly Budget Variance', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    variance_chart.set_x_axis({'name': 'Month', 'num_font': {'rotation': -45, 'size': 9, 'name': UI_FONT}})
    variance_chart.set_y_axis({'name': 'Variance', 'num_font': {'size': 9, 'name': UI_FONT}})
    variance_chart.set_style(10)
    variance_chart.set_size({'width': 720, 'height': 380})
    variance_chart.set_legend({'none': True})
    sheet.insert_chart('B67', variance_chart)

    # Cumulative Expense Area Chart
    sheet.merge_range('J66:R66', 'CUMULATIVE EXPENSE TRACKING', formats.sub_header)

    cumulative_chart = workbook.add_chart({'type': 'area'})
    cumulative_chart.add_series({
        'name': 'Cumulative Expenses',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$E$4:$E$15",
        'fill': {'color': '#3498DB', 'transparency': 30},
        'line': {'color': '#2980B9', 'width': 2}
    })
    cumulative_chart.set_title({'name': 'Cumulative Monthly Expenses', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    cumulative_chart.set_x_axis({'name': 'Month', 'num_font': {'size': 9, 'name': UI_FONT}})
    cumulative_chart.set_y_axis({'name': 'Amount', 'num_font': {'size': 9, 'name': UI_FONT}})
    cumulative_chart.set_style(10)
    cumulative_chart.set_size({'width': 720, 'height': 380})
    cumulative_chart.set_legend({'none': True})
    sheet.insert_chart('J67', cumulative_chart)

    # Income vs Expense Chart
    sheet.merge_range('B88:H88', 'INCOME VS EXPENSES', formats.sub_header)

    income_expense = workbook.add_chart({'type': 'column'})
    income_expense.add_series({
        'name': 'Total Income',
        'categories': "=Dashboard!$I$24:$I$35",
        'values': "=Dashboard!$L$24:$L$35",
        'fill': {'color': '#27AE60'},
    })
    income_expense.add_series({
        'name': 'Total Expenses',
        'categories': "=Dashboard!$B$30:$B$41",
        'values': "=Dashboard!$D$30:$D$41",
        'fill': {'color': '#E74C3C'},
    })
    income_expense.set_title({'name': 'Monthly Income vs Expenses', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    income_expense.set_x_axis({'num_font': {'rotation': -45, 'size': 9, 'name': UI_FONT}})
    income_expense.set_y_axis({'name': 'Amount', 'num_font': {'size': 9, 'name': UI_FONT}})
    income_expense.set_style(10)
    income_expense.set_size({'width': 720, 'height': 380})
    income_expense.set_legend({'position': 'top', 'font': {'name': UI_FONT, 'size': 10}})
    sheet.insert_chart('B89', income_expense)

    # Surplus Tracker
    sheet.merge_range('J88:R88', 'SURPLUS TRACKER', formats.sub_header)

    surplus_chart = workbook.add_chart({'type': 'column'})
    surplus_chart.add_series({
        'name': 'Budget Savings',
        'categories': "=Dashboard!$B$30:$B$41",
        'values': "=Dashboard!$E$30:$E$41",
        'fill': {'color': '#3498DB'},
    })
    surplus_chart.add_series({
        'name': 'Freelance Income',
        'categories': "=Dashboard!$B$30:$B$41",
        'values': "=Dashboard!$F$30:$F$41",
        'fill': {'color': '#9B59B6'},
    })
    surplus_chart.add_series({
        'name': 'Total Surplus',
        'categories': "=Dashboard!$B$30:$B$41",
        'values': "=Dashboard!$G$30:$G$41",
        'line': {'color': '#27AE60', 'width': 2.5},
        'marker': {'type': 'circle', 'size': 6, 'fill': {'color': '#27AE60'}}
    })
    surplus_chart.set_title({'name': 'Monthly Surplus Breakdown', 'name_font': {'size': 14, 'bold': True, 'name': UI_FONT}})
    surplus_chart.set_x_axis({'num_font': {'rotation': -45, 'size': 9, 'name': UI_FONT}})
    surplus_chart.set_y_axis({'name': 'Amount', 'num_font': {'size': 9, 'name': UI_FONT}})
    surplus_chart.set_style(10)
    surplus_chart.set_size({'width': 720, 'height': 380})
    surplus_chart.set_legend({'position': 'top', 'font': {'name': UI_FONT, 'size': 10}})
    sheet.insert_chart('J89', surplus_chart)

    # Freeze title row
    sheet.freeze_panes(1, 0)

    return sheet
