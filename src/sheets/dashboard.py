"""Dashboard sheet creation module."""

from datetime import date
from ..config import (
    CATEGORIES, BUDGETS, MONTHS, DAYS_IN_MONTH,
    CURRENCY_OPTIONS, CURRENCY_COLORS, UI_FONT, YEAR
)


def create_dashboard(workbook, formats):
    """Create the main dashboard sheet."""
    dashboard = workbook.add_worksheet('Dashboard')
    dashboard.set_tab_color('#1E3A5F')

    # Set column widths
    dashboard.set_column('A:A', 3)
    dashboard.set_column('B:B', 30)
    dashboard.set_column('C:C', 20)
    dashboard.set_column('D:D', 20)
    dashboard.set_column('E:E', 20)
    dashboard.set_column('F:F', 20)
    dashboard.set_column('G:G', 20)
    dashboard.set_column('H:H', 3)
    dashboard.set_column('I:I', 28)
    dashboard.set_column('J:J', 18)
    dashboard.set_column('K:K', 18)
    dashboard.set_column('L:L', 18)
    dashboard.set_column('M:M', 18)
    dashboard.set_column('N:N', 14)
    dashboard.set_column('O:O', 14)
    dashboard.set_column('P:P', 18)
    dashboard.set_column('Q:Q', 12)
    dashboard.set_column('R:R', 18)
    dashboard.set_column('S:S', 14)

    # Title
    dashboard.merge_range('B2:G2', f"FINANCIAL DASHBOARD {YEAR}", formats.title)
    dashboard.merge_range('B3:G3', f'January 1st - December 31st, {YEAR}', formats.subtitle)

    # Month Selector
    dashboard.merge_range('I2:J2', 'VIEW:', formats.selector)
    dashboard.merge_range('K2:L2', 'ALL TIME', formats.selector_input)
    dashboard.write_comment('K2', 'Select a month (January-December) or ALL TIME to view data')

    month_options = ['ALL TIME'] + MONTHS
    dashboard.data_validation('K2:L2', {
        'validate': 'list',
        'source': month_options,
        'dropdown': True
    })

    # Currency Selector
    dashboard.write('M2', 'CURRENCY:', formats.currency)
    dashboard.write('N2', 'CZK', formats.currency_input)
    dashboard.write_comment('N2', 'Select currency: CZK, USD, EUR, NGN')

    dashboard.data_validation('N2', {
        'validate': 'list',
        'source': CURRENCY_OPTIONS,
        'dropdown': True
    })

    # Hidden currency lookup
    dashboard.write('AA1', 'Currency', formats.text)
    dashboard.write('AB1', 'Display', formats.text)
    for i, curr in enumerate(CURRENCY_OPTIONS):
        dashboard.write(f'AA{i+2}', curr, formats.text)
        dashboard.write(f'AB{i+2}', curr, formats.text)
    dashboard.set_column('AA:AB', None, None, {'hidden': True})

    # Period info
    dashboard.write('B5', 'Tracking Period:', formats.sub_header)
    dashboard.merge_range('C5:E5', f'January 1, {YEAR} - December 31, {YEAR}',
                         formats.create_custom_format({
                             'bold': True, 'font_size': 12, 'align': 'center',
                             'border': 1, 'bg_color': '#D4E6F1'
                         }))
    dashboard.write('F5', 'Total Days:', formats.sub_header)
    dashboard.write('G5', 365, formats.create_custom_format({
        'bold': True, 'font_size': 12, 'align': 'center', 'border': 1, 'bg_color': '#D4E6F1'
    }))

    # Account Balance Section
    _create_account_balance_section(dashboard, formats)

    # Key Metrics Cards
    _create_key_metrics_section(dashboard, formats)

    # Category Status
    _create_category_status_section(dashboard, formats)

    # Year overview
    _create_year_overview_section(dashboard, formats)

    # Income Tracking
    _create_income_tracking_section(dashboard, formats)

    # Surplus Tracking
    _create_surplus_tracking_section(dashboard, formats)

    # Freelance by Currency
    _create_freelance_currency_section(dashboard, workbook, formats)

    # Dashboard Charts
    _create_dashboard_charts(dashboard, workbook, formats)

    # Selected Month View
    _create_selected_view_section(dashboard, workbook, formats)

    return dashboard


def _create_account_balance_section(dashboard, formats):
    """Create account balance tracking section."""
    dashboard.merge_range('B7:D7', 'ACCOUNT BALANCE TRACKING', formats.header)

    dashboard.write('B8', 'Opening Balance (Jan 1)', formats.text)
    dashboard.write('C8', 0, formats.money_input)
    dashboard.write_comment('C8', 'Enter your account balance on January 1st')

    dashboard.write('B9', 'Current/Closing Balance', formats.text)
    dashboard.write('C9', 0, formats.money_input)
    dashboard.write_comment('C9', 'Update with your current or year-end balance')

    dashboard.write('B10', 'Net Change (YTD)', formats.text)
    dashboard.write_formula('C10', '=C9-C8', formats.money_bold)

    dashboard.write('B11', 'Target Savings (Monthly)', formats.text)
    dashboard.write('C11', '5,000 - 10,000', formats.create_custom_format({
        'align': 'center', 'border': 1, 'bold': True, 'bg_color': '#D5F5E3'
    }))

    dashboard.write('B12', 'Financial Discipline Score', formats.text)
    dashboard.write_formula('C12',
        '=IF(C8=0,"Enter Opening Balance",IF(C10>=10000,"Excellent",IF(C10>=5000,"Good","Needs Attention")))',
        formats.create_custom_format({'align': 'center', 'border': 1, 'bold': True}))

    # Currency indicator
    dashboard.write('I5', 'All values in:', formats.text)
    dashboard.write_formula('J5', '=$N$2', formats.create_custom_format({
        'bold': True, 'font_size': 14, 'align': 'center', 'border': 2,
        'bg_color': '#E8F8F5', 'font_color': '#16A085', 'font_name': UI_FONT
    }))


def _create_key_metrics_section(dashboard, formats):
    """Create key financial metrics section."""
    dashboard.merge_range('B14:G14', 'KEY FINANCIAL METRICS (YEAR-TO-DATE)', formats.header)

    dashboard.write('B15', 'Annual Budget', formats.card_title)
    dashboard.write('C15', 'YTD Spent', formats.card_title)
    dashboard.write('D15', 'Remaining', formats.card_title)
    dashboard.write('E15', 'Budget Used', formats.card_title)
    dashboard.write('F15', 'YTD Savings', formats.card_title)
    dashboard.write('G15', 'Months Active', formats.card_title)

    dashboard.write_formula('B16', "=IFERROR('Monthly Budget'!D14*12,0)", formats.card_value)
    dashboard.write_formula('C16', "=IFERROR('Annual Summary'!E16,0)", formats.card_value)
    dashboard.write_formula('D16', "=IFERROR(B16-C16,0)", formats.card_value_green)
    dashboard.write_formula('E16', "=IFERROR(IF(B16>0,C16/B16,0),0)", formats.card_percent)
    dashboard.write_formula('F16', "=IFERROR(C10,0)", formats.card_value_green)
    dashboard.write_formula('G16', "=IFERROR(COUNTIF('Annual Summary'!E4:E15,\">0\"),0)",
        formats.create_custom_format({
            'bold': True, 'font_size': 18, 'font_color': '#1E3A5F',
            'align': 'center', 'valign': 'vcenter'
        }))

    # Alert Status
    dashboard.merge_range('B18:G18', 'BUDGET ALERT STATUS', formats.header)
    dashboard.merge_range('B19:G20', '', formats.create_custom_format({'border': 1}))
    dashboard.write_formula('B19',
        '=IF(E16>1,"ALERT: You have exceeded your annual budget! Review expenses immediately.",'
        'IF(E16>0.9,"WARNING: You have used over 90% of your annual budget.",'
        'IF(E16>0.75,"CAUTION: You have used over 75% of your annual budget.",'
        '"ON TRACK: Your spending is within budget.")))',
        formats.create_custom_format({
            'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1
        }))


def _create_category_status_section(dashboard, formats):
    """Create category status mini-table."""
    dashboard.merge_range('I7:M7', 'TOP SPENDING CATEGORIES (YTD)', formats.header)
    dashboard.write('I8', 'Category', formats.sub_header)
    dashboard.write('J8', 'Budget', formats.sub_header)
    dashboard.write('K8', 'Spent', formats.sub_header)
    dashboard.write('L8', 'Status', formats.sub_header)
    dashboard.write('M8', '% Used', formats.sub_header)

    for i, cat in enumerate(CATEGORIES[:8]):
        row = 9 + i
        dashboard.write(f'I{row}', cat, formats.get_category_format(cat))
        dashboard.write_formula(f'J{row}', f"='Monthly Budget'!D{i+3}*12", formats.money)
        dashboard.write_formula(f'K{row}', f"='Budget vs Actual'!D{i+3}", formats.money)
        dashboard.write_formula(f'L{row}',
            f'=IF(K{row}>J{row},"Over",IF(K{row}>J{row}*0.9,"Warn","OK"))',
            formats.text_center)
        dashboard.write_formula(f'M{row}', f'=IFERROR(IF(J{row}>0,K{row}/J{row},0),0)', formats.percent)

    # Quick navigation
    dashboard.merge_range('I18:M18', 'MONTHLY EXPENSE SHEETS', formats.header)
    dashboard.merge_range('I19:M20',
        'Each month (Jan-Dec) has pre-populated dates\nNavigate to monthly sheets to log daily expenses',
        formats.create_custom_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_size': 11
        }))


def _create_year_overview_section(dashboard, formats):
    """Create year overview mini-grid."""
    dashboard.merge_range('B22:G22', f'{YEAR} MONTHLY EXPENSE SUMMARY', formats.header)

    # First row of months (Jan-Jun)
    for i, month in enumerate(MONTHS[:6]):
        col = chr(ord('B') + i)
        dashboard.write(f'{col}23', month[:3], formats.get_month_format(month))
        dashboard.write_formula(f'{col}24',
            f"='{month} {YEAR}'!G{DAYS_IN_MONTH[month]+10}", formats.money)

    # Second row of months (Jul-Dec)
    for i, month in enumerate(MONTHS[6:]):
        col = chr(ord('B') + i)
        dashboard.write(f'{col}25', month[:3], formats.get_month_format(month))
        dashboard.write_formula(f'{col}26',
            f"='{month} {YEAR}'!G{DAYS_IN_MONTH[month]+10}", formats.money)


def _create_income_tracking_section(dashboard, formats):
    """Create income tracking section."""
    dashboard.merge_range('I22:M22', 'INCOME TRACKER (Salary + Freelance)', formats.header)

    dashboard.write('I23', 'Month', formats.sub_header)
    dashboard.write('J23', 'Salary', formats.sub_header)
    dashboard.write('K23', 'Freelance', formats.sub_header)
    dashboard.write('L23', 'Total Income', formats.sub_header)
    dashboard.write('M23', 'YTD Freelance', formats.sub_header)

    for i, month in enumerate(MONTHS):
        row = 24 + i
        dashboard.write(f'I{row}', month[:3], formats.get_month_format(month))
        dashboard.write_formula(f'J{row}', f"='{month} {YEAR}'!G5", formats.money)
        dashboard.write_formula(f'K{row}', f"='{month} {YEAR}'!G6", formats.money)
        dashboard.write_formula(f'L{row}', f"='{month} {YEAR}'!G7", formats.money)
        if i == 0:
            dashboard.write_formula(f'M{row}', f'=K{row}', formats.money)
        else:
            dashboard.write_formula(f'M{row}', f'=M{row-1}+K{row}', formats.money)

    # Total row
    dashboard.write('I36', 'TOTAL', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    dashboard.write_formula('J36', '=SUM(J24:J35)', formats.money_bold)
    dashboard.write_formula('K36', '=SUM(K24:K35)', formats.money_bold)
    dashboard.write_formula('L36', '=SUM(L24:L35)', formats.money_bold)
    dashboard.write_formula('M36', '=M35', formats.money_bold)


def _create_surplus_tracking_section(dashboard, formats):
    """Create monthly surplus tracking section."""
    dashboard.merge_range('B28:G28', 'MONTHLY SURPLUS TRACKER (Budget Savings + Freelance)', formats.header)

    dashboard.write('B29', 'Month', formats.sub_header)
    dashboard.write('C29', 'Budget', formats.sub_header)
    dashboard.write('D29', 'Spent', formats.sub_header)
    dashboard.write('E29', 'Savings', formats.sub_header)
    dashboard.write('F29', 'Freelance', formats.sub_header)
    dashboard.write('G29', 'Total Surplus', formats.sub_header)

    for i, month in enumerate(MONTHS):
        row = 30 + i
        dashboard.write(f'B{row}', month[:3], formats.get_month_format(month))
        dashboard.write_formula(f'C{row}', "='Monthly Budget'!D14", formats.money)
        dashboard.write_formula(f'D{row}', f"='{month} {YEAR}'!G{DAYS_IN_MONTH[month]+10}", formats.money)
        dashboard.write_formula(f'E{row}', f'=C{row}-D{row}', formats.money)
        dashboard.write_formula(f'F{row}', f"='{month} {YEAR}'!G6", formats.money)
        dashboard.write_formula(f'G{row}', f'=E{row}+F{row}', formats.money)

    # Total row
    dashboard.write('B42', 'TOTAL', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    dashboard.write_formula('C42', '=SUM(C30:C41)', formats.money_bold)
    dashboard.write_formula('D42', '=SUM(D30:D41)', formats.money_bold)
    dashboard.write_formula('E42', '=SUM(E30:E41)', formats.money_bold)
    dashboard.write_formula('F42', '=SUM(F30:F41)', formats.money_bold)
    dashboard.write_formula('G42', '=SUM(G30:G41)', formats.money_bold)

    # Cumulative Balance Summary
    dashboard.merge_range('I38:M38', 'CUMULATIVE BALANCE SUMMARY', formats.header)
    dashboard.write('I39', 'Total Salary (YTD):', formats.text)
    dashboard.write_formula('L39', '=J36', formats.money_bold)
    dashboard.write('I40', 'Total Freelance (YTD):', formats.text)
    dashboard.write_formula('L40', '=K36', formats.money_bold)
    dashboard.write('I41', 'Total Income (YTD):', formats.text)
    dashboard.write_formula('L41', '=L36', formats.money_bold)
    dashboard.write('I42', 'Total Expenses (YTD):', formats.text)
    dashboard.write_formula('L42', '=D42', formats.money_bold)
    dashboard.write('I43', 'Net Surplus (YTD):', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    dashboard.write_formula('L43', '=G42', formats.create_custom_format({
        'num_format': '#,##0 CZK', 'align': 'right', 'border': 1,
        'bold': True, 'bg_color': '#27AE60', 'font_color': 'white', 'font_size': 14
    }))


def _create_freelance_currency_section(dashboard, workbook, formats):
    """Create freelance earnings by currency section."""
    freelance_header = formats.create_custom_format({
        'bold': True, 'font_size': 12, 'font_color': 'white',
        'bg_color': '#16A085', 'align': 'center', 'valign': 'vcenter',
        'border': 2, 'font_name': UI_FONT
    })

    dashboard.merge_range('O2:S2', 'FREELANCE EARNINGS BY CURRENCY', freelance_header)

    dashboard.write('O3', 'Currency', formats.sub_header)
    dashboard.write('P3', 'Total Earned', formats.sub_header)
    dashboard.write('Q3', 'Months', formats.sub_header)
    dashboard.write('R3', 'Avg/Month', formats.sub_header)
    dashboard.write('S3', '% of Total', formats.sub_header)

    for i, currency in enumerate(CURRENCY_OPTIONS):
        row = 4 + i
        dashboard.write(f'O{row}', currency, formats.create_custom_format({
            'bold': True, 'align': 'center', 'border': 2,
            'bg_color': CURRENCY_COLORS[currency], 'font_color': 'white', 'font_name': UI_FONT
        }))

        # Build formulas
        sumif_parts = []
        countif_parts = []
        for month in MONTHS:
            sumif_parts.append(f"IF('{month} {YEAR}'!I6=\"{currency}\",'{month} {YEAR}'!G6,0)")
            countif_parts.append(f"IF('{month} {YEAR}'!I6=\"{currency}\",1,0)")

        dashboard.write_formula(f'P{row}', f'={"+".join(sumif_parts)}', formats.money)
        dashboard.write_formula(f'Q{row}', f'={"+".join(countif_parts)}', formats.text_center)
        dashboard.write_formula(f'R{row}', f'=IFERROR(IF(Q{row}>0,P{row}/Q{row},0),0)', formats.money)
        dashboard.write_formula(f'S{row}', f'=IFERROR(IF(P8>0,P{row}/P8,0),0)', formats.percent)

    # Total row
    dashboard.write('O8', 'TOTAL', formats.create_custom_format({
        'bold': True, 'align': 'center', 'border': 2, 'bg_color': '#D4E6F1', 'font_name': UI_FONT
    }))
    dashboard.write_formula('P8', '=SUM(P4:P7)', formats.money_bold)
    dashboard.write_formula('Q8', '=SUM(Q4:Q7)', formats.create_custom_format({
        'bold': True, 'align': 'center', 'border': 2, 'bg_color': '#D4E6F1', 'font_name': UI_FONT
    }))
    dashboard.write_formula('R8', '=IFERROR(IF(Q8>0,P8/Q8,0),0)', formats.money_bold)
    dashboard.write_formula('S8', '=IFERROR(SUM(S4:S7),0)', formats.percent_bold)


def _create_dashboard_charts(dashboard, workbook, formats):
    """Create embedded dashboard charts."""
    # Monthly Expenses Trend Chart
    trend_chart = workbook.add_chart({'type': 'line'})
    trend_chart.add_series({
        'name': 'Expenses',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$E$4:$E$15",
        'line': {'color': '#E74C3C', 'width': 2.5},
        'marker': {'type': 'circle', 'size': 6, 'fill': {'color': '#E74C3C'}}
    })
    trend_chart.add_series({
        'name': 'Budget',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$F$4:$F$15",
        'line': {'color': '#3498DB', 'width': 2, 'dash_type': 'dash'},
        'marker': {'type': 'none'}
    })
    trend_chart.set_title({'name': 'Monthly Expense Trend', 'name_font': {'size': 11, 'bold': True}})
    trend_chart.set_x_axis({'num_font': {'size': 8, 'rotation': -45}})
    trend_chart.set_y_axis({'num_font': {'size': 8}})
    trend_chart.set_legend({'position': 'bottom', 'font': {'size': 8}})
    trend_chart.set_size({'width': 450, 'height': 250})
    dashboard.insert_chart('O10', trend_chart)

    # Budget Allocation Pie Chart
    pie_chart = workbook.add_chart({'type': 'pie'})
    pie_chart.add_series({
        'name': 'Budget Allocation',
        'categories': f"='Monthly Budget'!$C$3:$C${len(CATEGORIES)+2}",
        'values': f"='Monthly Budget'!$D$3:$D${len(CATEGORIES)+2}",
        'data_labels': {'percentage': True, 'font': {'size': 7}},
    })
    pie_chart.set_title({'name': 'Budget Allocation', 'name_font': {'size': 11, 'bold': True}})
    pie_chart.set_legend({'position': 'right', 'font': {'size': 7}})
    pie_chart.set_size({'width': 400, 'height': 250})
    dashboard.insert_chart('O24', pie_chart)

    # Surplus/Deficit Bar Chart
    surplus_chart = workbook.add_chart({'type': 'column'})
    surplus_chart.add_series({
        'name': 'Surplus/Deficit',
        'categories': "='Annual Summary'!$B$4:$B$15",
        'values': "='Annual Summary'!$G$4:$G$15",
        'fill': {'color': '#27AE60'},
    })
    surplus_chart.set_title({'name': 'Monthly Surplus/Deficit', 'name_font': {'size': 11, 'bold': True}})
    surplus_chart.set_x_axis({'num_font': {'size': 8, 'rotation': -45}})
    surplus_chart.set_y_axis({'num_font': {'size': 8}})
    surplus_chart.set_legend({'none': True})
    surplus_chart.set_size({'width': 450, 'height': 250})
    dashboard.insert_chart('O38', surplus_chart)


def _create_selected_view_section(dashboard, workbook, formats):
    """Create selected month view section."""
    view_header = formats.create_custom_format({
        'bold': True, 'font_size': 12, 'font_color': 'white',
        'bg_color': '#9B59B6', 'align': 'center', 'valign': 'vcenter',
        'border': 1, 'text_wrap': True
    })

    dashboard.merge_range('B45:G45', 'SELECTED VIEW CATEGORY BREAKDOWN', view_header)
    dashboard.write('B46', 'Category', formats.sub_header)
    dashboard.write('C46', 'Monthly Budget', formats.sub_header)
    dashboard.write('D46', 'Selected Spent', formats.sub_header)
    dashboard.write('E46', 'Remaining', formats.sub_header)
    dashboard.write('F46', '% Used', formats.sub_header)
    dashboard.write('G46', 'Status', formats.sub_header)

    for i, cat in enumerate(CATEGORIES):
        row = 47 + i
        dashboard.write(f'B{row}', cat, formats.get_category_format(cat))
        dashboard.write_formula(f'C{row}', f"='Monthly Budget'!D{i+3}", formats.money)

        # Dynamic spent formula
        spent_parts = []
        for month in MONTHS:
            spent_parts.append(f'IF($K$2="{month}",\'{month} {YEAR}\'!M{i+6},')
        all_time_sum = '+'.join([f"'{m} {YEAR}'!M{i+6}" for m in MONTHS])
        spent_formula = ''.join(spent_parts) + f'IF($K$2="ALL TIME",{all_time_sum}/12,0)' + ')' * 12
        dashboard.write_formula(f'D{row}', f'={spent_formula}', formats.money)

        dashboard.write_formula(f'E{row}', f'=C{row}-D{row}', formats.money)
        dashboard.write_formula(f'F{row}', f'=IFERROR(IF(C{row}>0,D{row}/C{row},0),0)', formats.percent)
        dashboard.write_formula(f'G{row}',
            f'=IF(D{row}>C{row},"OVER",IF(D{row}>C{row}*0.9,"WARN","OK"))',
            formats.text_center)

    # Conditional formatting
    status_start = 47
    status_end = 47 + len(CATEGORIES) - 1
    dashboard.conditional_format(f'G{status_start}:G{status_end}', {
        'type': 'text', 'criteria': 'containing', 'value': 'OVER', 'format': formats.danger
    })
    dashboard.conditional_format(f'G{status_start}:G{status_end}', {
        'type': 'text', 'criteria': 'containing', 'value': 'WARN', 'format': formats.warning
    })
    dashboard.conditional_format(f'G{status_start}:G{status_end}', {
        'type': 'text', 'criteria': 'containing', 'value': 'OK', 'format': formats.good
    })

    # Total row
    total_row = 47 + len(CATEGORIES)
    dashboard.write(f'B{total_row}', 'TOTAL', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    dashboard.write_formula(f'C{total_row}', f'=SUM(C47:C{total_row-1})', formats.money_bold)
    dashboard.write_formula(f'D{total_row}', f'=SUM(D47:D{total_row-1})', formats.money_bold)
    dashboard.write_formula(f'E{total_row}', f'=SUM(E47:E{total_row-1})', formats.money_bold)
    dashboard.write_formula(f'F{total_row}',
        f'=IFERROR(IF(C{total_row}>0,D{total_row}/C{total_row},0),0)', formats.percent_bold)

    # Selected View Summary
    dashboard.merge_range('I45:M45', 'SELECTED VIEW SUMMARY', view_header)
    dashboard.write('I46', 'Selected Period:', formats.text)
    dashboard.write_formula('L46', '=$K$2', formats.create_custom_format({
        'bold': True, 'font_size': 14, 'align': 'center', 'border': 1, 'bg_color': '#F5EEF8'
    }))
    dashboard.write('I47', 'Total Budget:', formats.text)
    dashboard.write_formula('L47', f'=C{total_row}', formats.money_bold)
    dashboard.write('I48', 'Total Spent:', formats.text)
    dashboard.write_formula('L48', f'=D{total_row}', formats.money_bold)
    dashboard.write('I49', 'Net Balance:', formats.text)
    dashboard.write_formula('L49', f'=E{total_row}', formats.create_custom_format({
        'num_format': '#,##0 CZK', 'align': 'right', 'border': 1,
        'bold': True, 'bg_color': '#D5F5E3', 'font_size': 14
    }))
    dashboard.write('I50', 'Budget Utilization:', formats.text)
    dashboard.write_formula('L50', f'=F{total_row}', formats.create_custom_format({
        'num_format': '0.0%', 'align': 'center', 'border': 1, 'bold': True, 'font_size': 14
    }))

    # Freelance for selected period
    dashboard.write('I51', 'Freelance Income:', formats.text)
    freelance_parts = []
    for month in MONTHS:
        freelance_parts.append(f'IF($K$2="{month}",\'{month} {YEAR}\'!G6,')
    all_freelance = '+'.join([f"'{m} {YEAR}'!G6" for m in MONTHS])
    freelance_formula = ''.join(freelance_parts) + f'IF($K$2="ALL TIME",{all_freelance},0)' + ')' * 12
    dashboard.write_formula('L51', f'={freelance_formula}', formats.money_bold)

    dashboard.write('I52', 'Total Income:', formats.text)
    income_parts = []
    for month in MONTHS:
        income_parts.append(f'IF($K$2="{month}",\'{month} {YEAR}\'!G7,')
    all_income = '+'.join([f"'{m} {YEAR}'!G7" for m in MONTHS])
    income_formula = ''.join(income_parts) + f'IF($K$2="ALL TIME",{all_income},0)' + ')' * 12
    dashboard.write_formula('L52', f'={income_formula}', formats.create_custom_format({
        'num_format': '#,##0 CZK', 'align': 'right', 'border': 1,
        'bold': True, 'bg_color': '#D4E6F1', 'font_size': 12
    }))
