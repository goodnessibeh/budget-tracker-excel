"""Monthly expense sheets creation module."""

from datetime import date
from ..config import (
    CATEGORIES, MONTHS, DAYS_IN_MONTH, MONTH_COLORS,
    CATEGORY_COLORS, CURRENCY_OPTIONS, PAYMENT_METHODS, UI_FONT, YEAR
)


def create_monthly_expense_sheets(workbook, formats):
    """Create all 12 monthly expense tracking sheets."""
    sheets = []

    for month_idx, month in enumerate(MONTHS):
        sheet = _create_monthly_sheet(workbook, formats, month, month_idx)
        sheets.append(sheet)

    return sheets


def _create_monthly_sheet(workbook, formats, month, month_idx):
    """Create a single monthly expense sheet."""
    sheet_name = f'{month} {YEAR}'
    sheet = workbook.add_worksheet(sheet_name)
    sheet.set_tab_color(MONTH_COLORS[month])

    num_days = DAYS_IN_MONTH[month]
    month_num = month_idx + 1

    # Column widths
    sheet.set_column('A:A', 5)
    sheet.set_column('B:B', 6)
    sheet.set_column('C:C', 14)
    sheet.set_column('D:D', 12)
    sheet.set_column('E:E', 28)
    sheet.set_column('F:F', 38)
    sheet.set_column('G:G', 20)
    sheet.set_column('H:H', 16)
    sheet.set_column('I:I', 20)
    sheet.set_column('K:K', 28)
    sheet.set_column('L:L', 18)
    sheet.set_column('M:M', 18)
    sheet.set_column('N:N', 14)

    # Month header
    month_title = formats.create_custom_format({
        'bold': True, 'font_size': 24, 'font_color': 'white',
        'bg_color': MONTH_COLORS[month], 'align': 'center', 'valign': 'vcenter'
    })

    first_day = date(YEAR, month_num, 1)
    last_day = date(YEAR, month_num, num_days)

    sheet.merge_range('B1:I1', f'{month.upper()} {YEAR} EXPENSE TRACKER', month_title)
    sheet.set_row(0, 45)

    # Date range subtitle
    sheet.merge_range('B2:I2',
        f'{first_day.strftime("%B %d")} - {last_day.strftime("%B %d, %Y")} ({num_days} days)',
        formats.create_custom_format({
            'bold': True, 'font_size': 12, 'align': 'center',
            'bg_color': MONTH_COLORS[month], 'font_color': 'white'
        }))

    # Account balance section
    sheet.merge_range('B4:D4', 'MONTH BALANCE & INCOME', formats.get_month_header_format(month))
    sheet.write('B5', 'Opening Balance:', formats.text)
    sheet.write('C5', 0, formats.money_input)
    sheet.write('B6', 'Closing Balance:', formats.text)
    sheet.write('C6', 0, formats.money_input)
    sheet.write('B7', 'Net Change:', formats.text)
    sheet.write_formula('C7', '=C6-C5', formats.money_bold)

    # Income section
    sheet.merge_range('F4:I4', 'INCOME TRACKING', formats.get_month_header_format(month))
    sheet.write('F5', 'Salary (Fixed):', formats.text)
    sheet.write('G5', 60000, formats.money)
    sheet.write('H5', 'Currency:', formats.text)
    sheet.write_formula('I5', "=Dashboard!$N$2", formats.create_custom_format({
        'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#E8F8F5',
        'font_color': '#16A085', 'font_name': UI_FONT
    }))

    sheet.write('F6', 'Freelance Income:', formats.text)
    sheet.write('G6', 0, formats.money_input)
    sheet.write_comment('G6', 'Enter freelance/side income earned this month')
    sheet.write('H6', 'Currency:', formats.text)
    sheet.write('I6', 'CZK', formats.create_custom_format({
        'bold': True, 'align': 'center', 'border': 2, 'bg_color': '#E8F8F5',
        'font_color': '#16A085', 'font_name': UI_FONT, 'locked': False
    }))
    sheet.write_comment('I6', 'Select the currency for this freelance income')
    sheet.data_validation('I6', {
        'validate': 'list',
        'source': CURRENCY_OPTIONS,
        'dropdown': True
    })

    sheet.write('F7', 'Total Income:', formats.text)
    sheet.write_formula('G7', '=G5+G6', formats.money_bold)

    # Expense entry headers
    sheet.write('B9', '#', formats.get_month_header_format(month))
    sheet.write('C9', 'Date', formats.get_month_header_format(month))
    sheet.write('D9', 'Day', formats.get_month_header_format(month))
    sheet.write('E9', 'Category', formats.get_month_header_format(month))
    sheet.write('F9', 'Description', formats.get_month_header_format(month))
    sheet.write('G9', 'Amount', formats.get_month_header_format(month))
    sheet.write('H9', 'Payment', formats.get_month_header_format(month))
    sheet.write('I9', 'Running Total', formats.get_month_header_format(month))

    # Date formats
    weekend_date = formats.create_custom_format({
        'num_format': 'dd-mmm', 'align': 'center', 'border': 1, 'bg_color': '#FFE5E5'
    })
    weekday_date = formats.create_custom_format({
        'num_format': 'dd-mmm', 'align': 'center', 'border': 1, 'bg_color': '#F5F5F5'
    })
    day_fmt = formats.create_custom_format({
        'align': 'center', 'border': 1, 'bg_color': '#F5F5F5'
    })
    weekend_day = formats.create_custom_format({
        'align': 'center', 'border': 1, 'bg_color': '#FFE5E5', 'font_color': '#C0392B'
    })

    # Pre-populate dates
    for day in range(1, num_days + 1):
        row = day + 9
        current_date = date(YEAR, month_num, day)
        day_name = current_date.strftime('%a')
        is_weekend = current_date.weekday() >= 5

        sheet.write(f'B{row}', day, formats.text_center)
        sheet.write_datetime(f'C{row}', current_date, weekend_date if is_weekend else weekday_date)
        sheet.write(f'D{row}', day_name, weekend_day if is_weekend else day_fmt)
        sheet.write(f'E{row}', '', formats.text_input)
        sheet.write(f'F{row}', '', formats.text_input)
        sheet.write(f'G{row}', '', formats.money_input)
        sheet.write(f'H{row}', '', formats.text_input)

        if day == 1:
            sheet.write_formula(f'I{row}', f'=G{row}', formats.money)
        else:
            sheet.write_formula(f'I{row}', f'=I{row-1}+G{row}', formats.money)

        # Data validation
        sheet.data_validation(f'E{row}', {
            'validate': 'list',
            'source': CATEGORIES,
            'dropdown': True
        })
        sheet.data_validation(f'H{row}', {
            'validate': 'list',
            'source': PAYMENT_METHODS,
            'dropdown': True
        })

    # Conditional formatting for categories
    last_data_row = num_days + 9
    for cat, color in CATEGORY_COLORS.items():
        sheet.conditional_format(f'E10:E{last_data_row}', {
            'type': 'text',
            'criteria': 'containing',
            'value': cat,
            'format': formats.create_custom_format({
                'bg_color': color, 'font_color': 'white', 'bold': True, 'border': 1
            })
        })

    # Total row
    total_row = num_days + 10
    sheet.write(f'F{total_row}', 'MONTH TOTAL:', formats.create_custom_format({
        'bold': True, 'align': 'right', 'border': 1,
        'bg_color': MONTH_COLORS[month], 'font_color': 'white'
    }))
    sheet.write_formula(f'G{total_row}', f'=SUM(G10:G{last_data_row})', formats.create_custom_format({
        'num_format': '#,##0 CZK', 'align': 'right', 'border': 1,
        'bold': True, 'bg_color': MONTH_COLORS[month], 'font_color': 'white'
    }))

    # Category summary
    sheet.merge_range('K4:N4', f'{month} CATEGORY BREAKDOWN', formats.get_month_header_format(month))
    sheet.write('K5', 'Category', formats.sub_header)
    sheet.write('L5', 'Budget', formats.sub_header)
    sheet.write('M5', 'Spent', formats.sub_header)
    sheet.write('N5', 'Status', formats.sub_header)

    for i, cat in enumerate(CATEGORIES):
        row = i + 6
        sheet.write(f'K{row}', cat, formats.get_category_format(cat))
        sheet.write_formula(f'L{row}', f"='Monthly Budget'!D{i+3}", formats.money)
        sheet.write_formula(f'M{row}', f'=SUMIF(E:E,K{row},G:G)', formats.money)
        sheet.write_formula(f'N{row}',
            f'=IF(M{row}>L{row},"OVER",IF(M{row}>L{row}*0.9,"WARN","OK"))',
            formats.text_center)

        # Conditional formatting
        sheet.conditional_format(f'N{row}', {
            'type': 'text', 'criteria': 'containing', 'value': 'OVER',
            'format': formats.danger
        })
        sheet.conditional_format(f'N{row}', {
            'type': 'text', 'criteria': 'containing', 'value': 'WARN',
            'format': formats.warning
        })
        sheet.conditional_format(f'N{row}', {
            'type': 'text', 'criteria': 'containing', 'value': 'OK',
            'format': formats.good
        })

    # Category totals
    cat_total_row = len(CATEGORIES) + 6
    sheet.write(f'K{cat_total_row}', 'TOTAL', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    sheet.write_formula(f'L{cat_total_row}', f'=SUM(L6:L{cat_total_row-1})', formats.money_bold)
    sheet.write_formula(f'M{cat_total_row}', f'=SUM(M6:M{cat_total_row-1})', formats.money_bold)

    # Freeze panes
    sheet.freeze_panes(9, 1)

    return sheet
