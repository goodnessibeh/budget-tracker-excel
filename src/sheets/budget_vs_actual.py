"""Budget vs Actual sheet creation module."""

from ..config import CATEGORIES, MONTHS, YEAR


def create_budget_vs_actual(workbook, formats):
    """Create the budget vs actual comparison sheet."""
    sheet = workbook.add_worksheet('Budget vs Actual')
    sheet.set_tab_color('#F39C12')

    # Column widths
    sheet.set_column('A:A', 5)
    sheet.set_column('B:B', 32)
    sheet.set_column('C:C', 20)
    sheet.set_column('D:D', 20)
    sheet.set_column('E:E', 20)
    sheet.set_column('F:F', 16)
    sheet.set_column('G:G', 16)
    sheet.set_column('H:H', 18)

    # Title
    sheet.merge_range('B1:H1', f'BUDGET VS ACTUAL - YEAR {YEAR}', formats.title)
    sheet.merge_range('B2:H2', 'January 1st - December 31st', formats.subtitle)

    # Headers
    sheet.write('B3', 'Category', formats.header)
    sheet.write('C3', 'Annual Budget', formats.header)
    sheet.write('D3', 'YTD Actual', formats.header)
    sheet.write('E3', 'Remaining', formats.header)
    sheet.write('F3', '% Used', formats.header)
    sheet.write('G3', '% of Spend', formats.header)
    sheet.write('H3', 'Status', formats.header)

    # Category rows
    for i, cat in enumerate(CATEGORIES):
        row = i + 4
        sheet.write(f'B{row}', cat, formats.get_category_format(cat))
        sheet.write_formula(f'C{row}', f"='Monthly Budget'!D{i+3}*12", formats.money)

        # Sum from all monthly sheets
        sum_formula = '+'.join([f"'{m} {YEAR}'!M{i+6}" for m in MONTHS])
        sheet.write_formula(f'D{row}', f'={sum_formula}', formats.money)

        sheet.write_formula(f'E{row}', f'=C{row}-D{row}', formats.money)
        sheet.write_formula(f'F{row}', f'=IFERROR(IF(C{row}>0,D{row}/C{row},0),0)', formats.percent)
        sheet.write_formula(f'G{row}', f'=IFERROR(IF($D$15>0,D{row}/$D$15,0),0)', formats.percent)
        sheet.write_formula(f'H{row}',
            f'=IF(D{row}>C{row},"OVER BUDGET",IF(D{row}>C{row}*0.9,"WARNING",'
            f'IF(D{row}>C{row}*0.75,"CAUTION","ON TRACK")))',
            formats.text_center)

    # Conditional formatting
    total_row = len(CATEGORIES) + 4
    sheet.conditional_format(f'H4:H{total_row-1}', {
        'type': 'text', 'criteria': 'containing', 'value': 'OVER',
        'format': formats.danger
    })
    sheet.conditional_format(f'H4:H{total_row-1}', {
        'type': 'text', 'criteria': 'containing', 'value': 'WARNING',
        'format': formats.warning
    })
    sheet.conditional_format(f'H4:H{total_row-1}', {
        'type': 'text', 'criteria': 'containing', 'value': 'CAUTION',
        'format': formats.create_custom_format({
            'bg_color': '#F7DC6F', 'align': 'center', 'border': 1
        })
    })
    sheet.conditional_format(f'H4:H{total_row-1}', {
        'type': 'text', 'criteria': 'containing', 'value': 'ON TRACK',
        'format': formats.good
    })

    # Remaining column formatting
    sheet.conditional_format(f'E4:E{total_row-1}', {
        'type': 'cell', 'criteria': '<', 'value': 0,
        'format': formats.danger
    })

    # Progress bars
    sheet.conditional_format(f'F4:F{total_row-1}', {
        'type': 'data_bar', 'bar_color': '#3498DB', 'bar_solid': True,
        'min_value': 0, 'max_value': 1
    })

    # Total row
    sheet.write(f'B{total_row}', 'TOTAL', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    sheet.write_formula(f'C{total_row}', f'=SUM(C4:C{total_row-1})', formats.money_bold)
    sheet.write_formula(f'D{total_row}', f'=SUM(D4:D{total_row-1})', formats.money_bold)
    sheet.write_formula(f'E{total_row}', f'=SUM(E4:E{total_row-1})', formats.money_bold)
    sheet.write_formula(f'F{total_row}',
        f'=IFERROR(IF(C{total_row}>0,D{total_row}/C{total_row},0),0)', formats.percent_bold)

    # Freeze headers
    sheet.freeze_panes(3, 1)

    return sheet
