"""Monthly Budget sheet creation module."""

from ..config import CATEGORIES, BUDGETS, UI_FONT


def create_monthly_budget(workbook, formats):
    """Create the monthly budget allocation sheet."""
    sheet = workbook.add_worksheet('Monthly Budget')
    sheet.set_tab_color('#27AE60')

    # Column widths
    sheet.set_column('A:A', 5)
    sheet.set_column('B:B', 6)
    sheet.set_column('C:C', 32)
    sheet.set_column('D:D', 22)
    sheet.set_column('E:E', 22)
    sheet.set_column('F:F', 18)
    sheet.set_column('G:G', 28)

    # Title
    sheet.merge_range('B1:G1', "MONTHLY BUDGET", formats.title)
    sheet.set_row(0, 40)

    # Headers
    sheet.write('B2', '#', formats.header)
    sheet.write('C2', 'Category', formats.header)
    sheet.write('D2', 'Monthly', formats.header)
    sheet.write('E2', 'Annual', formats.header)
    sheet.write('F2', '% of Total', formats.header)
    sheet.write('G2', 'Color Code', formats.header)

    # Category rows
    for i, (cat, amount) in enumerate(zip(CATEGORIES, BUDGETS)):
        row = i + 3
        sheet.write(f'B{row}', i + 1, formats.text_center)
        sheet.write(f'C{row}', cat, formats.text)
        sheet.write(f'D{row}', amount, formats.money_input)
        sheet.write_formula(f'E{row}', f'=D{row}*12', formats.money)
        sheet.write_formula(f'F{row}', f'=IFERROR(IF($D$14>0,D{row}/$D$14,0),0)', formats.percent)
        sheet.write(f'G{row}', cat, formats.get_category_format(cat))

    # Total row
    total_row = len(CATEGORIES) + 3
    sheet.write(f'B{total_row}', '', formats.money_bold)
    sheet.write(f'C{total_row}', 'TOTAL MONTHLY BUDGET', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    sheet.write_formula(f'D{total_row}', f'=SUM(D3:D{total_row-1})', formats.money_bold)
    sheet.write_formula(f'E{total_row}', f'=SUM(E3:E{total_row-1})', formats.money_bold)
    sheet.write_formula(f'F{total_row}', f'=SUM(F3:F{total_row-1})', formats.percent_bold)

    # Income section
    sheet.merge_range(f'B{total_row+2}:G{total_row+2}', 'INCOME INFORMATION', formats.header)
    sheet.write(f'C{total_row+3}', 'Combined Monthly Income', formats.text)
    sheet.write(f'D{total_row+3}', 60000, formats.money_input)
    sheet.write_formula(f'E{total_row+3}', f'=D{total_row+3}*12', formats.money)

    sheet.write(f'C{total_row+4}', 'Target Monthly Savings', formats.text)
    sheet.write(f'D{total_row+4}', '5,000 - 10,000', formats.create_custom_format({
        'align': 'right', 'border': 1, 'bold': True, 'bg_color': '#D5F5E3'
    }))
    sheet.write(f'E{total_row+4}', '60,000 - 120,000', formats.create_custom_format({
        'align': 'right', 'border': 1, 'bold': True, 'bg_color': '#D5F5E3'
    }))

    sheet.write(f'C{total_row+5}', 'Target Savings Rate', formats.text)
    sheet.write(f'D{total_row+5}', '8% - 17%', formats.create_custom_format({
        'align': 'center', 'border': 1, 'bold': True, 'bg_color': '#D5F5E3'
    }))

    # Freeze headers
    sheet.freeze_panes(2, 0)

    return sheet
