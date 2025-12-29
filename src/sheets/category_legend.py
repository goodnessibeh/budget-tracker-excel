"""Category Legend sheet creation module."""

from ..config import CATEGORIES, MONTHS, DAYS_IN_MONTH, YEAR


def create_category_legend(workbook, formats):
    """Create the category legend sheet."""
    sheet = workbook.add_worksheet('Category Legend')
    sheet.set_tab_color('#95A5A6')

    # Column widths
    sheet.set_column('A:A', 5)
    sheet.set_column('B:B', 30)
    sheet.set_column('C:C', 15)
    sheet.set_column('D:D', 50)

    # Title
    sheet.merge_range('B1:D1', 'CATEGORY & COLOR LEGEND', formats.title)
    sheet.set_row(0, 40)

    # Category headers
    sheet.write('B2', 'Category', formats.header)
    sheet.write('C2', 'Color', formats.header)
    sheet.write('D2', 'Description', formats.header)

    # Category descriptions
    descriptions = {
        'Rent + Utilities': 'Monthly rent and utility bills (water, gas, heating)',
        'Electricity Deposit': 'Electricity deposit/prepayment',
        'Internet & Phone Bills': 'Internet and phone service bills',
        'Food & Groceries': 'Grocery shopping and food expenses',
        'Investments': 'Savings and investment contributions',
        'Date Nights': 'Entertainment and date activities',
        'Annual Costs Savings': 'Savings for annual expenses (insurance, subscriptions)',
        'Emergency Buffer': 'Emergency fund contributions',
        'Personal Care': 'Haircuts, grooming, and personal items',
        'Personal Savings': 'Personal savings and future goals',
        'Miscellaneous': 'Other uncategorized expenses'
    }

    for i, cat in enumerate(CATEGORIES):
        row = i + 3
        sheet.write(f'B{row}', cat, formats.text)
        sheet.write(f'C{row}', '', formats.get_category_format(cat))
        sheet.write(f'D{row}', descriptions.get(cat, ''), formats.text)

    # Month colors section
    sheet.merge_range('B21:D21', 'MONTH COLOR CODING', formats.header)
    sheet.write('B22', 'Month', formats.sub_header)
    sheet.write('C22', 'Color', formats.sub_header)
    sheet.write('D22', 'Date Range', formats.sub_header)

    for i, month in enumerate(MONTHS):
        row = i + 23
        num_days = DAYS_IN_MONTH[month]
        sheet.write(f'B{row}', month, formats.text)
        sheet.write(f'C{row}', '', formats.get_month_format(month))
        sheet.write(f'D{row}', f'{month} 1 - {month} {num_days}, {YEAR}', formats.text)

    # Freeze headers
    sheet.freeze_panes(2, 0)

    return sheet
