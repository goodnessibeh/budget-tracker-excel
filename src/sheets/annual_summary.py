"""Annual Summary sheet creation module."""

from ..config import MONTHS, DAYS_IN_MONTH, YEAR


def create_annual_summary(workbook, formats):
    """Create the annual summary sheet."""
    sheet = workbook.add_worksheet('Annual Summary')
    sheet.set_tab_color('#1ABC9C')

    # Column widths
    sheet.set_column('A:A', 5)
    sheet.set_column('B:B', 16)
    sheet.set_column('C:C', 20)
    sheet.set_column('D:D', 20)
    sheet.set_column('E:E', 20)
    sheet.set_column('F:F', 20)
    sheet.set_column('G:G', 20)
    sheet.set_column('H:H', 16)
    sheet.set_column('I:I', 12)

    # Title
    sheet.merge_range('B1:I1', f'{YEAR} ANNUAL FINANCIAL SUMMARY', formats.title)
    sheet.merge_range('B2:I2', 'January 1st - December 31st', formats.subtitle)

    # Headers
    sheet.write('B3', 'Month', formats.header)
    sheet.write('C3', 'Opening Bal.', formats.header)
    sheet.write('D3', 'Closing Bal.', formats.header)
    sheet.write('E3', 'Expenses', formats.header)
    sheet.write('F3', 'Budget', formats.header)
    sheet.write('G3', 'Variance', formats.header)
    sheet.write('H3', 'Status', formats.header)
    sheet.write('I3', 'Days', formats.header)

    # Monthly rows
    for i, month in enumerate(MONTHS):
        row = i + 4
        num_days = DAYS_IN_MONTH[month]
        sheet.write(f'B{row}', month, formats.get_month_format(month))
        sheet.write_formula(f'C{row}', f"='{month} {YEAR}'!C5", formats.money)
        sheet.write_formula(f'D{row}', f"='{month} {YEAR}'!C6", formats.money)
        sheet.write_formula(f'E{row}', f"='{month} {YEAR}'!G{num_days+10}", formats.money)
        sheet.write_formula(f'F{row}', "='Monthly Budget'!D14", formats.money)
        sheet.write_formula(f'G{row}', f'=F{row}-E{row}', formats.money)
        sheet.write_formula(f'H{row}',
            f'=IF(E{row}>F{row},"OVER",IF(E{row}>F{row}*0.9,"WARN","OK"))',
            formats.text_center)
        sheet.write(f'I{row}', num_days, formats.text_center)

        # Conditional formatting
        sheet.conditional_format(f'H{row}', {
            'type': 'text', 'criteria': 'containing', 'value': 'OVER',
            'format': formats.danger
        })
        sheet.conditional_format(f'H{row}', {
            'type': 'text', 'criteria': 'containing', 'value': 'WARN',
            'format': formats.warning
        })
        sheet.conditional_format(f'H{row}', {
            'type': 'text', 'criteria': 'containing', 'value': 'OK',
            'format': formats.good
        })

        # Variance coloring
        sheet.conditional_format(f'G{row}', {
            'type': 'cell', 'criteria': '<', 'value': 0,
            'format': formats.create_custom_format({
                'num_format': '#,##0 CZK', 'bg_color': '#FADBD8', 'border': 1
            })
        })
        sheet.conditional_format(f'G{row}', {
            'type': 'cell', 'criteria': '>=', 'value': 0,
            'format': formats.create_custom_format({
                'num_format': '#,##0 CZK', 'bg_color': '#D5F5E3', 'border': 1
            })
        })

    # Annual totals
    sheet.write('B16', 'YEAR TOTAL', formats.create_custom_format({
        'bold': True, 'align': 'left', 'border': 1, 'bg_color': '#D4E6F1'
    }))
    sheet.write_formula('C16', '=C4', formats.money_bold)
    sheet.write_formula('D16', '=D15', formats.money_bold)
    sheet.write_formula('E16', '=SUM(E4:E15)', formats.money_bold)
    sheet.write_formula('F16', '=SUM(F4:F15)', formats.money_bold)
    sheet.write_formula('G16', '=SUM(G4:G15)', formats.money_bold)
    sheet.write('I16', 365, formats.create_custom_format({
        'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D4E6F1'
    }))

    # Freeze headers
    sheet.freeze_panes(3, 1)

    return sheet
