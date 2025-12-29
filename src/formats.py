"""
Excel format definitions for the Budget Tracker.
All formats are created once and reused across sheets.
"""

from .config import UI_FONT, CATEGORY_COLORS, MONTH_COLORS


class FormatManager:
    """Manages all Excel formats for the workbook."""

    def __init__(self, workbook):
        """Initialize formats for the given workbook."""
        self.workbook = workbook
        self._create_formats()

    def _create_formats(self):
        """Create all format definitions."""
        wb = self.workbook
        font = UI_FONT

        # Header formats
        self.title = wb.add_format({
            'bold': True, 'font_size': 24, 'font_color': '#1E3A5F',
            'align': 'center', 'valign': 'vcenter', 'font_name': font
        })

        self.subtitle = wb.add_format({
            'bold': True, 'font_size': 14, 'font_color': '#4A6FA5',
            'align': 'center', 'valign': 'vcenter', 'font_name': font
        })

        self.header = wb.add_format({
            'bold': True, 'font_size': 12, 'font_color': 'white',
            'bg_color': '#1E3A5F', 'align': 'center', 'valign': 'vcenter',
            'border': 2, 'border_color': '#1E3A5F', 'text_wrap': True, 'font_name': font
        })

        self.sub_header = wb.add_format({
            'bold': True, 'font_size': 11, 'font_color': 'white',
            'bg_color': '#4A6FA5', 'align': 'center', 'valign': 'vcenter',
            'border': 2, 'border_color': '#4A6FA5', 'font_name': font
        })

        # Money formats
        self.money = wb.add_format({
            'num_format': '#,##0 CZK', 'align': 'right', 'border': 1,
            'font_name': font, 'font_size': 11
        })

        self.money_input = wb.add_format({
            'num_format': '#,##0 CZK', 'align': 'right', 'border': 1,
            'bg_color': '#E8F4FD', 'locked': False, 'font_name': font, 'font_size': 11
        })

        self.money_bold = wb.add_format({
            'num_format': '#,##0 CZK', 'align': 'right', 'border': 2,
            'bold': True, 'bg_color': '#D4E6F1', 'font_name': font, 'font_size': 11
        })

        # Percent formats
        self.percent = wb.add_format({
            'num_format': '0.0%', 'align': 'center', 'border': 1,
            'font_name': font, 'font_size': 11
        })

        self.percent_bold = wb.add_format({
            'num_format': '0.0%', 'align': 'center', 'border': 2,
            'bold': True, 'bg_color': '#D4E6F1', 'font_name': font, 'font_size': 11
        })

        # Text formats
        self.text = wb.add_format({
            'align': 'left', 'border': 1, 'valign': 'vcenter',
            'font_name': font, 'font_size': 11
        })

        self.text_center = wb.add_format({
            'align': 'center', 'border': 1, 'valign': 'vcenter',
            'font_name': font, 'font_size': 11
        })

        self.text_input = wb.add_format({
            'align': 'left', 'border': 1, 'bg_color': '#E8F4FD', 'locked': False,
            'font_name': font, 'font_size': 11
        })

        # Date formats
        self.date = wb.add_format({
            'num_format': 'dd-mmm-yyyy', 'align': 'center', 'border': 1,
            'bg_color': '#F5F5F5', 'font_name': font, 'font_size': 11
        })

        self.date_input = wb.add_format({
            'num_format': 'dd-mmm-yyyy', 'align': 'center', 'border': 1,
            'bg_color': '#E8F4FD', 'locked': False, 'font_name': font, 'font_size': 11
        })

        # Status formats
        self.good = wb.add_format({
            'align': 'center', 'border': 2,
            'bg_color': '#27AE60', 'font_color': 'white', 'bold': True,
            'font_name': font, 'font_size': 11
        })

        self.warning = wb.add_format({
            'align': 'center', 'border': 2,
            'bg_color': '#F39C12', 'font_color': 'white', 'bold': True,
            'font_name': font, 'font_size': 11
        })

        self.danger = wb.add_format({
            'align': 'center', 'border': 2,
            'bg_color': '#E74C3C', 'font_color': 'white', 'bold': True,
            'font_name': font, 'font_size': 11
        })

        # Dashboard card formats
        self.card_title = wb.add_format({
            'bold': True, 'font_size': 11, 'font_color': '#7F8C8D',
            'align': 'center', 'valign': 'vcenter', 'font_name': font
        })

        self.card_value = wb.add_format({
            'bold': True, 'font_size': 18, 'font_color': '#1E3A5F',
            'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0 CZK',
            'font_name': font
        })

        self.card_value_green = wb.add_format({
            'bold': True, 'font_size': 18, 'font_color': '#27AE60',
            'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0 CZK',
            'font_name': font
        })

        self.card_value_red = wb.add_format({
            'bold': True, 'font_size': 18, 'font_color': '#E74C3C',
            'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0 CZK',
            'font_name': font
        })

        self.card_percent = wb.add_format({
            'bold': True, 'font_size': 18, 'font_color': '#1E3A5F',
            'align': 'center', 'valign': 'vcenter', 'num_format': '0.0%',
            'font_name': font
        })

        # Selector formats
        self.selector = wb.add_format({
            'bold': True, 'font_size': 14, 'font_color': 'white',
            'bg_color': '#9B59B6', 'align': 'center', 'valign': 'vcenter',
            'border': 2, 'font_name': font
        })

        self.selector_input = wb.add_format({
            'bold': True, 'font_size': 14, 'font_color': '#1E3A5F',
            'bg_color': '#F5EEF8', 'align': 'center', 'valign': 'vcenter',
            'border': 2, 'locked': False, 'font_name': font
        })

        self.currency = wb.add_format({
            'bold': True, 'font_size': 14, 'font_color': 'white',
            'bg_color': '#16A085', 'align': 'center', 'valign': 'vcenter',
            'border': 2, 'font_name': font
        })

        self.currency_input = wb.add_format({
            'bold': True, 'font_size': 14, 'font_color': '#1E3A5F',
            'bg_color': '#E8F8F5', 'align': 'center', 'valign': 'vcenter',
            'border': 2, 'locked': False, 'font_name': font
        })

        # Weekend format for monthly sheets
        self.weekend = wb.add_format({
            'num_format': 'dd-mmm-yyyy', 'align': 'center', 'border': 1,
            'bg_color': '#FADBD8', 'font_name': font, 'font_size': 11
        })

        # Create category-colored formats
        self.category_formats = {}
        for cat, color in CATEGORY_COLORS.items():
            self.category_formats[cat] = wb.add_format({
                'align': 'left', 'border': 2, 'bg_color': color, 'font_color': 'white',
                'bold': True, 'font_name': font, 'font_size': 11
            })

        # Create month-colored formats
        self.month_formats = {}
        self.month_header_formats = {}
        for month, color in MONTH_COLORS.items():
            self.month_formats[month] = wb.add_format({
                'align': 'center', 'border': 2, 'bg_color': color, 'font_color': 'white',
                'bold': True, 'font_name': font, 'font_size': 11
            })
            self.month_header_formats[month] = wb.add_format({
                'bold': True, 'font_size': 12, 'font_color': 'white',
                'bg_color': color, 'align': 'center', 'valign': 'vcenter',
                'border': 2, 'text_wrap': True, 'font_name': font
            })

    def get_category_format(self, category):
        """Get format for a specific category."""
        return self.category_formats.get(category, self.text)

    def get_month_format(self, month):
        """Get format for a specific month."""
        return self.month_formats.get(month, self.text_center)

    def get_month_header_format(self, month):
        """Get header format for a specific month."""
        return self.month_header_formats.get(month, self.header)

    def create_custom_format(self, properties):
        """Create a custom format with given properties."""
        return self.workbook.add_format(properties)
