from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors


def get_header_font():
    font = Font(name='Calibri',
                size=28,
                bold=True,
                italic=False,
                vertAlign=None,
                underline='none',
                strike=False,
                color=colors.WHITE)
    return font

def greenFill():
    fill = PatternFill(fill_type='solid',start_color='00B050',end_color='00B050')
    return fill