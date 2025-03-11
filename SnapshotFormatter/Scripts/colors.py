import enum
from openpyxl.styles import PatternFill, Border, Side


class Color(str, enum.Enum):
    """Enumeration class of colors, including CISLA themes"""
    # values must be strings
    # Logo Colors
    LOGOGREEN = '008751'
    LOGOBLUE = '00539E'
    RED = 'E21836'
    ORANGE = 'F68E1E'
    BLACK = '000000'
    WHITE = 'FFFFFF'
    
    # Campaign Colors
    GREEN = '78BE20'
    BLUE = '489FDF'
    YELLOW = 'FFCE00'
    GREY = 'B2BEB5'
    BORDER = 'D9D9D9'
    HEADER = BLUE
    EMPTY = WHITE


def pattern_fill(cell, color: Color):
    thin_border = Border(left=Side(border_style='thin', color=Color.BORDER), 
                     right=Side(border_style='thin', color=Color.BORDER), 
                     top=Side(border_style='thin', color=Color.BORDER), 
                     bottom=Side(border_style='thin', color=Color.BORDER))
    cell.border = thin_border
    cell.fill = PatternFill(patternType='solid', fgColor=color)
    
    