"""Shared visual language for all three Azalea workbooks.

Color code (per build spec section 7):
 blue = hard input the user can change
 black = formula
 green = cross-sheet / cross-tab link
 red = external link (none used; reserved)

Number formats (spec rule 7): parentheses for negatives, '-' for zeros,
$#,##0 currency, 0.0% percentages, 0.0x multiples.
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

# ---- Palette -------------------------------------------------------------
NAVY = "1F3864"        # section headers
MIDBLUE = "2E5496"     # sub-headers
LIGHTBLUE = "D9E1F2"   # banded / input-block fill
PALEBLUE = "EEF3FB"    # zebra banding
GREYHDR = "F2F2F2"     # period header strip
INPUTFILL = "FFF2CC"   # yellow-ish fill behind editable inputs (CONFIRM cells)
GREENFILL = "E2EFDA"   # check-row OK fill
REDFLAG = "FCE4D6"     # warning / >10% variance

WHITE = "FFFFFF"
BLACK = "000000"
BLUE = "0000CC"        # input font
GREEN = "006100"       # cross-sheet link font
RED = "C00000"         # external link font / negative emphasis

# ---- Number formats ------------------------------------------------------
FMT_CUR = '$#,##0;($#,##0);"-"'
FMT_CUR2 = '$#,##0.00;($#,##0.00);"-"'
FMT_NUM = '#,##0;(#,##0);"-"'
FMT_NUM1 = '#,##0.0;(#,##0.0);"-"'
FMT_NUM2 = '0.00'
FMT_NUM4 = '0.0000'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_PCT2 = '0.00%;(0.00%);"-"'
FMT_MULT = '0.00x;(0.00x);"-"'
FMT_RATE = '$#,##0.00'
FMT_INT = '#,##0'

# ---- Fonts ---------------------------------------------------------------
BASE_FONT = "Calibri"

def f_title():
    return Font(name=BASE_FONT, size=16, bold=True, color=WHITE)

def f_subtitle():
    return Font(name=BASE_FONT, size=10, italic=True, color="D9D9D9")

def f_section():
    return Font(name=BASE_FONT, size=11, bold=True, color=WHITE)

def f_sub():
    return Font(name=BASE_FONT, size=10, bold=True, color=NAVY)

def f_label(bold=False):
    return Font(name=BASE_FONT, size=10, bold=bold, color=BLACK)

def f_input():
    return Font(name=BASE_FONT, size=10, color=BLUE)

def f_formula(bold=False):
    return Font(name=BASE_FONT, size=10, bold=bold, color=BLACK)

def f_link(bold=False):
    return Font(name=BASE_FONT, size=10, bold=bold, color=GREEN)

def f_total():
    return Font(name=BASE_FONT, size=10, bold=True, color=BLACK)

def f_note():
    return Font(name=BASE_FONT, size=8, italic=True, color="808080")

# ---- Fills ---------------------------------------------------------------
def fill(hex_):
    return PatternFill("solid", fgColor=hex_)

# ---- Borders -------------------------------------------------------------
_thin = Side(style="thin", color="BFBFBF")
_med = Side(style="medium", color=NAVY)
BORDER_THIN = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
BORDER_BOX = Border(left=_med, right=_med, top=_med, bottom=_med)
BORDER_TOP = Border(top=Side(style="thin", color=BLACK))
BORDER_TOPBOT = Border(top=Side(style="thin", color=BLACK),
                       bottom=Side(style="double", color=BLACK))

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def note(text):
    """Build a sized cell comment (source citation / flag)."""
    c = Comment(text, "Azalea Hospice")
    c.width = 320
    c.height = 120
    return c
