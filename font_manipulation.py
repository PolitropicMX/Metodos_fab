#    STYLING FONTS AND COLORS
import openpyxl
from openpyxl.styles import Font, Color

## FONT FUNCTION
## Font(
## name = None,
## sz = None,
## b = None,
## i = None,
## charset = None,
## u = None,
## strike = None,
## color = None,
## scheme = None,
## family = None,
## size = None,
## bold = None,
## italic = None,
## strikethrough = None,
## underline = None,
## vertAlign = None,
## outline = None,
## shadow = None,
## condense = None,
## extend = None)

wb = openpyxl.load_workbook("pruebafonts.xlsx")
ws = wb.active


ws['A1'].font = Font(name="Berlin Sans FB Demi",color="215967",size=20)

wb.save("pruebafonts.xlsx")
