import time
from datetime import datetime
from openpyxl.utils import get_column_letter


def getYear():
    years = []
    for year in range(2):
        years.append(datetime.now().year + year)
    return years


def coordinate_cell_to_string(col, row):
    col_letter = get_column_letter(col)
    cell_string = f"{col_letter}{row}"
    
    return cell_string
    
