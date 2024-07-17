from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

def import_column_from_xlsx(excel_path:str, linha:int =1, coluna:str ='A') -> list:
    try:
        load_wb = load_workbook(excel_path)
        sheet = load_wb.active
        
        cell_ID = [cell for cell in sheet[coluna][linha:]]
        cell_value = [cell.value for cell in cell_ID if cell.value is not None]
        
    except ValueError:
        raise ValueError("Digite uma coluna/linha existente!")

    except InvalidFileException:
        raise InvalidFileException("O caminho deverá levar até um arquivo .xlsx")
    
    return cell_value