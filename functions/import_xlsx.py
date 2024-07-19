from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

def import_column_from_xlsx(excel_path:str, linha:int =1, coluna:str ='A', index:bool =False) -> list:
    try:
        load_wb = load_workbook(excel_path)
        sheet = load_wb.active
        
        cell_ID = [cell for cell in sheet[coluna][linha:]]
        cell_value = [cell.value for cell in cell_ID if cell.value is not None]
        
        if index is True:
            indexed_cell_value = [(cell.row, cell.value) for cell in cell_ID if cell.value is not None]
            return indexed_cell_value # [(index, value)...]
        
        return cell_value # [value...]
    
    except ValueError:
        raise ValueError("Digite uma coluna/linha existente!")

    except InvalidFileException:
        raise InvalidFileException("O caminho deverá levar até um arquivo .xlsx")