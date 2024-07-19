import os 
import logging
from functions.import_xlsx import import_column_from_xlsx

logging.basicConfig(
        level=logging.INFO, 
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

class CreateDir:
    
    def __init__(self, excel_path:str, dir_path:str =None) -> None:
        self.excel_path = excel_path
        self.dir_path = dir_path
        self.companies_path = []
        self._process()
        
    def _process(self) -> None:
        self.creating_dir_path()
        companies_name = self.import_name_of_the_company()
        
        for company_name in companies_name:
            company_path = self.creating_dir_for_companies(
                dir_path=self.dir_path,
                company_name=company_name
            )
            
            self.companies_path.append(company_path)
            
    def creating_dir_path(self) -> None:
        logging.info('Criando path para diretorio')
        if self.dir_path is not None:
            if os.path.isdir(self.dir_path):
                logging.info('Path passado existente')
                self.dir_path = os.path.abspath(self.dir_path)
            else:
                logging.info('Path passado não existente')
                self.dir_path = "C:\\Mia"
                if os.path.isdir(self.dir_path) is False:
                    logging.info('Criando diretorio Mia')
                    os.mkdir(self.dir_path)
        else:
            logging.info('Path é None')
            self.dir_path = "C:\\Mia"
            if os.path.isdir(self.dir_path) is False:
                logging.info('Criando diretorio Mia')
                os.mkdir(self.dir_path)
        
    def creating_dir_for_companies(self, dir_path:str, company_name:str) -> str:
        company_name_without_space = company_name.replace(" ", "-")
        dir_path_join = os.path.join(dir_path, company_name_without_space)
        if os.path.isdir(dir_path_join) is False:
            logging.info(f'Criando pasta para a empresa {company_name}')
            os.mkdir(dir_path_join)
        
        return dir_path_join
    
    def import_name_of_the_company(self) -> list:
        return import_column_from_xlsx(
            excel_path=self.excel_path,
            linha=2, 
            coluna='A'
        )
    
    def get_companies_path_for_download(self) -> list:
        return self.companies_path
        
if __name__ == "__main__":
    create_dir = CreateDir("./excel_files/Empresas.xlsx")
    companies_path = create_dir.get_companies_path_for_download()
    print(companies_path)