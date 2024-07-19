from bot.dominio import DominioWeb
from bot.create_dir import CreateDir
from bot.efd import EFDContribuicoes
from dotenv import load_dotenv
import os

if __name__ == "__main__":
    # Pega o diretorio que este arquivo se encontra + dotenv/.env
    dotenv_path = os.path.join(os.path.dirname(__file__), "dotenv_files/.env") 
    load_dotenv(dotenv_path)
    
    excel_path = './excel_files/Empresas.xlsx'
    download_path = None
        
    create_dir = CreateDir(
        excel_path=excel_path
    )
    
    companies_path = create_dir.get_companies_path_for_download()
    
    dominio_web = DominioWeb(
        download_path=companies_path,
        excel_path=excel_path,
        username_web=os.environ.get("USERNAME1"),
        password_web=os.environ.get("PASSWORD1"),
        username_local=os.environ.get("USERNAME2"),
        password_local=os.environ.get("PASSWORD2")
    )
    
    companies_path_txt = dominio_web.get_companies_path_txt()
        
    """EFDContribuicoes(
        contribuicoes_path=companies_path_txt
    )"""
    