from rpa_2 import RPA_2
from dotenv import load_dotenv
import os

if __name__ == "__main__":
    # Pega o diretorio que este arquivo se encontra + dotenv/.env
    dotenv_path = os.path.join(os.path.dirname(__file__), "dotenv_files/.env") 
    load_dotenv(dotenv_path)
    
    RPA_2(
        excel_path="./excel_files/Empresas.xlsx",
        username_web=os.environ.get("USERNAME1"),
        password_web=os.environ.get("PASSWORD1"),
        username_local=os.environ.get("USERNAME2"),
        password_local=os.environ.get("PASSWORD2")
    )
    
    