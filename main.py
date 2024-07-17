from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pyautogui
from time import sleep
from dotenv import load_dotenv
from datetime import date
from calendar import monthrange
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class RPA_2:
    
    def __init__(self, empresa_id:int, username_web:str, password_web:str, username_local:str, password_local:str, download_path=None) -> None:
        self.username_web = username_web
        self.password_web = password_web
        self.username_local = username_local
        self.password_local = password_local
        self.empresa_id = empresa_id
        self.download_path = download_path        
        self.url = 'https://www.dominioweb.com.br/'
        self.driver = None
        self.wait_time = 10
        self._process()
    
    def _process(self) -> None:
        self.new_download_path() # Caso download_path não seja None, será criado um path
        self.open_browser() # Abre o navegador e entra na página da URL
        self.login_in_url() # Faz o login com email e senha
        self.click_in_TRComputerPluginWindows() # Clica em abrir
        self.click_in_Escrita_Fiscal() # Clica em Escrita Fiscal
        self.login_in_dominio() # Insere as informações necessárias
        self.click_to_login() # Clica no botão para efetuar o login
        self.wait_to_load() # Espera a escrita fiscal carregar
        self.click_in_Troca_de_Empresa() # Clica no icone de troca
        self.filling_the_empresa_code() # Preenche o campo com o código da empresa
        self.click_in_ativar() # Clica em ativar empresa
        self.open_contribuicoes() # Abre EFD contribuições
        self.typing_the_date() # Digita a data do primeiro e ultimo dia do mês passado
        self.typing_the_download_path() # Digita o path de download
        self.click_to_download() # Clica no botão de download
        self.wait_to_download() # Espera o download ser realizado
        
        sleep(10)

    def new_download_path(self) -> None:
        if self.download_path is not None:
            self.download_path = os.path.realpath(self.download_path)
            if os.path.realpath(self.download_path) and self.download_path[0:2] == "C:":
                self.download_path = "M:\\" + self.download_path[2::]
                return
        
        self.download_path = "M:\\Mia"
    
    def mouse_to_center(self) -> None:
        screen_width, screen_height = pyautogui.size()
        pyautogui.moveTo(screen_width/2, screen_height/2)
    
    def open_browser(self) -> None:
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.get(self.url)
    
    def login_in_url(self) -> None:
        sleep(1)
        email = WebDriverWait(
            self.driver,
            self.wait_time
        ).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[1]/span[2]/input"
                )
            )
        )
        
        email.send_keys(self.username_web)
        
        sleep(1)
        senha = WebDriverWait(
            self.driver,
            self.wait_time
        ).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[2]/span[2]/input"
                )
            )
        )
        
        senha.send_keys(self.password_web)
        
        sleep(1)
        entrar = WebDriverWait(
            self.driver,
            self.wait_time
        ).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/div/button"
                )
            )
        )
        
        entrar.click()
        
        # sleep(5) # Para o dev visualizar

    def click_in_TRComputerPluginWindows(self) -> None:
        try:
            self.mouse_to_center() # Move o mouse para o centro da tela

            logging.info("Procurando TRComputerPluginWindows")
            
            abrir_button = pyautogui.locateCenterOnScreen(
                "refer_images/Browser/TRComputerPluginWindows.png"
            )
            
            logging.info("TRComputerPluginWindows achado")
            
            sleep(0.1)
            pyautogui.moveTo(abrir_button)
            pyautogui.click()
            
        except pyautogui.ImageNotFoundException:
            logging.info("TRComputerPluginWindows não achado")
            self.click_in_TRComputerPluginWindows()

    def click_in_Escrita_Fiscal(self) -> None:
        try:
            self.mouse_to_center()
            
            logging.info("Procurando Escrita Fiscal")
            
            escrita_button = pyautogui.locateCenterOnScreen(
                "refer_images/Dominio/Inicial/EscritaFiscal.png", 
            )
            
            logging.info("Escrita Fiscal achada")
            pyautogui.moveTo(escrita_button)
            
            pyautogui.click(clicks=2)
            
        except pyautogui.ImageNotFoundException:
            logging.info("Escrita Fiscal não achada")
            self.click_in_Escrita_Fiscal()
    
    def login_in_dominio(self) -> None:
        try:
            logging.info("Procurando janela de Login")
            
            pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/DominioLogin.png")
            logging.info("Janela de Login achada")
            
            sleep(1)
            pyautogui.press('backspace',presses=10)
            pyautogui.write(self.username_local)
            pyautogui.press('tab')

            logging.info("Usuario adicionado")
            
            sleep(1)
            pyautogui.press('backspace',presses=10)
            pyautogui.write(self.password_local)
            
            logging.info("Senha adicionada")
        
        except pyautogui.ImageNotFoundException:
            logging.info("Janela de Login não achada")
            self.login_in_dominio()
    
    def click_to_login(self) -> None:
        try:
            logging.info("Procurando OK para login")
            ok_button = pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/DominioLoginOK.png")
            
            sleep(1)
            logging.info("OK para login achado")
            pyautogui.moveTo(ok_button)
            pyautogui.click()
        
        except pyautogui.ImageNotFoundException:
            logging.info("OK para login não achado")
            self.click_to_login()
    
    def wait_to_load(self) -> None:
        try:
            logging.info("Escrita carregando...")
            pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/FullLoad.png")

            sleep(1)
            logging.info("Escrita carregada")
        
        except pyautogui.ImageNotFoundException:
            self.wait_to_load()
    
    def click_in_Troca_de_Empresa(self) -> None:
        try:
            sleep(5)
            self.close_warning()
            self.close_alert()
                
            logging.info("Procurando Troca Icon de Empresa")
            icon_button = pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/TrocarEmpresaIcon.png")

            sleep(1)
            logging.info("Troca de Empresa Icon achado")
            pyautogui.moveTo(icon_button)
            pyautogui.click()
        
        except pyautogui.ImageNotFoundException:
            logging.info("Troca de Empresa Icon não achado")
            self.click_in_Troca_de_Empresa()
    
    def filling_the_empresa_code(self) -> None:
        try:
            logging.info("Procurando Troca de Empresa")
            pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/TrocarEmpresa.png")
            
            sleep(1)
            logging.info("Troca de Empresa achado")
            pyautogui.write(str(self.empresa_id))
        
        except pyautogui.ImageNotFoundException:
            logging.info("Troca de Empresa não achado")
            self.filling_the_empresa_code()
    
    def click_in_ativar(self) -> None:
        try:
            logging.info("Procurando Ativar troca de empresas")
            ativar_button = pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/AtivarEmpresa.png")
            
            sleep(1)
            logging.info("Ativar troca de empresas achado")
            pyautogui.moveTo(ativar_button)
            pyautogui.click()
            
            sleep(5)
            self.close_warning()
            self.close_alert()
        
        except pyautogui.ImageNotFoundException:
            logging.info("Ativar troca de empresas não achado")
            self.click_in_ativar()
    
    def close_warning(self, i=0) -> None:
        try:
            while(i<=2):
                sleep(1)
                logging.info(f"Procurando Warning {i+1}")
                pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/WarningEmpresa.png")
                self.mouse_to_center()
                
                logging.info("Warning encontrado")
                pyautogui.click()
                pyautogui.press('esc')
                
                sleep(0.1)
                self.close_question()
                break
            
            else:
                logging.info("Nenhum Warning encontrado")
                
        except pyautogui.ImageNotFoundException:
            i = i + 1
            self.close_warning(i)
    
    def close_alert(self, i=0) -> None:
        try:
            while(i<=1):
                sleep(1)
                logging.info(f"Procurando Alert {i+1}")
                pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/AlertEmpresa.png")
                self.mouse_to_center()
                
                logging.info("Alert encontrado")
                pyautogui.click()
                pyautogui.press('esc')
                break
                
            else:
                logging.info("Nenhum Alert encontrado")
            
        except pyautogui.ImageNotFoundException:
            i = i + 1
            self.close_alert(i)
    
    def close_question(self, i=0) -> None:
        try:
            while(i<=2):
                sleep(1)
                logging.info(f"Procurando Question {i+1}")
                pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/QuestionEmpresa.png")
                self.mouse_to_center()
                
                logging.info("Question encontrado")
                no_button = pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/QuestionNo.png")
                pyautogui.moveTo(no_button)
                pyautogui.click()
                break
                
            else:
                logging.info("Nenhum Question encontrado")
            
        except pyautogui.ImageNotFoundException:
            i = i + 1
            self.close_question(i)
      
    def open_contribuicoes(self) -> None:
        try:
            logging.info("Procurando Relatórios")
            relatorios_button = pyautogui.locateCenterOnScreen(
                "refer_images/Dominio/Escritura/Empresa/Relatorios.png"
            )
            
            logging.info("Relatórios achado")
            
            sleep(1)
            pyautogui.moveTo(relatorios_button)
            pyautogui.click()
            
            ordem = ["n", "f", "o"]
            
            for letra in ordem:
                logging.info(f"Tecla {letra} pressionada")
                pyautogui.press(letra)
            
        except pyautogui.ImageNotFoundException:
            logging.info("Relatórios não achado")
            self.open_contribuicoes()
    
    def typing_the_date(self) -> None:
        today_date = date.today()
        today_mouth = today_date.month
        today_year = today_date.year
                        
        if today_mouth == 1:
            last_mouth = 12
            today_year = today_year - 1      
        else:
            last_mouth = today_mouth - 1
                    
        last_day = str((monthrange(today_year, last_mouth))[1])
        last_mouth = str(last_mouth)
        today_year = str(today_year)

        if not(last_mouth in ["10", "11"]):
            last_mouth = "0" + str(last_mouth)
                
        first_day_date = "01" + last_mouth + today_year
        last_day_date = last_day + last_mouth + today_year
        
        sleep(1)
        pyautogui.write(first_day_date)
        pyautogui.press("tab")
        
        sleep(1)
        pyautogui.write(last_day_date)
    
    def typing_the_download_path(self) -> None:
        self.download_path = os.path.join(
            self.download_path, 
            f"{self.empresa_id}.txt"
        )
        
        sleep(1)
        pyautogui.press('tab', presses=2)
        pyautogui.write(self.download_path)
        
    def click_to_download(self) -> None:
        sleep(1)
        pyautogui.press('tab', presses=2)
        pyautogui.press('enter')
            
    def wait_to_download(self) -> None:
        try:
            pyautogui.locateCenterOnScreen(
                    "refer_images/Dominio/Escritura/Empresa/WarningEmpresa.png"
            )
            
            try:
                pyautogui.locateCenterOnScreen(
                    "refer_images/Dominio/Escritura/Empresa/EFD_Contribuicoes/DownloadComplete.png"
            )
                logging.info("Download concluido")
                
                self.mouse_to_center()
                pyautogui.click()
                pyautogui.press('esc')
                
            except pyautogui.ImageNotFoundException:
                
                try:
                    pyautogui.locateCenterOnScreen(
                    image="refer_images/Dominio/Escritura/Empresa/EFD_Contribuicoes/DownloadIncomplete.png"
                )
                    logging.info("Download Incompleto")
                    
                    self.mouse_to_center()
                    pyautogui.click()
                    pyautogui.press('esc')
                    
                except pyautogui.ImageNotFoundException:
                    self.wait_to_download()
            
        except pyautogui.ImageNotFoundException:
            logging.info("Esperando Download")
            self.wait_to_download()
        
    def __del__(self) -> None:
        if self.driver is not None:
            self.driver.close()
            logging.info("Processo Encerrado")
    
    def __str__(self) -> str:
        return self.username

if __name__ == "__main__":
    # Pega o diretorio que este arquivo se encontra + dotenv/.env
    dotenv_path = os.path.join(os.path.dirname(__file__), "dotenv_files/.env") 
    load_dotenv(dotenv_path)
    
    RPA_2(
        empresa_id=110,
        username_web=os.environ.get("USERNAME1"),
        password_web=os.environ.get("PASSWORD1"),
        username_local=os.environ.get("USERNAME2"),
        password_local=os.environ.get("PASSWORD2")
    )
    
    