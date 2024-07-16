from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pyautogui
from time import sleep
from dotenv import load_dotenv

import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class RPA_2:
    
    def __init__(self, empresa_id : int, username_web : str, password_web : str, username_local : str, password_local : str) -> None:
        self.username_web = username_web
        self.password_web = password_web
        self.username_local = username_local
        self.password_local = password_local
        self.empresa_id = empresa_id
        self.url = 'https://www.dominioweb.com.br/'
        self.driver = None
        self.wait_time = 10
        self._process()
    
    def _process(self) -> None:
        self.open_browser() # Abre o navegador e entra na página da URL
        self.login_in_url() # Faz o login com email e senha
        self.click_in_TRComputerPluginWindows() # Clica em abrir
        self.click_in_Escrita_Fiscal() # Clica em Escrita Fiscal
        self.login_in_dominio() # Insere as informações necessárias
        self.click_to_login() # Clica no botão para efetuar o login
        self.click_in_Troca_de_Empresa() # Clica no icone de troca
        self.filling_the_empresa_code() # Preenche o campo com o código da empresa
        self.click_in_ativar() # Clica em ativar empresa
        
        sleep(10)

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
            
            sleep(3)
            abrir_button = pyautogui.locateCenterOnScreen(
                "refer_images/TRComputerPluginWindows.png"
            )
            
            sleep(1)
            logging.info("TRComputerPluginWindows achado")
            pyautogui.moveTo(abrir_button)
            pyautogui.click()
            
        except pyautogui.ImageNotFoundException:
            logging.info("TRComputerPluginWindows não achado")
            self.click_in_TRComputerPluginWindows()

    def click_in_Escrita_Fiscal(self) -> None:
        try:
            self.mouse_to_center()
            
            logging.info("Procurando Escrita Fiscal")
            
            sleep(3)
            escrita_button = pyautogui.locateCenterOnScreen(
                "refer_images/EscritaFiscal.png"
            )
            
            sleep(1)
            logging.info("Escrita Fiscal achada")
            pyautogui.moveTo(escrita_button)
            
            pyautogui.click(clicks=2)
            
        except pyautogui.ImageNotFoundException:
            logging.info("Escrita Fiscal não achada")
            self.click_in_Escrita_Fiscal()
    
    def login_in_dominio(self) -> None:
        try:
            self.mouse_to_center()
            logging.info("Procurando janela de Login")
            
            sleep(3)
            pyautogui.locateCenterOnScreen("refer_images/DominioLogin.png")
            
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
            self.mouse_to_center()
            logging.info("Procurando OK para login")
            
            sleep(3)
            ok_button = pyautogui.locateCenterOnScreen("refer_images/DominioLoginOK.png")
            
            logging.info("OK para login achado")
            pyautogui.moveTo(ok_button)
            pyautogui.click()
        
        except pyautogui.ImageNotFoundException:
            logging.info("OK para login não achado")
            self.click_to_login()
    
    def click_in_Troca_de_Empresa(self) -> None:
        try:
            self.mouse_to_center()
            logging.info("Procurando Troca Icon de Empresa")
            
            sleep(3)
            icon_button = pyautogui.locateCenterOnScreen("refer_images/TrocarEmpresaIcon.png")
            
            logging.info("Troca de Empresa Icon achado")
            pyautogui.moveTo(icon_button)
            pyautogui.click()
        
        except pyautogui.ImageNotFoundException:
            logging.info("Troca de Empresa Icon não achado")
            self.click_in_Troca_de_Empresa()
    
    def filling_the_empresa_code(self) -> None:
        try:
            self.mouse_to_center()
            logging.info("Procurando Troca de Empresa")
            
            sleep(3)
            pyautogui.locateCenterOnScreen("refer_images/TrocarEmpresa.png")
            
            logging.info("Troca de Empresa achado")
            
            pyautogui.write(str(self.empresa_id))
        
        except pyautogui.ImageNotFoundException:
            logging.info("Troca de Empresa não achado")
            self.filling_the_empresa_code()
    
    def click_in_ativar(self) -> None:
        try:
            self.mouse_to_center()
            logging.info("Procurando Ativar troca de empresas")
            
            sleep(3)
            ativar_button = pyautogui.locateCenterOnScreen("refer_images/AtivarEmpresa.png")
            
            logging.info("Ativar troca de empresas achado")
            pyautogui.moveTo(ativar_button)
            pyautogui.click()
        
        except pyautogui.ImageNotFoundException:
            logging.info("Ativar troca de empresas não achado")
            self.click_in_ativar()

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
        username_web=os.environ.get("USERNAME"),
        password_web=os.environ.get("PASSWORD"),
        username_local=os.environ.get("USERNAME2"),
        password_local=os.environ.get("PASSWORD2")
    )
    
    