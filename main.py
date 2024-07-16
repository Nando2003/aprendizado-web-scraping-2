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
    
    def __init__(self, username : str, password : str) -> None:
        self.username = username
        self.password = password
        self.url = 'https://www.dominioweb.com.br/'
        self.driver = None
        self.wait_time = 10
        self._process()
    
    def _process(self) -> None:
        self.open_browser() # Abre o navegador e entra na página da URL
        self.login_in_url() # Faz o login com email e senha
        self.click_in_TRComputerPluginWindows() # Clica em abrir
        self.click_in_Escrita_Fiscal() # Clica em Escrita Fiscal
        
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
        
        email.send_keys(self.username)
        
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
        
        senha.send_keys(self.password)
        
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
    
            sleep(10)
            logging.info("Procurando TRComputerPluginWindows")
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
            
            sleep(20)
            logging.info("Procurando Escrita Fiscal")
            escrita_button = pyautogui.locateCenterOnScreen(
                "refer_images/EscritaFiscal.png"
            )
            
            sleep(1)
            logging.info("Escrita Fiscal achada")
            pyautogui.moveTo(escrita_button)
            
            for i in range(2):
                pyautogui.click()
            
        except pyautogui.ImageNotFoundException:
            logging.info("Escrita Fiscal não achada")
            self.click_in_Escrita_Fiscal()
    
    def __del__(self) -> None:
        if self.driver is not None:
            self.driver.close()
            logging.info("Processo Encerrado")
    
    def __str__(self) -> str:
        return self.username

if __name__ == "__main__":
    # Pega o diretorio que este arquivo se encontra + .env
    dotenv_path = os.path.join(os.path.dirname(__file__), ".env") 
    load_dotenv(dotenv_path)
    
    RPA_2(
        username=os.environ.get("USERNAME"),
        password=os.environ.get("PASSWORD")
    )
    
    