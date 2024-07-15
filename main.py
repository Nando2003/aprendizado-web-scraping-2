from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
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
        
        sleep(5) # Para o dev visualizar

    def click_in_TRComputerPluginWindows(self) -> None:
        self.mouse_to_center() # Move o mouse para o centro da tela
        
        pyautogui.locateAllOnScreen("./refer_images/TRComputerPluginWindows.png")
        pyautogui.click()
    
    def __del__(self) -> None:
        sleep(10)
        if self.driver is not None:
            self.driver.close()
    
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
    
    