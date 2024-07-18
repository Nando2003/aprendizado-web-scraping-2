# Biblioteca para web scraping
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Biblioteca para manipular o ponteiro do mouse
import pyautogui

# Bibliotecas relacionadas a tempo
from time import sleep
from datetime import date
from calendar import monthrange

# Manipular arquivos
import os

# Funções criadas pelo dev
from import_xlsx import import_column_from_xlsx
from editing_xlsx import editing_xlsx

# Logging.info('')
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class RPA_2:
    
    def __init__(self, excel_path:str, username_web:str, password_web:str, username_local:str, password_local:str, download_path:str=None) -> None:
        self.username_web = username_web
        self.password_web = password_web
        self.username_local = username_local
        self.password_local = password_local
        self.excel_path = excel_path
        self.empresas_id = self.list_of_empresas() 
        # self.empresas_id = [413, 346]
        self.download_path = download_path
        self.status = []     
        self.url = 'https://www.dominioweb.com.br/'
        self.driver = None
        self.wait_time = 10
        self._process()
    
    def _process(self) -> None:
        self.new_download_path() # Caso download_path não seja None, será criado um path
        self.open_browser() # Abre o navegador e entra na página da URL
        self.login_in_url() # Faz o login com email e senha
        self.open_dominio() # Clica em abrir
        self.click_in_escrita_fiscal() # Clica em Escrita Fiscal
        self.login_in_dominio() # Insere as informações necessárias
        self.click_to_login() # Clica no botão para efetuar o login
        self.wait_to_load() # Espera a escrita fiscal carregar
        
        for empresa_id in self.empresas_id:
            self.click_in_troca_de_empresa() # Clica no icone de troca
            self.filling_the_empresa_code(empresa_id=empresa_id) # Preenche o campo com o código da empresa
            self.click_in_ativar() # Clica em ativar empresa
            self.open_contribuicoes() # Abre EFD contribuições
            self.typing_the_date() # Digita a data do primeiro e ultimo dia do mês passado
            self.typing_the_download_path(empresa_id=empresa_id) # Digita o path de download
            self.click_to_download() # Clica no botão de download
            self.wait_to_download(empresa_id=empresa_id) # Espera o download ser realizado
            self.close_download() # Fecha a aba de download
        
        self.editing_excel() # Edita o excel com os status de cada empresa
        
        sleep(10)

    def list_of_empresas(self) -> list:
        return import_column_from_xlsx(
            excel_path=self.excel_path,
            linha=2, 
            coluna='B'
        )
    
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

    def open_dominio(self) -> None:
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
            self.open_dominio()

    def click_in_escrita_fiscal(self) -> None:
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
            self.click_in_escrita_fiscal()
    
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
            
            sleep(5)
            self.close_warning()
            self.close_alert()
        
        except pyautogui.ImageNotFoundException:
            self.wait_to_load()
    
    def click_in_troca_de_empresa(self) -> None:
        try:    
            logging.info("Procurando Troca Icon de Empresa")
            icon_button = pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/TrocarEmpresaIcon.png")

            sleep(1)
            logging.info("Troca de Empresa Icon achado")
            pyautogui.moveTo(icon_button)
            pyautogui.click()
        
        except pyautogui.ImageNotFoundException:
            logging.info("Troca de Empresa Icon não achado")
            self.click_in_troca_de_empresa()
    
    def filling_the_empresa_code(self, empresa_id) -> None:
        try:
            logging.info("Procurando Troca de Empresa")
            pyautogui.locateCenterOnScreen("refer_images/Dominio/Escritura/Empresa/TrocarEmpresa.png")
            
            sleep(1)
            logging.info("Troca de Empresa achado")
            logging.info(f"Empresa digitada: {empresa_id}")
            pyautogui.write(str(empresa_id))
        
        except pyautogui.ImageNotFoundException:
            logging.info("Troca de Empresa não achado")
            self.filling_the_empresa_code(empresa_id)
    
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
            
            ordem = ["n", "f"]
            
            for letra in ordem:
                logging.info(f"Tecla {letra} pressionada")
                pyautogui.press(letra)
            
            sleep(1)
            efd_button = pyautogui.locateCenterOnScreen(
                "refer_images/Dominio/Escritura/Empresa/EFD_Contribuicoes/EFD_Contribuicoes.png"
            )
            
            pyautogui.moveTo(efd_button)
            pyautogui.click()
            
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
            last_mouth = today_mouth - 1 # Correto
            # last_mouth = today_mouth - 2 # Teste
                    
        last_day = str((monthrange(today_year, last_mouth))[1])
        last_mouth = str(last_mouth)
        today_year = str(today_year)

        if not(last_mouth in ["10", "11"]):
            last_mouth = "0" + str(last_mouth)
                
        first_day_date = "01" + last_mouth + today_year
        last_day_date = last_day + last_mouth + today_year
        
        sleep(0.5)
        pyautogui.press("backspace")
        pyautogui.write(first_day_date)
        pyautogui.press("tab")
        
        sleep(0.5)
        pyautogui.press("backspace")
        pyautogui.write(last_day_date)
    
    def typing_the_download_path(self, empresa_id) -> None:
        complete_download_path = os.path.join(
            self.download_path, 
            f"{empresa_id}.txt"
        )
        
        sleep(0.1)
        pyautogui.press('tab', presses=2)
        pyautogui.press('backspace')
        pyautogui.write(complete_download_path)
        
    def click_to_download(self) -> None:
        sleep(0.1)
        pyautogui.press('tab', presses=2)
        pyautogui.press('enter')
            
    def wait_to_download(self, empresa_id) -> None:
        try:
            pyautogui.locateCenterOnScreen(
                    "refer_images/Dominio/Escritura/Empresa/WarningEmpresa.png"
            )
            
            sleep(0.5)
            
            try:
                pyautogui.locateCenterOnScreen(
                    "refer_images/Dominio/Escritura/Empresa/EFD_Contribuicoes/DownloadComplete.png"
            )
                logging.info("Download concluido")
                
                self.mouse_to_center()
                pyautogui.click()
                pyautogui.press('esc')
                
                self.status.append(("CONCLUIDO", "90ee90"))
                
            except pyautogui.ImageNotFoundException:
                
                logging.info("Download Incompleto")
                pyautogui.screenshot(imageFilename=f"./screenshot/{empresa_id}.png")
                    
                self.mouse_to_center()
                pyautogui.click()
                pyautogui.press('esc')
                    
                self.status.append(("FALHA", "ff6961"))
                    
        except pyautogui.ImageNotFoundException:
            try:
                sleep(0.1)
                pyautogui.locateCenterOnScreen(
                    "refer_images/Dominio/Escritura/Empresa/AlertEmpresa.png"
                )
                logging.info("Dados não digitados")
                        
                self.mouse_to_center()
                pyautogui.click()
                pyautogui.press('esc')
                    
                self.status.append(("FALTA DE DADOS", "ffffe0"))
                
            except pyautogui.ImageNotFoundException:
                try:
                    sleep(0.1)
                    pyautogui.locateCenterOnScreen(
                        "refer_images/Dominio/Escritura/Empresa/QuestionEmpresa.png"
                    )
                    logging.info("Copiando o ultimo mês digitado")
                            
                    self.mouse_to_center()
                    pyautogui.click()
                    pyautogui.press('y')
                    
                    self.wait_to_download(empresa_id)
                    
                except pyautogui.ImageNotFoundException:
                    logging.info("Esperando Download")
                    self.wait_to_download(empresa_id)
    
    def editing_excel(self) -> None:
        if editing_xlsx(
          excel_path=self.excel_path,
          data=self.status,
          linha=3,
          coluna='C'
        ):
            logging.info("Excel editado")
        else:
            logging.info("Erro ao editar o Excel")
    
    def close_download(self) -> None:
        sleep(0.1)
        pyautogui.press('tab')
        pyautogui.press('enter')
        logging.info("Fechando aba de download")
    
    def close_dominio(self) -> None:
        logging.info('Fechando Escrita Fiscal...')
        sleep(0.1)
        pyautogui.hotkey('alt', 'f4')
        sleep(0.1)
        pyautogui.press('enter')
        
        sleep(1)
        self.mouse_to_center()
        pyautogui.click()
        
        logging.info('Fechando Domínio WEB...')
        sleep(0.1)
        pyautogui.hotkey('alt', 'f4')
        sleep(0.1)
        pyautogui.press('enter')
    
    def __del__(self) -> None:
        self.close_dominio() # Fechando Domínio
        
        if self.driver is not None:
            self.driver.close()
            logging.info("Processo Encerrado")
    
    def __str__(self) -> str:
        return self.username
