import pyautogui
import logging
from time import sleep
import os

logging.basicConfig(
        level=logging.INFO, 
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

class EFDContribuicoes:
    
    def __init__(self, contribuicoes_path:list) -> None:
        self.contribuicoes_path = contribuicoes_path
        self._process()
        
    def _process(self) -> None:
        self.open_efd()
        
        for path in self.contribuicoes_path:
            self.import_shortcut()
            self.import_file(path)
            self.close_pop_ups_download()
            self.download_file(path)
            self.close_escritura()
            self.delete_efd_from_db()
            self.close_pop_ups_delete()
            
        sleep(5)
    
    def open_efd(self) -> None:
        pyautogui.press('winleft')
        
        sleep(0.5)
        pyautogui.write('EFD Contribui')
        
        sleep(0.5)
        pyautogui.press('enter')
        
        logging.info('Abrindo EFD Contribuições')
        
        sleep(30)
        
    def import_shortcut(self) -> None:
        pyautogui.hotkey('ctrl', 'i')
        logging.info('Pressionando ctrl + i')
    
    def import_file(self, path:str) -> None:
        sleep(3)
        if os.path.isfile(path):
            if path.lower().endswith('.txt'):
                pyautogui.write(path)
                pyautogui.press('enter')
                logging.info('Arquivo importado')
            else:
                raise ValueError(f"Erro: {self.contribuicoes_path} não é um arquivo .txt")
        else:
            raise FileNotFoundError(f"Erro: {self.contribuicoes_path} não é um arquivo válido")
    
    def close_pop_ups_download(self) -> None:
        try:
            pyautogui.locateCenterOnScreen(
                image='refer_images/EFD_Contribuicoes/EscrituraAssinada.png'
            )
            logging.info('Fechando pop-ups')
            pyautogui.press('enter', presses=2, interval=1.5)
            
            while(True):
                try:
                    pyautogui.locateCenterOnScreen(
                        image='refer_images/EFD_Contribuicoes/ArquivoAvisos.png'
                    )
                    
                    sleep(1)
                    pyautogui.press('enter', presses=2, interval=1.5)
                    break
                
                except pyautogui.ImageNotFoundException:
                    ...
            
        except pyautogui.ImageNotFoundException:
            self.close_pop_ups_download()
    
    def download_file(self, path:str) -> None:
        sleep(3)
        download_path = path[:-3] + ".pdf"
        for i in range(7):
            pyautogui.press('tab')
            
        pyautogui.press('enter')
        
        sleep(3)
        pyautogui.write(download_path)
        pyautogui.press('enter')
        
        logging.info('Baixando pdf')
    
    def close_escritura(self) -> None:
        sleep(1)
        pyautogui.hotkey('ctrl', 'f')
        
        logging.info('Fechando escritura')
    
    def delete_efd_from_db(self) -> None:
        sleep(1)
        pyautogui.hotkey('ctrl', 'e')
        sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        sleep(0.5)
        pyautogui.press('enter')
        
        logging.info('Deletando EFD do banco de dados')
        
    def close_pop_ups_delete(self) -> None:
        sleep(1)
        pyautogui.press('enter')
        logging.info('Fechando pop-ups')
        
        while(True):
            try:
                pyautogui.locateCenterOnScreen(
                    image='refer_images/EFD_Contribuicoes/EscrituraExcluida.png'
                )
                pyautogui.press('enter')
                break
            
            except pyautogui.ImageNotFoundException:
                ...
        
        sleep(3)
        
    def __del__(self) -> None:
        pyautogui.hotkey(
            'alt', 'f4'
        )
        logging.info('Processo encerrado')

if __name__ in "__main__":
    EFDContribuicoes(
        'C:\Relatorio\GEM-SHIPPING-LTDA\\05.2024-Sped_Contribuicao.txt',
        'C:\Relatorio\LUZ-PUBLICIDADE-LTDA\-05.2024-Sped_Contribuicao.txt'
    )