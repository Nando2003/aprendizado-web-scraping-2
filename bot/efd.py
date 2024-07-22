from functions import editing_xlsx, import_column_from_xlsx
from time import sleep
import pyautogui
import logging
import os

logging.basicConfig(
        level=logging.INFO, 
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

class EFDContribuicoes:
    
    def __init__(self, excel_path:list) -> None:
        self.excel_path = excel_path
        self.pdfs_path = []
        self.index, self.txt_path = self.import_index_path_from_excel()
        self._process()
        
    def _process(self) -> None:
        self.open_efd()
        
        for path in self.txt_path:
            self.import_shortcut()
            self.import_file(path)
            self.close_pop_ups_download()
            self.download_file(path)
            self.close_escritura()
            self.delete_efd_from_db()
            self.close_pop_ups_delete()
            
        self.editing_the_paths_in_excel(self.index, self.pdfs_path)
        sleep(5)
    
    def import_index_path_from_excel(self) -> list:
        index_data = import_column_from_xlsx(
            excel_path=self.excel_path,
            linha=3,
            coluna='D',
            index=True
        )
        return self.split_index_and_data(index_data)
        
    def split_index_and_data(self, index_data:list) -> list:
        index, data_value = zip(*[(index, value) for index, value in index_data])
        return list(index), list(data_value)

    def editing_the_paths_in_excel(self, index_list:list, pdfs_path:list) -> None:
        for index, path in zip(index_list, pdfs_path):
            editing_xlsx(
                excel_path=self.excel_path,
                data=[path],
                linha=index,
                coluna='D',
                hyperlink=True
            )
    
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
                raise ValueError(f"Erro: {path} não é um arquivo .txt")
        else:
            raise FileNotFoundError(f"Erro: {path} não é um arquivo válido. Verifique se ele está aberto.")
    
    def close_pop_ups_download(self) -> None:
        try:
            pyautogui.locateCenterOnScreen(
                image='refer_images/EFD_Contribuicoes/ArquivoAvisos.png'
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
        download_path = path[:-3] + "pdf"
        
        for i in range(7):
            pyautogui.press('tab')
            
        pyautogui.press('enter')
        
        sleep(3)
        pyautogui.write(download_path)
        pyautogui.press('enter')
        
        logging.info('Baixando pdf')
        self.pdfs_path.append(download_path)
    
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
        close_button = pyautogui.locateCenterOnScreen(
            image='refer_images/EFD_Contribuicoes/CloseButton.png',
            minSearchTime=5
        )
        
        pyautogui.moveTo(close_button)
        pyautogui.click()
        
        logging.info('Processo encerrado')

if __name__ in "__main__":
    EFDContribuicoes(
        'C:\Relatorio\GEM-SHIPPING-LTDA\\05.2024-Sped_Contribuicao.txt',
        'C:\Relatorio\LUZ-PUBLICIDADE-LTDA\-05.2024-Sped_Contribuicao.txt'
    )