import time
import pandas as pd
from diretorio import dir_dict, chrome_driver_path
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By

class Fundamental:

    def __init__(self) -> None:
        self.__driver = webdriver.Chrome(
            executable_path = chrome_driver_path + r'\chromedriver.exe'
        )

    def get_page(self,url):
        self.__driver.get(url)

    def get_table(self):
        try:
            time.sleep(10)
            self.__driver.implicitly_wait(30)
            elemento_tabela = self.__driver.find_element(
                By.ID,
                'resultado'
            )
        except Exception:
            print("Problemas no carregamento da tabela", Exception)
        else:
            tabela_html = elemento_tabela.get_attribute("outerHTML")
            soup = BeautifulSoup(tabela_html, "lxml")
            tabela = soup.find_all("table")
            dataframe = pd.read_html(str(tabela), decimal=",")
            dataframe_filtrado = dataframe[0].dropna(axis=0, thresh=4)
            print(dataframe_filtrado)
            dataframe_filtrado.to_excel(dir_dict['dir_base_documentos'] + r'\Cotacao_fundamentals.xlsx')

    def teste(self):
        self.get_page('https://fundamentus.com.br/resultado.php')
        self.get_table()

Fundamental().teste()
