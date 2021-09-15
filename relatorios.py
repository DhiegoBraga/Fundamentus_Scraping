from os import cpu_count
from sys import path
import time
from typing import Counter
import zipfile
from bs4 import BeautifulSoup
from numpy import extract
from diretorio import Dir_analise as dir
import pandas as pd
import sistemas, dados, database

class Relatorios_Fluig:

    def __init__(self,LISTA_UT):
        self.driver_fluig = sistemas.Fluig()
        self.html_fluig = dados.Sistemas['fluig']
        self.dir_fluig = dir().dir_sistema('fluig')
        self.lista_codigout = LISTA_UT

    def save_table_pedido(self,lista_nomeclatura,lista_codigout):
        self.driver_fluig.get_logged()
        dir_pedidos = dir().dir_makefolder(self.dir_fluig,'BD_Pedidos')
        for nomeclatura_ut, codigo_ut in zip (lista_nomeclatura, lista_codigout):
            dataframe_geral = pd.DataFrame()
            self.driver_fluig.get_report(self.html_fluig['urls']['url_pedidos'],nomeclatura_ut, codigo_ut)
            if self.driver_fluig.get_table()[1] != 0:
                pass
            else:
                tabela_html = self.driver_fluig.get_table()[0]
                soup = BeautifulSoup(tabela_html,'lxml')
                tabela = soup.find_all("table")
                dataframe = pd.read_html(str(tabela), decimal=",")
                dataframe_filtrado = dataframe[0].dropna(axis=0, thresh=4)
                dataframe_filtrado.insert(0, "Centro Custo", str(codigo_ut))
                dataframe_geral = pd.concat([dataframe_geral, dataframe_filtrado])
                dataframe_geral["Pedido"] = dataframe_geral["Pedido"].astype("int")
                dataframe_geral["Pedido"] = dataframe_geral["Pedido"].astype("str")
                dataframe_geral.to_excel(
                    dir_pedidos + r'\Pedido_UT_' + str(codigo_ut) + '.xlsx'
                )

    def save_table_modulo(self,lista_nomeclatura, lista_codigout):
        self.driver_fluig.get_logged()
        dir_modulos = dir().dir_makefolder(self.dir_fluig,'BD_Modulos')
        for nomeclatura_ut, codigo_ut in zip (lista_nomeclatura, lista_codigout):
            dataframe_geral = pd.DataFrame()
            self.driver_fluig.get_report(self.html_fluig['urls']['url_modulo'],nomeclatura_ut, codigo_ut)
            if self.driver_fluig.get_table()[1] != 0:
                pass
            else:
                tabela_html = self.driver_fluig.get_table()[0]
                soup = BeautifulSoup(tabela_html,'lxml')
                tabela = soup.find_all("table")
                dataframe = pd.read_html(str(tabela), decimal=",")
                dataframe_filtrado = dataframe[0].dropna(axis=0, thresh=4)
                dataframe_filtrado.insert(0, "Centro Custo", str(codigo_ut))
                dataframe_geral = pd.concat([dataframe_geral, dataframe_filtrado])
                dataframe_geral.to_excel(
                    dir_modulos + r'\Modulo_UT_' + str(codigo_ut) + '.xlsx'
                )

    def save_hierarquia(self):
        self.driver_fluig.get_logged()
        dir_hierarquia = dir().dir_makefolder(self.dir_fluig,'BD_Hierarquia')
        for ut in self.lista_codigout:
            self.driver_fluig.get_hierarquia(self.html_fluig['urls']['url_hierarquia'],ut)
            dic = self.driver_fluig.get_hierarquia_text(ut)
            dataframe = pd.DataFrame().from_dict(dic)
            dataframe.to_excel(
                dir_hierarquia + r'\Hierarquia_UT_' + str(ut) + '.xlsx'
            )
            
class Relatorio_Mercado:

    def __init__(self, LISTA_UT):
        self.driver_mercado = sistemas.Mercado_Eletronico()
        self.html_mercado = dados.Sistemas['mercado']
        self.dir_mercado = dir().dir_sistema('mercado')
        self.lista_ut = LISTA_UT

    def save_report(self):
        self.driver_mercado.get_logged()  
        lista_processo = []
        lista_ut_processado = []
        dir_rm = dir().dir_makefolder(
            self.dir_mercado,              
            "BD_RMs")
        for ut in self.lista_ut:
            count_report = 0
            while count_report == 0:
                self.driver_mercado.get_report_page()
                self.driver_mercado.get_report_from_ut(ut)
                count = 0
                while count == 0:
                    status = self.driver_mercado.obter_labels_processo()[0]
                    if status != "Processado":
                        self.driver_mercado.atualizar_processo()
                    elif status == "Erro":
                        count = 1
                    else:
                        self.driver_mercado.gerar_relatorio()
                        lista_processo.append(self.driver_mercado.obter_labels_processo()[1])
                        lista_ut_processado.append(ut)
                        count = 1
                        count_report = 1
        time.sleep(10)
        for processo, ut in zip(lista_processo, lista_ut_processado):
            file_zip = (
                dir().dir_base()
                + dir().get_separator()
                + "Downloads"
                + dir().get_separator()
                + self.html_mercado['Report_info']
                + processo 
                + '.zip'
            )
            dir().extract_file(file_zip,dir_rm)
            dir().remove_file(file_zip)
            old_file = (
                dir_rm
                + dir().get_separator()
                + self.html_mercado['Report_info']
                + processo 
                + '.xlsx'
            )
            new_file = (
                dir_rm
                + dir().get_separator()
                + "RMs_UT_"
                + ut 
                + '.xlsx'
            )
            dir().rename_file(old_file,new_file)
        return lista_processo, lista_ut_processado
