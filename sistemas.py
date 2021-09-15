from datetime import time

from selenium.webdriver.common import by
import dados, time
from diretorio import Dir_analise as dir
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By

class Fluig:

    def __init__(self):
        self.__driver = webdriver.Chrome(
            executable_path= dir().dir_base() + dados.Sistemas['webdriver_config']['executable_path']
        )
        self.html_fluig = dados.Sistemas['fluig']

    def get_logged(self):
        self.__driver.get(self.html_fluig['urls']['url_login'])
        try:
            self.__driver.implicitly_wait(30)
            self.__driver.find_element(By.ID, self.html_fluig['html_elements']['html_element_login'])
        except Exception:
            print('Não consegui')
        else:
            self.__driver.find_element(By.ID, self.html_fluig['html_elements']['html_element_login']).send_keys('dhiego.braga@manserv.com.br')
            self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_element_password']).send_keys('Db.240996')
            self.__driver.find_element(By.CLASS_NAME,self.html_fluig['html_elements']['html_element_submit_bt']).click()
        try:
            self.__driver.implicitly_wait(30)
            self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_procced_bt'])
        except Exception:
            print('Cannot get')
        else:     
            self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_procced_bt']).click()

    def get_report(self,url,nomeclatura_ut,codigo_ut):
        self.__driver.get(url)
        try:
           self.__driver.implicitly_wait(30)
           self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_tipo_cb'])
        except Exception:
            print('A')
        else:
            Select(self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_tipo_cb'])).select_by_visible_text('UT')
            try:
                self.__driver.implicitly_wait(30)
                self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_coligada_cb'])
            except Exception:
                print('B')
            else:
                Select(self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_coligada_cb'])).select_by_visible_text(nomeclatura_ut)
                self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_codigo_ut']).send_keys(codigo_ut)
                self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_pesquisar_bt']).click()

    def get_table(self):
        try:    
            self.__driver.implicitly_wait(90)
            self.__driver.find_element(By.XPATH,
            self.html_fluig['html_elements']['html_table_pedido'])
        except Exception:
            count = 1
            tabela_html = ''
            pass
        else:
            elemento_tabela = self.__driver.find_element(By.XPATH,
            self.html_fluig['html_elements']['html_table_pedido'])
            tabela_html = elemento_tabela.get_attribute("outerHTML")
            count = 0
        return tabela_html, count

    def get_hierarquia(self,url,codigo_ut):
        self.__driver.get(url)
        try:
           self.__driver.implicitly_wait(30)
           self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_tipo_cb'])
        except Exception:
            print('A')
        else:
            Select(self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_tipo_cb'])).select_by_visible_text('UT')
            try:
                self.__driver.implicitly_wait(30)
                self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_numero_ut'])
            except Exception:
                print('B')
            else:
                self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_numero_ut']).send_keys(codigo_ut)
                self.__driver.find_element(By.ID,self.html_fluig['html_elements']['html_pesquisar_bt']).click()

    def get_hierarquia_text(self,ut):
        dicionario_hierarquia = {
            "numero_coligada": [],
            "nome_coligada": [],
            "UT": [],
            "nome_ut": [],
            "nome_presidente": [],
            "email_presidente": [],
            "nome_diretorgeral": [],
            "email_diretorgeral": [],
            "nome_diretor": [],
            "email_diretor": [],
            "nome_gerente": [],
            "email_gerente": [],
            "nome_coordenador": [],
            "email_coordenador": [],
        }
        try:
            numero_coligada = self.__driver.find_element(
                By.ID,
                self.html_fluig['html_elements']["hierarquia_info"]["label_coligada"],
            )
            nome_coligada = self.__driver.find_element(
                By.ID,
                self.html_fluig['html_elements']["hierarquia_info"]["label_coligada"],
            )
            nome_ut = self.__driver.find_element(
                By.ID,
                self.html_fluig['html_elements']["hierarquia_info"]["label_nomeclatura_ut"],
            )            
            nome_presidente = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_nome_presidente"],
            )
            email_presidente = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_email_presidente"],
            )
            nome_diretorgeral = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_nome_diretorgeral"],
            )
            email_diretorgeral = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_email_diretorgeral"],
            )
            nome_diretor = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_nome_diretor"],
            )
            email_diretor = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_email_diretor"],
            )
            nome_gerente = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_nome_gerente"],
            )
            email_gerente = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_email_gerente"],
            )
            nome_coordenador = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_nome_coordenador"],
            )
            email_coordenador = self.__driver.find_element(
                By.XPATH,
                self.html_fluig['html_elements']["hierarquia_info"]["label_email_coordenador"],
            )
        except Exception:
            print("Problemas no carregamento da tabela", Exception)
        else:
            dicionario_hierarquia["UT"].append(str(ut))
            dicionario_hierarquia["numero_coligada"].append(
                (numero_coligada.text[0:numero_coligada.text.find(' ')])
            )
            dicionario_hierarquia["nome_coligada"].append(
                nome_coligada.text[(numero_coligada.text.find(' ') + 1):]
            )
            dicionario_hierarquia["nome_ut"].append(
                nome_ut.text[(len(str(ut)) + 1):]
            )    
            dicionario_hierarquia["nome_presidente"].append(
                nome_presidente.text[6:]
            )                               
            dicionario_hierarquia["email_presidente"].append(
                email_presidente.text[8:]
            )
            dicionario_hierarquia["nome_diretorgeral"].append(
                nome_diretorgeral.text[6:]
            )
            dicionario_hierarquia["email_diretorgeral"].append(
                email_diretorgeral.text[8:]
            )
            dicionario_hierarquia["nome_diretor"].append(nome_diretor.text[6:])
            dicionario_hierarquia["email_diretor"].append(
                email_diretor.text[8:]
            )
            dicionario_hierarquia["nome_gerente"].append(nome_gerente.text[6:])
            dicionario_hierarquia["email_gerente"].append(
                email_gerente.text[8:]
            )
            dicionario_hierarquia["nome_coordenador"].append(
                nome_coordenador.text[6:]
            )
            dicionario_hierarquia["email_coordenador"].append(
                email_coordenador.text[8:]
            )
        return dicionario_hierarquia

class Mercado_Eletronico:

    def __init__(self):
        self.__driver = webdriver.Chrome(
            executable_path= dir().dir_base() + dados.Sistemas['webdriver_config']['executable_path']
        )
        self.html_mercado = dados.Sistemas['mercado']

    def get_logged(self):
        self.__driver.get(self.html_mercado['urls']['url_login'])
        try:
            self.__driver.implicitly_wait(30)
            self.__driver.find_element(By.ID,self.html_mercado['html_elements']['txtbox_path_usuario'])
        except Exception:
            print('Não conseguiu')
        else:
            self.__driver.find_element(By.ID, self.html_mercado['html_elements']['txtbox_path_usuario']).send_keys('12062299699')
            self.__driver.find_element(By.ID,self.html_mercado['html_elements']['txtbox_path_senha']).send_keys('Db.240996')
            self.__driver.find_element(By.ID,self.html_mercado['html_elements']['button_path_submit']).click()
                    
    def get_report_page(self):
        self.__driver.get(self.html_mercado['urls']['url_relatorio'])
        try:
            self.__driver.implicitly_wait(30)
            self.__driver.find_element(By.ID,self.html_mercado['html_elements']['link_path_relatorio108'])
        except Exception:
            print('B')
        else:
            self.__driver.find_element(By.ID,self.html_mercado['html_elements']['link_path_relatorio108']).click()

    def get_report_from_ut(self,ut):
        try:
           self.__driver.implicitly_wait(30)
           self.__driver.find_element(By.ID,self.html_mercado['html_elements']['combox_path_centrocusto'])
        except Exception:
            print('A')
        else:
            Select(self.__driver.find_element(By.ID,self.html_mercado['html_elements']['combox_path_centrocusto'])).select_by_visible_text(ut)
            self.__driver.find_element(By.ID,self.html_mercado['html_elements']['checkbox_path_sendemail']).click()
            self.__driver.find_element(By.ID,self.html_mercado['html_elements']['button_path_processar_rel']).click()
            time.sleep(1)
            self.__driver.switch_to.alert.accept()

    def obter_labels_processo(self):
        self.__driver.implicitly_wait(30)
        try:
            self.__driver.find_element(
                By.XPATH,
                self.html_mercado['html_elements']['label_path_statusprocess']
            ).text
        except Exception:
            print('B')
        else:
            status = str(
                self.__driver.find_element(
                    By.XPATH,
                    self.html_mercado['html_elements']['label_path_statusprocess']
                ).text
            )
            num_processo = str(
                self.__driver.find_element(
                    By.XPATH,
                    self.html_mercado['html_elements']['label_path_numprocess']
                ).text
            )
        return status, num_processo

    def atualizar_processo(self):
        self.__driver.implicitly_wait(30)
        time.sleep(2)
        self.__driver.find_element(
            By.ID,
            self.html_mercado['html_elements']['button_path_atual_rel']
        ).click()

    def gerar_relatorio(self):
        self.__driver.implicitly_wait(30)
        time.sleep(2)
        self.__driver.find_element(
            By.XPATH,
            self.html_mercado['html_elements']['button_path_gerar_rel']
        ).click()

    def check_processo(self):
        if self.obter_labels_processo()[0] == "Processado":
            acao = "extrair_relatorio"
            numero_processo = self.obter_labels_processo()[1]
        elif self.obter_labels_processo()[0] == "Erro":
            acao = "refazer_extracao"
            numero_processo = ""
        else:
            acao = "atualizar"
            numero_processo = ""
        return acao, numero_processo
