import pandas as pd
from dados import Sistemas as dados_sistema
from database import Database as db
from diretorio import Dir_analise as Dir 

class Agrupando_Relatorios:

    def __init__(self) -> None:
        self.db_base = pd.DataFrame().from_dict(db().Sharepoint().dicionario_UTs())
    
    def lista_ut_suprimentos(self):
        df_suprimentos = self.db_base
        df_suprimentos = df_suprimentos[df_suprimentos['Relatorio_Suprimentos'] == True]
        return df_suprimentos['UT'].to_list()

    def lista_ut_coligada(self):
        dir_base = Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        df_ut_coligada = pd.read_excel(dir_base + Dir().get_separator() + "DF_Hierarquia" + '.xlsx')
        lista_ut = df_ut_coligada['UT']
        lista_coligada = df_ut_coligada['nome_coligada']
        return lista_ut, lista_coligada

    def DF_Hierarquia(self):
        dir_listado = Dir().dir_makefolder(Dir().dir_sistema("fluig"),'BD_Hierarquia')
        dir_base = Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        lista_caminhos_hierarquia = Dir().list_files(dir_listado)
        dataframe_hierarquia = pd.DataFrame()
        for arquivo in lista_caminhos_hierarquia:
            df_temp = pd.read_excel(dir_listado + Dir().get_separator() + arquivo)
            dataframe_hierarquia = pd.concat([dataframe_hierarquia,df_temp])
        for arquivo in lista_caminhos_hierarquia:
            Dir().remove_file(dir_listado + Dir().get_separator() + arquivo) 
        for coluna in dados_sistema['fluig']['hierarquia_colunas']['colunas_modificar']:
            lista_tratada = []
            lista_nomes = dataframe_hierarquia[coluna].tolist()
            for nome in lista_nomes:
                nome_tratado = nome.title()
                lista_tratada.append(nome_tratado)
            dataframe_hierarquia[coluna] = lista_tratada
        del(dataframe_hierarquia['Unnamed: 0'])
        dataframe_hierarquia.to_excel(dir_base + Dir().get_separator() + "DF_Hierarquia" + '.xlsx')

    def DF_Modulos(self):
        dir_listado = Dir().dir_makefolder(Dir().dir_sistema("fluig"),'BD_Modulos')
        dir_base = Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        lista_caminhos_hierarquia = Dir().list_files(dir_listado)
        dataframe_hierarquia = pd.DataFrame()
        for arquivo in lista_caminhos_hierarquia:
            df_temp = pd.read_excel(dir_listado + Dir().get_separator() + arquivo)
            dataframe_hierarquia = pd.concat([dataframe_hierarquia,df_temp])
        for arquivo in lista_caminhos_hierarquia:
            Dir().remove_file(dir_listado + Dir().get_separator() + arquivo) 
        for coluna in ["Produto"]:
            lista_tratada_produto = []
            lista_tratada_codigo_produto = []
            lista_produto = dataframe_hierarquia[coluna].tolist()
            for produto in lista_produto:
                codigo_produto = produto[0:12]
                produto = produto[14:]
                lista_tratada_codigo_produto.append(codigo_produto)
                lista_tratada_produto.append(produto)
            dataframe_hierarquia["Codigo Produto"] = lista_tratada_codigo_produto
            dataframe_hierarquia["Produto"] = lista_tratada_produto
        for coluna in ["Fornecedor"]:
            lista_tratada_fornecedor = []
            lista_tratada_codigo_fornecedor = []
            lista_fornecedor = dataframe_hierarquia[coluna].tolist()
            for fornecedor in lista_fornecedor:
                codigo_fornecedor = fornecedor[0:18]
                fornecedor = fornecedor[21:]
                lista_tratada_codigo_fornecedor.append(codigo_fornecedor)
                lista_tratada_fornecedor.append(fornecedor)
            dataframe_hierarquia["CNPJ Fornecedor"] = lista_tratada_codigo_fornecedor
            dataframe_hierarquia["Fornecedor"] = lista_tratada_fornecedor
        del(dataframe_hierarquia['Unnamed: 0'])
        dataframe_hierarquia.to_excel(dir_base + Dir().get_separator() + "BD_Modulos" + '.xlsx')

    def DF_Pedidos(self):
        dir_listado = Dir().dir_makefolder(Dir().dir_sistema("fluig"),'BD_Pedidos')
        dir_base = Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        lista_caminhos_hierarquia = Dir().list_files(dir_listado)
        dataframe_hierarquia = pd.DataFrame()
        for arquivo in lista_caminhos_hierarquia:
            df_temp = pd.read_excel(dir_listado + Dir().get_separator() + arquivo)
            dataframe_hierarquia = pd.concat([dataframe_hierarquia,df_temp])
        for arquivo in lista_caminhos_hierarquia:
            Dir().remove_file(dir_listado + Dir().get_separator() + arquivo) 
        for coluna in ["Produto"]:
            lista_tratada_produto = []
            lista_tratada_codigo_produto = []
            lista_produto = dataframe_hierarquia[coluna].tolist()
            for produto in lista_produto:
                codigo_produto = produto[0:11]
                produto = produto[11:]
                lista_tratada_codigo_produto.append(codigo_produto)
                lista_tratada_produto.append(produto)
            dataframe_hierarquia["Codigo Produto"] = lista_tratada_codigo_produto
            dataframe_hierarquia["Produto"] = lista_tratada_produto
        for coluna in ["Fornecedor"]:
            lista_tratada_fornecedor = []
            lista_tratada_codigo_fornecedor = []
            lista_fornecedor = dataframe_hierarquia[coluna].tolist()
            for fornecedor in lista_fornecedor:
                codigo_fornecedor = fornecedor[0:18]
                fornecedor = fornecedor[18:]
                lista_tratada_codigo_fornecedor.append(codigo_fornecedor)
                lista_tratada_fornecedor.append(fornecedor)
            dataframe_hierarquia["CNPJ Fornecedor"] = lista_tratada_codigo_fornecedor
            dataframe_hierarquia["Fornecedor"] = lista_tratada_fornecedor
        del(dataframe_hierarquia['Unnamed: 0'])
        dataframe_hierarquia.to_excel(dir_base + Dir().get_separator() + "BD_Pedidos" + '.xlsx')