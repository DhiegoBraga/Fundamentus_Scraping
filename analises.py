from numpy.core.numeric import NaN
from numpy.lib.arraysetops import isin
import pandas as pd
import relatorios
import time
from datetime import datetime
from dados import Sistemas as dados_sistema
from database import Database as db
from diretorio import Dir_analise as Dir 

class Configs:

    def __init__(self):
        self.Configs_Data = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Configs')
            + Dir().get_separator() + "Configs_Data" + '.xlsx')
    def Data_UT(self):
        BD_config = self.Configs_Data[self.Configs_Data['Status'] == 'Ativo']
        ut = BD_config['UT'].tolist()
        primeira_data = BD_config['Primeira Data'].astype(str).tolist()
        ultima_data = BD_config['Ultima Data'].astype(str).tolist()
        return ut,primeira_data,ultima_data

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

    def DF_Justificativa(self):
        dir_base = Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        df_bdrm = pd.read_excel(dir_base + Dir().get_separator() + "BD_RMs" + '.xlsx') 
        dir_justificativa = (dir_base + Dir().get_separator() + "BD_Justificativa" + '.xlsx')   
        if Dir().check_file(dir_justificativa) == False:
            dict_vazio = {
                'RM': [],
                'Justificativa':[],
                'N_Prisma':[]
            }
            empty_dfjust = pd.DataFrame.from_dict(dict_vazio)
            empty_dfjust.to_excel(dir_justificativa)
        else:
            pass
        count = 1
        while count != 0:
            df_justificativa = pd.read_excel(dir_justificativa)
            lista_justificativa = df_justificativa['RM'].tolist()
            lista_bdrm = df_bdrm['Requisição'].drop_duplicates().tolist()
            lista_fazer = []
            for item in lista_bdrm:
                if item in lista_justificativa:
                    pass
                else:
                    print(item)
                    lista_fazer.append(str(item))
            dict_just = relatorios.Relatorio_Mercado('','','').save_justificativa(lista_fazer)
            temp_df = pd.DataFrame.from_dict(dict_just)
            df_justificativa = pd.concat([df_justificativa,temp_df])
            df_justificativa.to_excel(dir_justificativa)
            time.sleep(5)
            lista_justificativa = df_justificativa['RM'].tolist()
            lista_bdrm = df_bdrm['Requisição'].tolist()
            for item in lista_justificativa:
                if item in lista_bdrm:
                    pass
                else:
                    lista_fazer.append(str(item))
            count = len(lista_fazer)
            

    def DF_RMs_Mercado(self):
        dir_listado = Dir().dir_makefolder(Dir().dir_sistema("mercado"),'BD_RMs')
        dir_base = Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        lista_caminho_rms = Dir().list_files(dir_listado)
        dataframe_rm = pd.DataFrame()
        for arquivo in lista_caminho_rms:
            df_temp = pd.read_excel(dir_listado + Dir().get_separator() + arquivo,1)
            dataframe_rm = pd.concat([dataframe_rm,df_temp])
        for coluna in dados_sistema['mercado']['colunas_a_deletar']:
            del(dataframe_rm[coluna])
        dataframe_rm['Andamento Req.'] = dataframe_rm['Andamento Req.'].replace({
                'RFQ':'Em Cotação',
                'Não Enviado':'Pendencia de Compra'})
        def status_geral_rm():
            lista_tratada = []
            lista_KPI = []
            lista_status_rm = dataframe_rm['Status Requisição'].tolist()
            lista_status_pedido = dataframe_rm['Andamento Req.'].tolist()
            for rm,pedido in zip(lista_status_rm,lista_status_pedido):
                if rm == 'CANCELADO':
                    lista_tratada.append('Finalizado')
                    lista_KPI.append(0)
                elif rm == 'EM PROCESSAMENTO':
                    lista_tratada.append(f'Pendente (Aprovação)')
                    lista_KPI.append(1)
                elif rm == 'APROVADO' and pedido == 'Pedido':
                    lista_tratada.append('Concluido')
                    lista_KPI.append(0)
                elif rm == 'APROVADO' and pedido == 'Cancelado':
                    lista_tratada.append('Finalizado')
                    lista_KPI.append(0)
                elif rm != 'APROVADO':
                    lista_tratada.append('Divergência')
                    lista_KPI.append(1) 
                else:
                    lista_tratada.append(f'Pendente ({pedido})')
                    lista_KPI.append(1)         
            return lista_tratada, lista_KPI
        status = status_geral_rm()
        dataframe_rm['Resumo andamento']  = status[0]
        dataframe_rm['Resumo andamento KPI']  = status[1]

        dataframe_rm.to_excel(dir_base + Dir().get_separator() + "BD_RMs" + '.xlsx')

    def DF_NFs_TOTVS(self):
        dir_listado = Dir().dir_makefolder(Dir().dir_sistema("totvs"),'BD_NF_UTs')
        dir_base = Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        lista_caminho_nfs = Dir().list_files(dir_listado)
        dataframe_nf = pd.DataFrame()
        for arquivo in lista_caminho_nfs:
            df_temp = pd.read_excel(dir_listado + Dir().get_separator() + arquivo)
            dataframe_nf = pd.concat([dataframe_nf,df_temp])
        for coluna in dados_sistema['totvs']['colunas_a_deletar']:
            del(dataframe_nf[coluna])
        dataframe_nf.to_excel(dir_base + Dir().get_separator() + "BD_NFs" + '.xlsx')

class Limpando_BDs:

    def __init__(self):
        self.BD_NF_Bruto = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
            + Dir().get_separator() + "BD_NFs" + '.xlsx')
        self.BD_NF_Manual = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("totvs"),'BD_TOTVS_MANUAL')
            + Dir().get_separator() + "TOTVS_MANUAL" + '.xlsx')

    def comparando_bds(self):
        BD_NF_Bruto = self.BD_NF_Bruto
        BD_NF_Manual = self.BD_NF_Manual
        BD_NF_Bruto['Cod_unico'] = (
            BD_NF_Bruto['Número do Documento'].astype(str) 
            + BD_NF_Bruto['Cliente/Fornecedor'].astype(str)
            + BD_NF_Bruto['Centro de Custo'].astype(str))
        BD_NF_Manual['Cod_unico'] = (
            BD_NF_Manual['Número do Documento'].astype(str) 
            + BD_NF_Manual['Cliente/Fornecedor'].astype(str)
            + BD_NF_Manual['Centro de Custo'].astype(str))
        lista_status = []
        lista_NFs_brutos = BD_NF_Bruto['Cod_unico'].astype(str).tolist()
        lista_NFs_manual = BD_NF_Manual['Cod_unico'].astype(str).tolist()
        for nf in lista_NFs_brutos:
            if nf in lista_NFs_manual:
                lista_status.append('Tratado')
            else:
                lista_status.append('Nao tratado')
        BD_NF_Bruto['Status Tratativa'] = lista_status
        BD_NF_Bruto = BD_NF_Bruto[BD_NF_Bruto['Status Tratativa'] == 'Nao tratado']
        return BD_NF_Bruto
    
    def Desfragmentando_campo_historico(self):
        BD_NF_Bruto = self.comparando_bds()
        BD_NF_Bruto['Histórico'] = BD_NF_Bruto['Histórico'].fillna('-')
        lista_historico = BD_NF_Bruto['Histórico'].tolist()
        lista_categoria = []
        lista_n_pedido = []
        lista_n_solicitacao = []
        lista_usuario_email = []
        for historico in lista_historico:
            if historico.find('Previsão de Compras') >= 0:
                lista_categoria.append('Previsão de Compras')
                lista_n_pedido.append('-')
                lista_n_solicitacao.append('-')
                lista_usuario_email.append('-')
            elif historico.find('PED:') >= 0:
                lista_categoria.append('Pedido Mercado')
                lista_n_pedido.append(historico[
                    (historico.find('PED:') + 5):
                    ((historico.find('PED:')) + 12)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('PC -') >= 0:
                lista_categoria.append('Pedido')
                lista_n_pedido.append(historico[
                    (historico.find('PC -') + 5):
                    ((historico.find('PC -')) + 12)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('PC:') >= 0:
                lista_categoria.append('Pedido')
                lista_n_pedido.append(historico[
                    (historico.find('PC:') + 4):
                    ((historico.find('PC:')) + 11)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('PC-') >= 0:
                lista_categoria.append('Pedido')
                lista_n_pedido.append(historico[
                    (historico.find('PC-') + 4):
                    ((historico.find('PC-')) + 11)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('PC ') >= 0:
                lista_categoria.append('Pedido')
                lista_n_pedido.append(historico[
                    (historico.find('PC ') + 4):
                    ((historico.find('PC ')) + 11)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('Pc ') >= 0:
                lista_categoria.append('Pedido')
                lista_n_pedido.append(historico[
                    (historico.find('Pc ') + 4):
                    ((historico.find('Pc ')) + 11)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('Pedido Compra Dir. n.') >= 0:
                lista_categoria.append('Pedido Mercado')
                lista_n_pedido.append(historico[
                    (historico.find('Pedido Compra Dir. n.') + 21):
                    ((historico.find('PED:')) + 27)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('PPC:') >= 0:
                lista_categoria.append('Pedido Modulo')
                lista_n_pedido.append(historico[
                    (historico.find('PPC:') + 5):
                    ((historico.find('PPC:')) + 12)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('PPC -') >= 0:
                lista_categoria.append('Pedido Modulo')
                lista_n_pedido.append(historico[
                    (historico.find('PPC -') + 6):
                    ((historico.find('PPC -')) + 13)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('MD ') >= 0:
                lista_categoria.append('Pedido Modulo')
                lista_n_pedido.append(historico[
                    (historico.find('MD ') + 4):
                    ((historico.find('MD ')) + 11)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('ACD:') >= 0:
                lista_categoria.append('Pedido ACD')
                lista_n_pedido.append(historico[
                    (historico.find('ACD:') + 5):
                    ((historico.find('ACD:')) + 12)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            elif historico.find('ACD') >= 0:
                lista_categoria.append('Pedido ACD')
                lista_n_pedido.append(historico[
                    (historico.find('ACD') + 4):
                    ((historico.find('ACD')) + 11)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])

            elif historico.find('Pedido de Contrato n.') >= 0:
                lista_categoria.append('Pedido Modulo')
                lista_n_pedido.append(historico[
                    (historico.find('Pedido de Contrato n.') + 22):
                    ((historico.find('Pedido de Contrato n.')) + 28)
                ])
                lista_n_solicitacao.append(historico[
                    (historico.find('Número Fluig:') + 14):
                    ((historico.find('Número Fluig:')) + 21)
                ])
                lista_usuario_email.append(historico[
                    (historico.find('E-mail:') + 7):])
            
            elif historico.find('Prest. de Contas') >= 0 or historico.find('Prest.Contas') >= 0:
                lista_categoria.append('Prest. de Contas')
                lista_n_pedido.append('-')
                lista_n_solicitacao.append('-')
                lista_usuario_email.append('-')
            elif historico.find('Tributação referente a lançamento Ref') >= 0:
                lista_categoria.append('Tributação')
                lista_n_pedido.append('-')
                lista_n_solicitacao.append('-')
                lista_usuario_email.append('-')
            elif historico.find('GRFC') >= 0:
                lista_categoria.append('GRFC')
                lista_n_pedido.append('-')
                lista_n_solicitacao.append('-')
                lista_usuario_email.append('-')
            elif historico.find('GRRF') >= 0:
                lista_categoria.append('GRRF')
                lista_n_pedido.append('-')
                lista_n_solicitacao.append('-')
                lista_usuario_email.append('-')                                
            else:
                lista_categoria.append('-')
                lista_n_pedido.append('-')
                lista_n_solicitacao.append('-')
                lista_usuario_email.append('-')
        return lista_categoria,lista_n_pedido,lista_n_solicitacao,lista_usuario_email

    def save(self):
        DF_TOTVS = self.comparando_bds()
        DF_TOTVS['Categoria'] = self.Desfragmentando_campo_historico()[0]
        DF_TOTVS['N Pedido'] = self.Desfragmentando_campo_historico()[1]
        DF_TOTVS['N Solicitação'] = self.Desfragmentando_campo_historico()[2]
        DF_TOTVS['Email Solicitante'] = self.Desfragmentando_campo_historico()[3]
        DF_TOTVS.to_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
            + Dir().get_separator() + "BD_NFs_Pendente" + '.xlsx')

class Relatorios_frequencia:

    def __init__(self):
        self.dir_base_freq = Dir().dir_makefolder(Dir().dir_sistema("BD_Frequencia"),'Suprimentos') 
        self.BD_RM_Base = pd.read_excel(
            (Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos'))
            + Dir().get_separator() + "BD_RMs" + '.xlsx',index_col=0)
        self.BD_Pedido = pd.read_excel(
            (Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos'))
            +  Dir().get_separator() + "BD_Pedidos" + '.xlsx',index_col=0)
        self.BD_NF_Manual = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("totvs"),'BD_TOTVS_MANUAL')
            + Dir().get_separator() + "TOTVS_MANUAL" + '.xlsx')
        '''
        self.BD_Modulo = pd.read_excel(
            (Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos'))
            + Dir().get_separator() + "BD_Modulos" + '.xlsx',index_col=0)
        '''
        self.BD_hierarquia = pd.read_excel(
            (Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos'))
            + Dir().get_separator() + "DF_Hierarquia" + '.xlsx',index_col=0)

    def resumo_status_rm(self):
        BD_RM = self.BD_RM_Base
        BD_pedido_resumo = self.linkar_pedido()
        BD_frequencia_resumokpi = BD_RM.groupby(['Requisição']).agg(
            {
            'Resumo andamento KPI':'sum',
            'Preço Total Requisição':'sum',
            'Preço Total Pedido':'sum'
            }).reset_index()
        BD_frequencia_pedidos = BD_RM.groupby(['C. Custo','Requisição'])['Pedido Cliente'].nunique().reset_index() 
        BD_frequencia_resumo_andamentos = BD_RM.groupby(['C. Custo','Requisição'])['Resumo andamento'].nunique().reset_index() 
        BD_frequencia_resumo_andamentos_lista = BD_RM.groupby(['Requisição','Resumo andamento'])['Item Requisição'].count().reset_index() 
        del(BD_frequencia_pedidos['C. Custo'])
        BD_frequencia_rm = BD_frequencia_pedidos.merge(BD_frequencia_resumo_andamentos,how='left',on='Requisição')
        BD_frequencia_rm = BD_frequencia_rm.merge(BD_frequencia_resumokpi,how='left',on='Requisição')
        lista_status = []
        for rm,resumo_andamento, resumo_kpi in zip(BD_frequencia_rm['Requisição'],BD_frequencia_rm['Resumo andamento'],BD_frequencia_rm['Resumo andamento KPI']):
            if int(resumo_andamento) == 1 and int(resumo_kpi) == 0:
                for rm_resumo, resumo in zip(BD_frequencia_resumo_andamentos_lista['Requisição'],BD_frequencia_resumo_andamentos_lista['Resumo andamento']):
                    if rm == rm_resumo:
                        lista_status.append(resumo)
                    else:
                        pass
            elif int(resumo_andamento) == 2 and int(resumo_kpi) == 0:
                lista_status.append('Concluido')
            else:
                lista_status.append('Pendente')
        BD_frequencia_rm['Status Processo'] = lista_status
        BD_frequencia_rm['Data_Status'] = datetime.today().strftime('%d/%m/%Y')
        BD_frequencia_rm = BD_frequencia_rm.merge(self.frequencia_notafiscal(),how='left',on='Requisição')
        BD_frequencia_rm = BD_frequencia_rm.rename(columns={'Valor Original':'Valor Total NFs'})
        BD_frequencia_rm['Valor Total NFs'] = BD_frequencia_rm['Valor Total NFs'].fillna('-')
        lista_status_pedido = []
        lista_status_pedido_kpi = []
        lista_status_nf = []
        lista_status_nf_kpi = []       
        lista_obs = [] 
        lista_processo_kpi = []
        for rm,status_processo,pedido_cliente,valor_pedido,valor_nfs in zip(BD_frequencia_rm['Requisição'].tolist(),BD_frequencia_rm['Status Processo'].tolist(),BD_frequencia_rm['Pedido Cliente'].tolist(),
        BD_frequencia_rm['Preço Total Pedido'].tolist(),
        BD_frequencia_rm['Valor Total NFs'].tolist()):
            if status_processo == 'Finalizado' and valor_nfs == '-':
                lista_processo_kpi.append(-1000)
                lista_status_pedido.append('Sem Informação')
                lista_status_pedido_kpi.append(-1000)
                lista_status_nf.append('Sem Informação')
                lista_status_nf_kpi.append(-1000)                
                lista_obs.append("RM cancelada, não gerou pedido de compra e faturamento.")
            elif status_processo == 'Pendente' and pedido_cliente == 0 and valor_nfs == '-':
                lista_processo_kpi.append(10)
                lista_status_pedido.append('Sem Informação')
                lista_status_pedido_kpi.append(-1000)
                lista_status_nf.append('Sem Informação')
                lista_status_nf_kpi.append(-1000)                
                lista_obs.append("RM divergente, não gerou pedido de compra e faturamento.")           
            elif status_processo == 'Concluido' and valor_nfs == '-':
                lista_processo_kpi.append(0)
                BD_status = BD_pedido_resumo[BD_pedido_resumo['Requisição'] == rm]['Status']
                BD_status = BD_status.fillna('-')
                if 'RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançamento Efetuado'
                    status_kpi = 0
                elif 'PENDENTE' in (BD_status.tolist()):
                    status = 'Lançamento Pendente'
                    status_kpi = 10
                elif 'PARC RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançado Parcialmente'
                    status_kpi = 5
                else:
                    status = 'Sem informação'
                    status_kpi = -1000                    
                lista_status_pedido.append(status)
                lista_status_pedido_kpi.append(status_kpi)
                lista_status_nf.append('Sem Informação')
                lista_status_nf_kpi.append(-1000)        
                if pedido_cliente == 1:      
                    lista_obs.append("RM Aprovada, 1 pedido de compra gerado, faturamento não vinculado.")   
                else:
                    lista_obs.append(f"RM Aprovada, {str(pedido_cliente)} pedidos de compras gerados, faturamento não vinculado.")       
            elif status_processo == 'Concluido'and valor_nfs != '-':
                lista_processo_kpi.append(0)
                BD_status = BD_pedido_resumo[BD_pedido_resumo['Requisição'] == rm]['Status']
                BD_status = BD_status.fillna('-')
                if 'RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançamento Efetuado'
                    status_kpi = 0
                elif 'PENDENTE' in (BD_status.tolist()):
                    status = 'Lançamento Pendente'
                    status_kpi = 10
                elif 'PARC RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançado Parcialmente'
                    status_kpi = 5
                else:
                    status = 'Sem informação'
                    status_kpi = -1000                    
                lista_status_pedido.append(status)
                lista_status_pedido_kpi.append(status_kpi)
                if valor_pedido == valor_nfs:
                    status_fat = 'Faturamento Integral'
                    status_fat_kpi = 0
                elif valor_pedido > valor_nfs:
                    status_fat = 'Faturamento Parcial'
                    status_fat_kpi = 5
                else:
                    status_fat = 'Faturamento a maior'   
                    status_fat_kpi = 10                 
                lista_status_nf.append(status_fat)
                lista_status_nf_kpi.append(status_fat_kpi)        
                if pedido_cliente == 1:      
                    lista_obs.append("RM Aprovada, 1 pedido de compra gerado, faturamento vinculado.")   
                else:
                    lista_obs.append(f"RM Aprovada, {str(pedido_cliente)} pedidos de compras gerados, faturamento vinculado.")   

            elif status_processo == 'Pendente'and valor_nfs == '-':
                lista_processo_kpi.append(10)
                BD_status = BD_pedido_resumo[BD_pedido_resumo['Requisição'] == rm]['Status']
                BD_status = BD_status.fillna('-')
                if 'RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançamento Efetuado'
                    status_kpi = 0
                elif 'PENDENTE' in (BD_status.tolist()):
                    status = 'Lançamento Pendente'
                    status_kpi = 10
                elif 'PARC RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançado Parcialmente'
                    status_kpi = 5
                else:
                    status = 'Sem informação'
                    status_kpi = -1000                    
                lista_status_pedido.append(status)
                lista_status_pedido_kpi.append(status_kpi)
                lista_status_nf.append('Sem Informação')
                lista_status_nf_kpi.append(-1000)         
                lista_obs.append("RM Aprovada, Pendente tratativa")   
            elif status_processo == 'Pendente'and valor_nfs != '-':
                lista_processo_kpi.append(10)
                BD_status = BD_pedido_resumo[BD_pedido_resumo['Requisição'] == rm]['Status']
                BD_status = BD_status.fillna('-')
                if 'RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançamento Efetuado'
                    status_kpi = 0
                elif 'PENDENTE' in (BD_status.tolist()):
                    status = 'Lançamento Pendente'
                    status_kpi = 10
                elif 'PARC RECEBIDO' in (BD_status.tolist()):
                    status = 'Lançado Parcialmente'
                    status_kpi = 5
                else:
                    status = 'Sem informação'
                    status_kpi = -1000                    
                lista_status_pedido.append(status)
                lista_status_pedido_kpi.append(status_kpi)
                if int(valor_pedido) == int(valor_nfs):
                    status_fat = 'Faturamento Integral'
                    status_fat_kpi = 0
                elif int(valor_pedido) > int(valor_nfs):
                    status_fat = 'Faturamento Parcial'
                    status_fat_kpi = 5
                else:
                    status_fat = 'Faturamento a maior'   
                    status_fat_kpi = 10                 
                lista_status_nf.append(status_fat)
                lista_status_nf_kpi.append(status_fat_kpi)          
                lista_obs.append("Pendencia no processo")                                               
            else:
                lista_processo_kpi.append(-1000)
                lista_status_pedido.append("")
                lista_status_pedido_kpi.append(0)
                lista_status_nf.append('')
                lista_status_nf_kpi.append(0)   
                lista_obs.append("")                     
        BD_frequencia_rm['Status Processamento KPI'] = lista_processo_kpi
        BD_frequencia_rm['Status Pedido'] = lista_status_pedido
        BD_frequencia_rm['Status Pedido KPI'] = lista_status_pedido_kpi
        BD_frequencia_rm['Status Faturamento'] = lista_status_nf
        BD_frequencia_rm['Status Faturamento KPI'] = lista_status_nf_kpi
        BD_frequencia_rm['Observação'] = lista_obs      
        lista_status_geral = []
        for status_processo,status_pedido,status_fat in zip (BD_frequencia_rm['Status Processo'].tolist(),BD_frequencia_rm['Status Pedido'].tolist(),BD_frequencia_rm['Status Faturamento'].tolist()):
            if status_processo == 'Finalizado':
                lista_status_geral.append("Finalizado - 1ª Fase")
            elif status_processo == 'Pendente':
                lista_status_geral.append("Pendente - 1ª Fase")
            else:
                if status_pedido == 'Sem Informação' and status_pedido == 'Sem Informação':
                    lista_status_geral.append("Pendente - 2ª Fase (Informação)")
                elif status_pedido == "Lançamento Efetuado"	and status_pedido == 'Faturamento Integral':
                    lista_status_geral.append("Concluido - 3ª Fase")
                elif status_pedido == "Lançamento Efetuado"	and status_pedido == 'Faturamento a maior':
                    lista_status_geral.append("Concluido - 3ª Fase")
                elif status_pedido == "Lançamento Efetuado"	and status_pedido == 'Faturamento Parcial':
                    lista_status_geral.append("Pendente - 3ª Fase")
                elif status_pedido == "Lançamento Efetuado"	and status_pedido == 'Sem Informação':
                    lista_status_geral.append("Pendente - 3ª Fase (Informação)")                    
                else:
                    lista_status_geral.append("Pendente - 2ª Fase")
        BD_frequencia_rm['Status Geral'] = lista_status_geral
        BD_frequencia_rm.to_excel(self.dir_base_freq + Dir().get_separator() + "BD_Resumo_StatusGeral" + '.xlsx')

    def frequencia_RM_Pedido(self):
        def limpar_pontuacao_cc(lista):
            lista_tratada = []
            for item in lista:
                novo = item.replace('.','')
                lista_tratada.append(novo)
            return lista_tratada
        def concatenando_valores(lista_cc,lista_cnpj,lista_ped):
            lista_tratada = []
            for cc, cnpj, ped in zip(lista_cc,lista_cnpj,lista_ped):
                novo = cc+cnpj+ped
                lista_tratada.append(novo)
            return lista_tratada

        BD_RM = self.BD_RM_Base
        BD_RM['Pedido Cliente'] = BD_RM['Pedido Cliente'].fillna(0)
        BD_RM['Pedido Cliente'] = BD_RM['Pedido Cliente'].astype(int)
        BD_RM['Pedido Cliente'] = BD_RM['Pedido Cliente'].astype(str)
        BD_RM['Pré-Pedido'] = BD_RM['Pré-Pedido'].fillna(0)
        BD_RM['Pré-Pedido'] = BD_RM['Pré-Pedido'].astype(int)
        BD_RM['Pré-Pedido'] = BD_RM['Pré-Pedido'].astype(str)      
        BD_RM['Pré-Pedido'] = BD_RM['Pré-Pedido'].replace(
            ['0'],'Sem numero mercado' )        
        BD_RM['Pedido Cliente'] = BD_RM['Pedido Cliente'].replace(
            ['0'],'Sem numero pedido' )
        BD_frequencia = BD_RM.groupby(['Requisição', 'Pedido Cliente','C. Custo','Pré-Pedido','Categoria da Requisição']).agg(
            {
                'Qtde. Pedido':'count',
                'Preço Total Requisição':'sum',
                'Preço Total Pedido':'sum'
            }
        ).reset_index()
        BD_frequencia['Codigo Unico'] = concatenando_valores(
            limpar_pontuacao_cc(BD_frequencia['C. Custo'].astype(str).tolist()),
            BD_frequencia['Pré-Pedido'].astype(str).tolist(),
            BD_frequencia['Pedido Cliente'].astype(str)
        )
        return BD_frequencia
        
    def frequencia_pedido(self):

        def limpar_pontuacao_cc(lista):
            lista_tratada = []
            for item in lista:
                novo = item.replace('.','')
                lista_tratada.append(novo)
            return lista_tratada

        def limpar_pontuacao_cnpj(lista):
            lista_tratada = []
            for item in lista:
                novo = item.replace('.','').replace('/','').replace('-','')
                lista_tratada.append(novo)
            return lista_tratada

        def concatenando_valores(lista_cc,lista_cnpj,lista_ped):
            lista_tratada = []
            for cc, cnpj, ped in zip(lista_cc,lista_cnpj,lista_ped):
                novo = cc+cnpj+ped
                lista_tratada.append(novo)
            return lista_tratada
        
        BD_Pedido = self.BD_Pedido
        BD_Pedido['Pedido Cliente'] = BD_Pedido['Pedido'].astype(int)
        BD_Pedido['Pedido Cliente'] = BD_Pedido['Pedido Cliente'].astype(str)
        BD_Pedido['Pré-Pedido'] = BD_Pedido['Nr. ME'].fillna(0)
        BD_Pedido['Pré-Pedido'] = BD_Pedido['Pré-Pedido'].astype(int)
        BD_Pedido['Pré-Pedido'] = BD_Pedido['Pré-Pedido'].astype(str)        
        BD_Pedido['Pré-Pedido'] = BD_Pedido['Pré-Pedido'].replace(
            ['0'],'Sem numero pré-pedido' )               
        BD_frequencia = BD_Pedido.groupby(['Centro Custo','Pedido','Status','CNPJ Fornecedor','Pré-Pedido','Fornecedor']).agg(
            {
                'Produto':'count'
            }
        ).reset_index()
        print(BD_frequencia)
        BD_frequencia['Codigo Unico'] = concatenando_valores(
            limpar_pontuacao_cc(BD_frequencia['Centro Custo'].astype(str).tolist()),
            BD_frequencia['Pré-Pedido'].astype(str).tolist(),
            BD_frequencia['Pedido'].astype(str)
        )
        #BD_frequencia.to_excel(self.dir_base_freq + Dir().get_separator() + "BD_Freq_Pedidos_teste" + '.xlsx')
        return BD_frequencia
        
    def resumo_pedido(self,BD_frequencia):
        lista_status_geral = []
        lista_status_geral_kpi = []
        BD_frequencia['Status'] =BD_frequencia['Status'].fillna('')
        for item in BD_frequencia['Status'].tolist():
            if item == 'RECEBIDO':
                lista_status_geral_kpi.append(0)
            elif item == '':
                lista_status_geral_kpi.append(-1000)
            else:
                lista_status_geral_kpi.append(1)
        BD_frequencia['Status KPI'] = lista_status_geral_kpi
        BD_frequencia = BD_frequencia.groupby(['Requisição'])['Status KPI'].sum().reset_index()
        for item in BD_frequencia['Status KPI'].tolist():
            if item == 0:
                lista_status_geral.append('Concluido')
            elif item < 0:
                lista_status_geral.append('Sem Informação')                
            else:
                lista_status_geral.append('Pendente')
        BD_frequencia['Status'] = lista_status_geral
        return BD_frequencia

    def linkar_pedido(self):
        BD_Frequencia_RM = self.frequencia_RM_Pedido()
        BD_Frequencia_Pedido = self.frequencia_pedido()
        for coluna in ['Centro Custo','Pedido','Pré-Pedido','Produto']:
            del BD_Frequencia_Pedido[coluna]
        BD_frequencia_rm = BD_Frequencia_RM.merge(BD_Frequencia_Pedido,how='left',on='Codigo Unico')
        BD_resumo_pedido_rm = self.resumo_pedido(BD_frequencia_rm)
        BD_resumo_pedido_rm.to_excel(self.dir_base_freq + Dir().get_separator() + "BD_Freq_Pedidos_RM" + '.xlsx')
        BD_frequencia_rm.to_excel(self.dir_base_freq + Dir().get_separator() + "BD_Freq_Pedidos" + '.xlsx')
        return BD_frequencia_rm

    def frequencia_notafiscal(self):
        BD_NF = self.BD_NF_Manual
        DB_frequ_NF = BD_NF.groupby(['Requisição'])['Valor Original'].sum().reset_index() 
        return DB_frequ_NF
        #DB_frequ_NF.to_excel(self.dir_base_freq + Dir().get_separator() + "BD_Freq_RMxNF" + '.xlsx')

class Linkando_relatorios:

    def __init__(self):

        self.BD_RM_Base = pd.read_excel(
            (Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos'))
            + Dir().get_separator() + "BD_RMs" + '.xlsx',index_col=0)
        self.BD_Pedido = pd.read_excel(
            (Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos'))
            +  Dir().get_separator() + "BD_Pedidos" + '.xlsx',index_col=0)
        self.BD_TOTVS_MANUAL  = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("totvs"),'BD_TOTVS_MANUAL')
            + Dir().get_separator() + "TOTVS_MANUAL" + '.xlsx')
        self.BD_NF_Base = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        + Dir().get_separator() + "BD_NFs_Base" + '.xlsx')
        self.BD_RM_FREQ = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Frequencia"),'Suprimentos')
            + Dir().get_separator() + "BD_Freq_RM_Pedidos" + '.xlsx')
        self.BD_PED_FREQ = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Frequencia"),'Suprimentos')
            + Dir().get_separator() + "BD_Freq_Pedidos" + '.xlsx')
        
    def Codigo_unico_totvs(self):
        def limpar_pontuacao_cc(lista):
            lista_tratada = []
            for item in lista:
                novo = item.replace('.','')
                lista_tratada.append(novo)
            return lista_tratada

        def limpar_pontuacao_cnpj(lista):
            lista_tratada = []
            for item in lista:
                novo = item.replace('.','').replace('/','').replace('-','')
                lista_tratada.append(novo)
            return lista_tratada

        def concatenando_valores(lista_cc,lista_cnpj,lista_ped):
            lista_tratada = []
            for cc, cnpj, ped in zip(lista_cc,lista_cnpj,lista_ped):
                novo = cc+cnpj+ped
                lista_tratada.append(novo)
            return lista_tratada
        BD_TOTVS_MANUAL = self.BD_TOTVS_MANUAL
        BD_TOTVS_MANUAL = BD_TOTVS_MANUAL[BD_TOTVS_MANUAL['Categoria'].isin(['Pedido Mercado','Pedido','Pedido ACD','Pedido Modulo'])]
        BD_TOTVS_MANUAL['Codigo Unico'] = concatenando_valores(
            limpar_pontuacao_cc(BD_TOTVS_MANUAL['Centro de Custo'].astype(str).tolist()),
            limpar_pontuacao_cnpj(BD_TOTVS_MANUAL['CNPJ/CPF'].astype(str).tolist()),
            BD_TOTVS_MANUAL['N Pedido'].astype(str)
        )
        BD_TOTVS_MANUAL.to_excel(
        Dir().dir_makefolder(Dir().dir_sistema("BD_Base"),'Suprimentos')
        + Dir().get_separator() + "BD_NFs_Base" + '.xlsx')

    def RM_e_Pedido(self):
        BD_Req = self.BD_RM_FREQ
        BD_Pedido = self.BD_PED_FREQ
        DF_concat = BD_Req.merge(BD_Pedido,how='left',on='Pré-Pedido')
        DF_concat.to_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Frequencia"),'Suprimentos')
            + Dir().get_separator() + "BD_RM_Pedido" + '.xlsx')      

    def RM_Ped_NF(self):
        BD_RM_Ped = pd.read_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Frequencia"),'Suprimentos')
            + Dir().get_separator() + "BD_RM_Pedido" + '.xlsx')       
        BD_NF = self.BD_NF_Base
        DF_Concat = BD_NF.merge(BD_RM_Ped,how='left',on='Codigo Unico')
        DF_Concat.to_excel(
            Dir().dir_makefolder(Dir().dir_sistema("BD_Frequencia"),'Suprimentos')
            + Dir().get_separator() + "BD_RM_Ped_NF" + '.xlsx')  

#Relatorios_frequencia().frequencia_pedido()
#Relatorios_frequencia().resumo_rm()
Agrupando_Relatorios().DF_Justificativa()
