from dados import database_info as database_configs
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

class Database:
    class Sharepoint:
        def __init__(self):
            self.cliente_context = ClientContext(
                database_configs["Sharepoint"]["url_site"]
            ).with_credentials(
                UserCredential(
                    database_configs["Sharepoint"]["Usuario_dados"]["Usuario"],
                    database_configs["Sharepoint"]["Usuario_dados"]["Senha"],
                )
            )

        def Itens_lista_UT(self):
            sp_list = database_configs["Sharepoint"]["list_name"]
            sp_lists = self.cliente_context.web.lists
            s_list = sp_lists.get_by_title(sp_list)
            l_items = s_list.get_items()
            self.cliente_context.load(l_items)
            self.cliente_context.execute_query()
            return l_items

        def dicionario_UTs(self):
            itens = self.Itens_lista_UT()
            dicionario = {
                "UT": [],
                "Coligada": [],
                "Nomeclatura": [],
                "Em Atividade": [],
                "Relatorio_Telemetria": [],
                "Relatorio_Pool": [],
                "Relatorio_Suprimentos": [],
            }
            for item in itens:
                dicionario["UT"].append(item.properties["UT"])
                dicionario["Coligada"].append(item.properties["Nomeclatura_UT"])
                dicionario["Nomeclatura"].append(item.properties["Nome_UT"])
                dicionario["Em Atividade"].append(item.properties["UTAtiva_x003f_"])
                dicionario["Relatorio_Telemetria"].append(
                    item.properties["Utiliza_x00e7__x00e3_oTelemetria"]
                )
                dicionario["Relatorio_Pool"].append(
                    item.properties["Utiliza_x00e7__x00e3_oPool"]
                )
                dicionario["Relatorio_Suprimentos"].append(
                    item.properties["Utiliza_x00e7__x00e3_oSuprimento"]
                )
            return dicionario