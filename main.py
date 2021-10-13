import analises
import relatorios

class Create_Base_Sheets:

    def __init__(self):        
        self.lista_suprimentos_base = analises.Agrupando_Relatorios().lista_ut_suprimentos()
        self.ut_mercado = analises.Configs().Data_UT()[0]
        self.lista_data_inicio =analises.Configs().Data_UT()[1]
        self.lista_data_fim =analises.Configs().Data_UT()[2]

    def hierarquia_base(self):
        relatorios.Relatorios_Fluig(self.lista_suprimentos_base).save_hierarquia()
        analises.Agrupando_Relatorios().DF_Hierarquia()

    def modulo_base(self):
        lista_ut_coligada = analises.Agrupando_Relatorios().lista_ut_coligada()
        relatorios.Relatorios_Fluig("").save_table_modulo(lista_ut_coligada[1],lista_ut_coligada[0])
        analises.Agrupando_Relatorios().DF_Modulos()

    def pedido_base(self):
        lista_ut_coligada = analises.Agrupando_Relatorios().lista_ut_coligada()
        relatorios.Relatorios_Fluig("").save_table_pedido(lista_ut_coligada[1],lista_ut_coligada[0])
        analises.Agrupando_Relatorios().DF_Pedidos()

    def bd_rms(self):
        relatorios.Relatorio_Mercado(self.ut_mercado,self.lista_data_inicio,self.lista_data_fim).save_report()
        analises.Agrupando_Relatorios().DF_RMs_Mercado()

    def bd_nfs(self):
        analises.Agrupando_Relatorios().DF_NFs_TOTVS()
        analises.Limpando_BDs().comparando_bds()
        analises.Limpando_BDs().save()

    def all_bd(self):
        #self.hierarquia_base()
        self.bd_nfs()
        #self.bd_rms()
        
class Analisando_Relatorios:

    def __init__(self) -> None:
        pass

    def dados_analisados_mercado(self):
        analises.Relatorios_frequencia().resumo_status_rm()

    def dados_analisados_fluig(self):
        analises.Relatorios_frequencia().linkar_pedido()

    def dados_analisados_totvs(self):
        analises.Relatorios_frequencia().frequencia_notafiscal()

#Create_Base_Sheets().all_bd()
#Analisando_Relatorios().dados_analisados_mercado()
#Analisando_Relatorios().dados_analisados_fluig()
#Analisando_Relatorios().dados_analisados_totvs()