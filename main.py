import analises
import relatorios

class Create_Base_Sheets:

    def __init__(self):        
        self.lista_suprimentos_base = analises.Agrupando_Relatorios().lista_ut_suprimentos()

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

Create_Base_Sheets().pedido_base()
