Sistemas = {

    'webdriver_config':{
        'executable_path':r'\Documents\Python Scripts\Suprimentos_Manserv\Suprimentos_Manserv\chromedriver.exe'
    },
    'fluig':{
        'urls':{
            'url_login':'https://manserv.fluigidentity.com/cloudpass/?forward=%2FSPInitPost%2FreceiveSSORequest%2F2cbp2mgimnsg0v8d1434473628905%2Fo7fzggeaqhc1fut31471876606303&id=login_form',
            'url_modulo':'http://fluig.manserv.com.br/portal/p/1/consulta-saldo-contrato',
            'url_pedidos':'http://fluig.manserv.com.br/portal/p/1/consulta-pedido-compra',
            'url_hierarquia':'http://fluig.manserv.com.br/portal/p/1/consulta-hierarquia'
        },
        'html_elements':{
            'html_element_login':'emailAddress',
            'html_element_password':'password',
            'html_element_submit_bt':'login-label',
            'html_procced_bt':'proceed-button',
            'html_tipo_cb':'tipo',
            'html_coligada_cb':'coligadaCodigo',
            'html_codigo_ut':'utCodigo',
            'html_numero_ut':'numero',
            'html_pesquisar_bt':'btn-pesquisar',
            'html_table_pedido':'/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div/table',
            "hierarquia_info": {
                "label_coligada":"targetInfo1",
                "label_nomeclatura_ut":"targetInfo2",
                "label_nome_presidente": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[7]/p[2]",
                "label_email_presidente": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[7]/p[3]",
                "label_nome_diretorgeral": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[6]/p[2]",
                "label_email_diretorgeral": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[6]/p[3]",
                "label_nome_diretor": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[5]/p[2]",
                "label_email_diretor": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[5]/p[3]",
                "label_nome_gerente": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[4]/p[2]",
                "label_email_gerente": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[4]/p[3]",
                "label_nome_coordenador": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[3]/p[2]",
                "label_email_coordenador": "/html/body/div[3]/div[2]/div/div[1]/section/form/fieldset[2]/div[2]/div[1]/div/div[3]/p[3]",
            },
        },
        'hierarquia_colunas':{
            "colunas_modificar": ['nome_presidente','nome_diretorgeral','nome_diretor','nome_gerente','nome_coordenador']
        },
    },
    'mercado':{
        "urls": {
            "url_login": "https://www.me.com.br/do/Login.mvc/",
            "url_relatorio": "https://www.me.com.br/ME/ExtracaoRelatorio.aspx",
            "url_base_rm": "https://www.me.com.br/DO/Request/Home.mvc/Show/",
        },
        "html_elements": {
            "txtbox_path_usuario": "LoginName",
            "txtbox_path_senha": "RAWSenha",
            "button_path_submit": "SubmitAuth",
            "link_path_relatorio109": "ctl00_conteudo_wzdExtracao_frmRelatorio_dtlRelatorios_ctl09_myLinkButton",
            "link_path_relatorio108": "ctl00_conteudo_wzdExtracao_frmRelatorio_dtlRelatorios_ctl08_myLinkButton",
            "combox_path_centrocusto": "ctl00_conteudo_wzdExtracao_frmFiltros_CentroCusto",
            "button_path_processar_rel": "ctl00_conteudo_wzdExtracao_StepNavigationTemplateContainerID_ButtonBar2_btn_ctl00_conteudo_wzdExtracao_StepNavigationTemplateContainerID_ButtonBar2_btnFirst",
            "checkbox_path_sendemail": "ctl00_conteudo_wzdExtracao_frmParametrizacao_chkEnviaEmail",
            "label_path_statusprocess": "/html/body/main/form/div[4]/div/div[2]/table/tbody/tr/td[5]",
            "label_path_numprocess": "/html/body/main/form/div[4]/div/div[2]/table/tbody/tr/td[1]",
            "button_path_gerar_rel": "/html/body/main/form/div[4]/div/div[2]/table/tbody/tr/td[9]/img",
            "button_path_atual_rel": "ctl00_conteudo_btnBar2_btn_ctl00_conteudo_btnBar2_ImageButton1",
            "label_path_justificativa": "txt_Preferences_Attributes_1__valor",
            "label_path_prisma": "txt_Preferences_Attributes_4__valor",
        },
        "Report_info":"Relatorio_Relatorio de Acompanhamento - Data Aprovacao Req. _"
    }
}

database_info = {
    "user": "admin",
    "password": "Db.130996",
    "host": "database-1.cvhlssk69mjn.sa-east-1.rds.amazonaws.com",
    "port": "3306",
    "Sharepoint": {
        "url_site": "https://tiadminmanserv.sharepoint.com/sites/ManservCode/",
        "list_name": "Lista_UTs",
        "Usuario_dados": {
            "Usuario": "dhiego.braga@manserv.com.br",
            "Senha": "rvry196$",
        },
    },
}