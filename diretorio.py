import os
import time
import zipfile
from sys import platform

class Dir_analise:
    def __init__(self):
        self.operation_system = platform

    def get_separator(self):
        if self.operation_system == "linux" or self.operation_system == "linux2":
            separator = r"/"
        elif self.operation_system == "win32":
            separator = r"\\"
        else:
            separator = r":"
        return separator

    def check_dir(self, path):
        check = os.path.exists(path)
        if not check:
            os.makedirs(path)
        else:
            pass

    def check_file(self, path):
        check = os.path.exists(path)
        return check

    def dir_base(self):
        diretorio_base = os.path.expanduser('~')
        return diretorio_base

    def dir_sistema(self,sistema):
        diretorio_sis = self.dir_base() + self.get_separator() + 'temp' + self.get_separator() + sistema
        self.check_dir(diretorio_sis)
        return diretorio_sis

    def dir_makefolder(self,caminho_base,nova_pasta):
        novo_caminho = caminho_base + self.get_separator() + nova_pasta
        self.check_dir(novo_caminho)
        return novo_caminho

    def extract_file(self, path_zip, path_xlsx):
        with zipfile.ZipFile(path_zip) as zip_ref:
            zip_ref.extractall(path_xlsx)

    def list_files(self, diretorio):
        lista_caminhos_arquivos = []
        arquivos = os.listdir(diretorio)
        for item in arquivos:
            lista_caminhos_arquivos.append(item)
        return lista_caminhos_arquivos        

    def remove_file(self,path):
        os.remove(path)
    
    def rename_file(self,old_path,new_path):
        os.rename(old_path,new_path)