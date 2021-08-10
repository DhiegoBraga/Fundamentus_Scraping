import os

dir_dict = {
    'dir_base_documentos':os.path.expanduser('~') + r'\Documents',
    'dir_base_codigo':os.path.abspath(os.getcwd()),
    'dir_pasta_codigo':r'\Fundamentus_Scraping'
}

chrome_driver_path = dir_dict['dir_base_codigo'] + dir_dict['dir_pasta_codigo']
print(chrome_driver_path,dir_dict['dir_base_documentos'])