import time

tempo_inicial = time.time()

import os
import tkinter as tk
from tkinter import filedialog
from lxml import etree
from openpyxl import Workbook
from tkinter import simpledialog

# Cria a janela do tkinter e a esconde
root = tk.Tk()
root.withdraw()

# Exibe uma caixa de diálogo para o usuário selecionar a pasta
folder_path = filedialog.askdirectory()
# Define o namespace da NFe
ns = {
    "nfe": "http://www.portalfiscal.inf.br/nfe"
}

# Define o nome das colunas da planilha
columns = ["COMBUSTIVEL", "DATA", "NOTA", "ITEM", "V", "CODIGO", "VALOR", "QUANTIDADE"]



# Itera sobre todas as pastas e subpastas no diretório especificado
cnpj_posto = simpledialog.askstring("CNPJ", "ESPECIFIQUE O CNPJ DO POSTO")
for root, dirs, files in os.walk(folder_path):
    # Cria uma nova planilha do Excel e adiciona os cabeçalhos das colunas
    wb = Workbook()
    ws = wb.active
    ws.title = "Notas Fiscais"
    ws.append(columns)
    x = 1
    data_ano = ""
    y = 1
    z = 1
    # Itera sobre todos os arquivos XML na pasta atual e extrai as informações relevantes
    for filename in files:
        if filename.endswith(".xml"):
            xml_path = os.path.join(root, filename)
            
            with open(xml_path, "rb") as f:
                xml = f.read()

            try:
                root_element = etree.fromstring(xml)
            except etree.XMLSyntaxError:
                print(f"Erro ao ler arquivo {xml_path}")
                continue

            root_element = etree.fromstring(xml)

            for item in root_element.xpath("//nfe:det", namespaces=ns):
                cnpj_emit = item.xpath("ancestor::nfe:infNFe/nfe:emit/nfe:CNPJ/text()", namespaces=ns)
                if len(cnpj_emit) > 0:
                    cnpj_emit = cnpj_emit[0][:14]
                else:
                    cnpj_emit = "---"
                if cnpj_emit == cnpj_posto:
                    print(x, 'NOTAS PROCESSADAS DO MES:', folder_name)
                    x = x + 1
                    produto = item.xpath("nfe:prod/nfe:xProd/text()", namespaces=ns)[0]
                    data_emissao = item.xpath("ancestor::nfe:infNFe/nfe:ide/nfe:dhEmi/text()", namespaces=ns)[0][:10]
                    data_ano = item.xpath("ancestor::nfe:infNFe/nfe:ide/nfe:dhEmi/text()", namespaces=ns)[0][:4]
                    id_nfe = item.xpath("ancestor::nfe:NFe/nfe:infNFe/@Id", namespaces=ns)[0]
                    n_item = float(item.xpath("@nItem", namespaces=ns)[0])
                    codigo_produto = item.xpath("nfe:prod/nfe:cProd/text()", namespaces=ns)[0]
                    valor_unitario = float(item.xpath("nfe:prod/nfe:vUnCom/text()", namespaces=ns)[0])
                    quantidade = float(item.xpath("nfe:prod/nfe:qCom/text()", namespaces=ns)[0])

                    # Adiciona as informações extraídas na planilha do Excel
                    ws.append([produto, data_emissao, id_nfe, n_item, y, codigo_produto, valor_unitario, quantidade])
                else:
                    print(z, 'NOTAS NÃO PROCESSADAS DO MES:', folder_name)
                    z = z + 1
    # Define o nome da pasta como o nome da planilha
    folder_name = os.path.basename(root)
    worksheet_name = f"{data_ano}{folder_name}_{cnpj_posto}_0101.xlsx"
    wb.save(os.path.join(folder_path, worksheet_name))

# Define o tamanho mínimo do arquivo em bytes
min_file_size = 4 * 1024

# Itera sobre todos os arquivos Excel na pasta atual e exclui os arquivos com menos de 6KB
for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx"):
        file_path = os.path.join(folder_path, filename)
        file_size = os.path.getsize(file_path)
        if file_size < min_file_size:
            os.remove(file_path)
            print(f"Arquivo {filename} excluído por ser menor que 6KB")

            
   
    # Imprime o número total de notas fiscais processadas na pasta atual
    print(f"Todas as notas fiscais foram processadas na pasta {folder_name}")

tempo_final = time.time()
tempo_total = tempo_final - tempo_inicial
tempo_total_min = tempo_total / 60
print(f"Tempo total de execução: {tempo_total} segundos")
print(f"Tempo total de execução: {tempo_total_min} minutos")