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
columns = ["COMBUSTIVEL", "DATA", "NOTA", "ITEM", "CODIGO", "PREÇO", "%", "QUANTIDADE", "ANO", "MES", "DIA", "PERIODO", "UF"]

# Pergunta CNPJ do posto
cnpj_posto = simpledialog.askstring("CNPJ", "ESPECIFIQUE O CNPJ DO POSTO")

# Itera sobre todas as pastas e subpastas no diretório especificado
for root, dirs, files in os.walk(folder_path):
    # Cria uma nova planilha do Excel e adiciona os cabeçalhos das colunas
    wb = Workbook()
    ws = wb.active
    ws.title = "Notas Fiscais"
    ws.append(columns)
    x = 1
    y = 1
    cnpj_dest = "---"
    A = ""
    B = ""
    C = 1
    D = 2
    # Itera sobre todos os arquivos XML na pasta atual e extrai as informações relevantes
    for filename in files:
        if filename.endswith(".xml"):
            xml_path = os.path.join(root, filename)
            folder_name = os.path.basename(root)
            
            with open(xml_path, "rb") as f:
                xml = f.read()
            try:
                root_element = etree.fromstring(xml)
            except etree.XMLSyntaxError:
                print(f"Erro ao ler arquivo {xml_path}")
                continue
            root_element = etree.fromstring(xml)

            for item in root_element.xpath("//nfe:det", namespaces=ns):
                cnpj_dest = item.xpath("ancestor::nfe:infNFe/nfe:dest/nfe:CNPJ/text()", namespaces=ns)
                if len(cnpj_dest) > 0:
                    cnpj_dest = cnpj_dest[0][:14]
                else:
                    cnpj_dest = "---"
                cnpj_emit = item.xpath("ancestor::nfe:infNFe/nfe:emit/nfe:CNPJ/text()", namespaces=ns)
                if len(cnpj_emit) > 0:
                    cnpj_emit = cnpj_emit[0][:14]
                else:
                    cnpj_emit = "---"
                    
                if cnpj_dest == cnpj_posto and cnpj_emit != cnpj_posto:
                    print(x, 'NOTAS PROCESSADAS DO MES:', folder_name)
                    x = x+1
                    Uf = item.xpath("ancestor::nfe:infNFe/nfe:emit/nfe:enderEmit/nfe:UF/text()", namespaces=ns)[0]
                    produto = item.xpath("nfe:prod/nfe:xProd/text()", namespaces=ns)[0]
                    data_emissao = item.xpath("ancestor::nfe:infNFe/nfe:ide/nfe:dhEmi/text()", namespaces=ns)[0][:10]
                    ano = int(data_emissao[:4])
                    mes = int(data_emissao[5:7])
                    dia = int(data_emissao[8:10])
                    id_nfe = item.xpath("ancestor::nfe:NFe/nfe:infNFe/@Id", namespaces=ns)[0]
                    n_item = float(item.xpath("@nItem", namespaces=ns)[0])
                    codigo_produto = item.xpath("nfe:prod/nfe:cProd/text()", namespaces=ns)[0]
                    quantidade = float(item.xpath("nfe:prod/nfe:qCom/text()", namespaces=ns)[0])

                    if 1 <= dia <= 15:
                        #str(ano, mes, dia)
                        ws.append([produto, data_emissao, id_nfe, n_item, codigo_produto, A, B, quantidade, ano, mes, dia, C, Uf])
                    else:
                        #str(ano, mes, dia)
                        ws.append([produto, data_emissao, id_nfe, n_item, codigo_produto, A, B, quantidade, ano, mes, dia, D, Uf])
                else:
                    print(y, 'NOTA NÃO PROCESSADA DO MES:', folder_name)
                    y=y+1
    print(f'Pasta {os.path.basename(root)} processada')
    print(f'Notas processadas {x-1}\nNotas não processadas {y-1}')
    
    # Define o nome da pasta como o nome da planilha
    folder_name = os.path.basename(root)
    worksheet_name = f"COMPRAS-{cnpj_posto} - {folder_name}.xlsx"
    wb.save(os.path.join(folder_path, worksheet_name))

    # Imprime o número total de notas fiscais processadas na pasta atual
    print(f"Todas as notas fiscais foram processadas na pasta {folder_name}")
        