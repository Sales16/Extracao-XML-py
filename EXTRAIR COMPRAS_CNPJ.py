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

keywords = ["CANCELADA", "CANCELADO", "CANCEL"]

# Define o nome das colunas da planilha
columns = ["COMBUSTIVEL", "DATA", "NOTA", "ITEM", "CODIGO", "PREÇO", "%", "QUANTIDADE", "Brest", "ICMSrest", "Bdest", "ICMSdest", "Bc", "ICMS", "BCST", "texto"]

# Itera sobre todas as pastas e subpastas no diretório especificado
cnpj_posto = simpledialog.askstring("CNPJ", "ESPECIFIQUE O CNPJ DO POSTO")
for root, dirs, files in os.walk(folder_path):
    # Cria uma nova planilha do Excel e adiciona os cabeçalhos das colunas
    wb = Workbook()
    ws = wb.active
    ws.title = "Notas Fiscais"
    ws.append(columns)
    x = 1
    y = 1
    vbcstret = 0
    vicmsstret = 0
    vbcstdest = 0
    vicmsstdest = 0
    vbc = 0
    vicms = 0
    vbcst = 0
    cnpj_dest = 0
    A = ""
    B = ""
    # Itera sobre todos os arquivos XML na pasta atual e extrai as informações relevantes
    for filename in files:
        if filename.endswith(".xml"):
            # Apaga Cancelados
            if any(keyword in filename for keyword in keywords):
                file_path = os.path.join(root, filename)
                print(f"Arquivo Apagado: {file_path}")
                os.remove(file_path)
            else:
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
                        produto = item.xpath("nfe:prod/nfe:xProd/text()", namespaces=ns)[0]
                        data_emissao = item.xpath("ancestor::nfe:infNFe/nfe:ide/nfe:dhEmi/text()", namespaces=ns)[0][:10]
                        id_nfe = item.xpath("ancestor::nfe:NFe/nfe:infNFe/@Id", namespaces=ns)[0]
                        n_item = float(item.xpath("@nItem", namespaces=ns)[0])
                        codigo_produto = item.xpath("nfe:prod/nfe:cProd/text()", namespaces=ns)[0]
                        quantidade = float(item.xpath("nfe:prod/nfe:qCom/text()", namespaces=ns)[0])
                        icms = item.xpath("nfe:imposto/nfe:ICMS", namespaces=ns)
                
                        if icms:
                            icmsST = icms[0].xpath("nfe:ICMSST", namespaces=ns)
                            icms60 = icms[0].xpath("nfe:ICMS60", namespaces=ns)
                            if icmsST:
                                vbcstret = item.xpath(".//nfe:ICMSST/nfe:vBCSTRet/text()", namespaces=ns)
                                if len(vbcstret) > 0:
                                    vbcstret = float(vbcstret[0])
                                else:
                                    vbcstret = 0
                                vicmsstret = item.xpath(".//nfe:ICMSST/nfe:vICMSSTRet/text()", namespaces=ns)
                                if len(vicmsstret) > 0:
                                    vicmsstret = float(vicmsstret[0])
                                else:
                                    vicmsstret = 0
                                vbcstdest = item.xpath(".//nfe:ICMSST/nfe:vBCSTDest/text()", namespaces=ns)
                                if len(vbcstdest) > 0:
                                    vbcstdest = float(vbcstdest[0])
                                else:
                                    vbcstdest = 0
                                vicmsstdest = item.xpath(".//nfe:ICMSST/nfe:vICMSSTDest/text()", namespaces=ns)
                                if len(vicmsstdest) > 0:
                                    vicmsstdest = float(vicmsstdest[0])
                                else:
                                    vicmsstdest = 0
                            if icms60:
                                vbcstret = item.xpath(".//nfe:ICMS60/nfe:vBCSTRet/text()", namespaces=ns)
                                if len(vbcstret) > 0:
                                    vbcstret = float(vbcstret[0])
                                else:
                                    vbcstret = 0
                                vicmsstret = item.xpath(".//nfe:ICMS60/nfe:vICMSSTRet/text()", namespaces=ns)
                                if len(vicmsstret) > 0:
                                    vicmsstret = float(vicmsstret[0])
                                else:
                                    vicmsstret = 0
                                vbcstdest = item.xpath(".//nfe:ICMS60/nfe:vBCSTDest/text()", namespaces=ns)
                                if len(vbcstdest) > 0:
                                    vbcstdest = float(vbcstdest[0])
                                else:
                                    vbcstdest = 0
                                vicmsstdest = item.xpath(".//nfe:ICMS60/nfe:vICMSSTDest/text()", namespaces=ns)
                                if len(vicmsstdest) > 0:
                                    vicmsstdest = float(vicmsstdest[0])
                                else:
                                    vicmsstdest = 0
                        vbc = item.xpath("ancestor::nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vBC/text()", namespaces=ns)
                        if len(vbc) > 0:
                            vbc = float(vbc[0])
                        else:
                            vbc = 0
                        vicms = item.xpath("ancestor::nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vICMS/text()", namespaces=ns)
                        if len(vicms) > 0:
                            vicms = float(vicms[0])
                        else:
                            vicms = 0
                        vbcst = item.xpath("ancestor::nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vBCST/text()", namespaces=ns)
                        if len(vbcst) > 0:
                            vbcst = float(vbcst[0])
                        else:
                            vbcst = 0
                        inf_adic = item.xpath("ancestor::nfe:infNFe/nfe:infAdic/nfe:infCpl/text()", namespaces=ns)
                        if inf_adic:
                            inf_adic = inf_adic[0]
                        else:
                            inf_adic = 0

                        # Adiciona as informações extraídas na planilha do Excel
                        ws.append([produto, data_emissao, id_nfe, n_item, codigo_produto, A, B, quantidade, vbcstret, vicmsstret, vbcstdest, vicmsstdest, vbc, vicms, vbcst, inf_adic])
                    else:
                        print(y, 'NOTA NÃO PROCESSADA DO MES:', folder_name)
                        y=y+1
    print(f'Pasta {os.path.basename(root)} processada')
    print(f'Notas processadas {x-1}\nNotas não processadas {y-1}')
    
    # Define o nome da pasta como o nome da planilha
    folder_name = os.path.basename(root)
    worksheet_name = f"COMPRAS-{cnpj_posto}-{folder_name}.xlsx"
    wb.save(os.path.join(folder_path, worksheet_name))

    # Imprime o número total de notas fiscais processadas na pasta atual
    print(f"Todas as notas fiscais foram processadas na pasta {folder_name}")
        