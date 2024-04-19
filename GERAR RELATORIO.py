import os
import csv
import tkinter as tk
from tkinter import filedialog
import pandas as pd

root = tk.Tk()
root.withdraw()
diretorio = filedialog.askdirectory()
coluna = 1

dados = []

# Iterar sobre todos os arquivos CSV nas pastas
for pasta, subpastas, arquivos in os.walk(diretorio):
    for arquivo in arquivos:
        if arquivo.endswith(".csv"):
            caminho_arquivo = os.path.join(pasta, arquivo)

            # Abrir o arquivo CSV e capturar o último valor na coluna desejada
            with open(caminho_arquivo, "rb") as f:
                conteudo = f.read()
                conteudo_sem_nulos = conteudo.replace(b'\x00', b'')
                leitor_csv = csv.reader(conteudo_sem_nulos.decode('iso-8859-1').splitlines())
                ultima_linha = None
                for linha in leitor_csv:
                    if linha:
                        ultima_linha = linha
                if ultima_linha is not None and len(ultima_linha) >= 2:
                    dados.append((pasta,arquivo, ultima_linha[0]+","+ultima_linha[1]))
                    print(arquivo, "capturada com sucesso")
                else:
                    dados.append((pasta, arquivo, "NAO POSSUI INFORMAÇÃO"))
                    print(arquivo, "capturada com sucesso")

df = pd.DataFrame(dados, columns=["pasta", "Arquivo", "Últimas informações"])

# Salvar o DataFrame em uma planilha do Excel
caminho_excel = r"E:\PROJETO\relatorio.xlsx"  # Substitua pelo caminho correto para o arquivo do Excel
df.to_excel(caminho_excel, index=False)

print("Planilha do Excel salva com sucesso!")