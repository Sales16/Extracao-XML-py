import os

import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

diretorio = filedialog.askdirectory()

for arquivo in os.listdir(diretorio):
    if arquivo.endswith(".csv"):
        if "COMPRA" not in arquivo:
            caminho_arquivo = os.path.join(diretorio, arquivo)
            try:
                os.startfile(caminho_arquivo)
                print(f"Arquivo {arquivo} Aberto")
            except IsADirectoryError:
                pass
                print(f"Erro ao abrir o arquivo {arquivo}")
    else:
        print(f"Arquivo {arquivo} é um arquivo de compra")
print('Todos arquivos abertos')