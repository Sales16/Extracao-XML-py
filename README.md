# Extracao-XML-py
Scripts feitos em python para extração de dados contidos em XMLs de notas ficais de postos de gasolina.

* 1 - EXTRAIR VENDAS - CNPJ
Extrai os dados de vendas dos XMLs de venda dos postos de combustiveis especificado pelo CNPJ.

Inicia o cronometro.
Especifica o diretorio principal.
Especifica o CNPJ do posto.
Itera em todas as pastas do diretorio especificado.
Itera em todos os arquivos XML da pasta atual e extrai as informações relevantes contidas no XML.
Salva os dados extraidos em uma planilha XLSX.
Define o nome da planilha.
Apaga todas planilhas que forem menor que 4KB
Finaliza o cronometro
Exibe informações como (Notas processadas, Notas não processadas, tempo total de execução em M\S).

** NESCESSARIO ORGANIZAÇÃO DOS ARQUIVOS EM PASTAS COMO NO EXEMPLO DE DIRETORIO ABAIXO **

    Diretorio-principal         ---Diretorio Raiz   ** DIRETORIO A SER ESPECIFICADO NO SCRIPT **
        2018                        ---Ano
            01                          ---Mês
                .xml                        ---Arquivos .xml de vendas que serão extraido as informações
                .xml
                .xml
            02
                .xml
                .xml
                .xml
            03
            XX
        2019
        XXXX
        COMPRAS                     ---Pastas dos XMLs de compras
            COMPRA 2018
                .xml


* 2 - EXTRAIR COMPRAS - CNPJ
Extrai os dados de compras dos XMLs de compras dos postos de combustiveis especificado pelo CNPJ.

Especifica o diretorio das compras.
Especifica o CNPJ do posto.
Itera em todas as pastas do diretorio especificado.
Itera em todos os arquivos XML da pasta atual.
Apaga todos os XMLs cancelados.
Extrai as informações relevantes contidas no XML.
Se não possuir a informações, ela é considerada 0.
Salva os dados extraidos em uma planilha XLSX.
Define o nome da planilha.
Exibe informações como (Notas processadas, Notas não processadas).

** NESCESSARIO ORGANIZAÇÃO DOS ARQUIVOS EM PASTAS COMO NO EXEMPLO DE DIRETORIO ABAIXO **

    Diretorio-principal         ---Diretorio Raiz
        2018                        ---Ano
            01                          ---Mês
                .xml                        ---XML de venda (NÃO SERÁ USADO)
        2019
        COMPRAS                     ---Pasta que contem todas compras dos anos  ** DIRETORIO A SER ESPECIFICADO NO SCRIPT **
            COMPRA 2018                 ---Todas compras do ano(os XMLs que será extraido as informações deverá estar nestas pastas)
                .xml                        ---XML de COMPRA
                .xml
                .xml
            COMPRA 2019
            COMPRA XXXX

3 - EXTRAIR COMPRAS - COTEPE
Extrai os dados de compras dos XMLs de compras dos postos de combustiveis especificado pelo CNPJ. A extração será feita para a utilização da tabela COTEPE.

Especifica o diretorio principal.
Especifica o CNPJ do posto.
Itera em todas as pastas do diretorio especificado.
Itera em todos os arquivos XML da pasta atual.
Extrai as informações relevantes contidas no XML.
Se não possuir a informações, ela é considerada 0.
Informa os dados nescessarios para a utilização da COTEPE.
Salva os dados extraidos em uma planilha XLSX.
Define o nome da planilha.
Exibe informações como (Notas processadas, Notas não processadas).

** NESCESSARIO ORGANIZAÇÃO DOS ARQUIVOS EM PASTAS COMO NO EXEMPLO DE DIRETORIO ABAIXO **

    Diretorio-principal         ---Diretorio Raiz
        2018                        ---Ano
            01                          ---Mês
                .xml                        ---XML de venda (NÃO SERÁ USADO)
        2019
        COMPRAS                     ---Pasta que contem todas compras dos anos  ** DIRETORIO A SER ESPECIFICADO NO SCRIPT **
            COMPRA 2018                 ---Todas compras do ano(os XMLs que será extraido as informações deverá estar nestas pastas)
                .xml                        ---XML de COMPRA
                .xml
                .xml
            COMPRA 2019
            COMPRA XXXX

4 - GERAR RELATORIO
Gera um relatorio com os dados dos CSVs de cada posto, organizando essas dados e os salva em uma planilha fixa.

Especifica o diretorio dos postos.
Itera em todos os CSVs.
Extrai apenas os dados nescessarios para o relatorio.
Salva os dados em uma planilha fixa.
