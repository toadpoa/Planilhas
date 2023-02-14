import pandas as pd
import os

def consolidar_planilhas():
    # lista para armazenar as informações consolidadas
    informacoes = []

    # caminho para a pasta com as planilhas a serem consolidadas
    pasta = '/workspaces/Planilhas/Relatórios/'

    # itera por todos os arquivos na pasta
    for arquivo in os.listdir(pasta):
        # verifica se é um arquivo .xlsx
        if arquivo.endswith('.xlsx'):
            # lê a primeira tabela do arquivo
            df = pd.read_excel(os.path.join(pasta, arquivo), sheet_name=1)

            # seleciona as informações desejadas, começando na coluna 1 linha 4 e finalizando na coluna 7 linha 21
            df = df.iloc[2:21, 0:7]

            # adiciona informação sobre o nome do arquivo original às informações consolidadas
            df['Nome do Arquivo Original'] = arquivo

            # adiciona informações consolidadas a lista informacoes
            informacoes.append(df)

    # cria tabela consolidada a partir da lista de informacoes
    tabela_consolidada = pd.concat(informacoes, ignore_index=True)


    # retorna tabela consolidada
    return tabela_consolidada
    

# chama a função consolidar_planilhas e armazena o resultado em uma variável
tabela = consolidar_planilhas()
df_final = pd.DataFrame(tabela, columns=['Nome do Arquivo Original', 'Serviços', 'Quantidade', 'vlr Faturado'])

# exibe o resultado
print(tabela)
print(df_final)

# Salva o dataframe final como um arquivo .xlsx
tabela.to_excel("Relatório Final.xlsx", index=False)
