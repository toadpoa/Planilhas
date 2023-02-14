import os
import pandas as pd

# Cria Lista Secretaria
secretaria = ['GP', 'PGM', 'SMAMUS']

# Cria a lista Serviços
servicos = ['Administração e Manutenção de Redes Locais', 'Disponibilização de Servidor Computacional', 'Disponibilização de Servidor de Arquivos']

# Pasta com arquivos .xlsx
pasta = '/workspaces/Planilhas/Relatórios'

# Lista para armazenar os resultados
resultados = []

# Percorre todos os arquivos na pasta
for arquivo in os.listdir(pasta):
    if arquivo.endswith('.xlsx'):
        nome_arquivo = arquivo
        arquivo = os.path.join(pasta, arquivo)
        df = pd.read_excel(arquivo, sheet_name=1, skiprows=2)
        
        for index, row in df.iterrows():
            if row[0] in servicos:
                secretaria_encontrada = 'não identificado'
                for secretaria_possivel in secretaria:
                    if secretaria_possivel in str(df.iloc[2,:]):
                        secretaria_encontrada = secretaria_possivel
                resultados.append([nome_arquivo, secretaria_encontrada, row[0], index+3])

# Cria DataFrame com resultados
df_resultados = pd.DataFrame(resultados, columns=['Nome do Arquivo Original', 'Secretaria', 'Serviços', 'Linha'])

# Mostra resultados
print(df_resultados)
