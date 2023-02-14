import os
import pandas as pd

# Caminho da pasta com os arquivos
path = '/workspaces/Planilhas/Relatórios/'

# Inicialização do dataframe consolidado
df_consolidated = pd.DataFrame(columns=['Nome do Arquivo Original', 'Serviços', 'Quantidade', 'vlr faturado'])

# Loop para ler cada arquivo na pasta
for filename in os.listdir(path):
    if filename.endswith('.xlsx'):
        df = pd.read_excel(os.path.join(path, filename), skiprows=2, nrows=18)
       
        # Seleciona as colunas de interesse
        df = df.iloc[:, [0, 1, 6]]
        
        # Renomeia as colunas para padronizar
        df.columns = ['Serviços', 'Quantidade', 'vlr faturado']
        
        df['Nome do Arquivo Original'] = filename
        df_consolidated = pd.concat([df_consolidated, df], ignore_index=True)

# Reordenação das colunas na tabela consolidada
df_consolidated = df_consolidated[['Nome do Arquivo Original', 'Serviços', 'Quantidade', 'vlr faturado']]

# Salvando a tabela consolidada em um arquivo .xlsx
df_consolidated.to_excel('relatorio_consolidado.xlsx', index=False)
print(df_consolidated)
