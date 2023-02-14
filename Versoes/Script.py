import os
import pandas as pd

# Local dos arquivos
data_dir = "/workspaces/Planilhas/Relatórios"

# Lista de arquivos
files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx')]

# Resultado final
result = []

# Loop por arquivos
for file in files:
    # Lê a primeira tabela do arquivo
    df = pd.read_excel(os.path.join(data_dir, file), sheet_name=0, header=None, skiprows=2)
    
    # Armazena as informações desejadas
    result.append([file, df.iloc[0, 1], df.iloc[0, 2], df.iloc[0, 3]])

# Cria o dataframe final
df_final = pd.DataFrame(result, columns=['Nome do Arquivo original', 'Serviço', 'Quantidade', 'vlr Faturado'])

# Salva o dataframe final como um arquivo .xlsx
df_final.to_excel("Relatório Final.xlsx", index=False)

# Exibe a tabela final
print(df_final)