import os
import pandas as pd
import re

secretarias = ['GP', 'PGM', 'SMAMUS', 'SMAP', 'SMC', 'SMDET', 'SMDS', 'SMED', 'SMGOV', 'SMOI', 'SMP', 'SMS - Geral', 'SMS - AB', 'SMS - CGVS', 'SMSEG', 'SMS - HMIPV', 'SMS - HPS', 'SMS - PA', 'SMS - SAMU', 'SMS - SM', 'SMSURB', 'SMTC', 'SMPAE', 'DMLU', 'EPTC', 'FASC', 'PREVIMPA', 'CARRIS', 'DEMHAB', 'DMAE', 'SMF', 'IMESF', 'SMS-CR(Estado)', 'SMHARF', 'SMMU', 'SMELJ']
mes_competencia = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
ano_competencia = ['2020', '2021', '2022', '2023', '2024', '2025']
servicos = ['Administração e Manutenção de Redes Locais',
            'Disponibilização de Servidor Computacional',
            'Disponibilização de Servidor de Arquivos',
            'Hospedagem de Aplicação e Armazenamento de Dados - Sistemas',
            'Gestão e operação da Rede Digital de Telefonia Municipal (RDTM)', 
            'Suporte técnico a evento sazonal', 
            'Consultoria de Infraestrutura', 
            'Outro (informar):', 
            'Suporte local a hardware e software - estação com garantia',
            'Suporte local a hardware e software - estação sem garantia',
            'Suporte a impressoras',
            'Administração de Correio Eletrônico (e-mail)',
            'Administração e Manutenção de Câmeras de Videomonitoramento - indoor',
            'Administração e Manutenção de Câmeras de Videomonitoramento - Outdoor',
            'Administração e Manutenção de Rádio Wi-Fi - indoor',
            'Administração e Manutenção de Rádio Wi-Fi - outdoor',
            'Suporte e manutenção de Terminal de Radiocomunicação Digital',
            'Administração e Manutenção da Rede de Radiocomunicação Digital',
            'Gestão da Rede Infovia',
            'Análise e Desenvolvimento de sistemas de informação']

# Pasta com arquivos .xlsx
path = '/workspaces/Planilhas/Relatórios'

# Inicializa lista para armazenar resultados
resultados = []

# Loop para ler todos os arquivos .xlsx na pasta
for filename in os.listdir(path):
    if filename.endswith(".xlsx"):
        # Leitura do arquivo
        df = pd.read_excel(os.path.join(path, filename), sheet_name=1, skiprows=0)

        for index, row in df.iterrows():
            if row[0] in servicos:
                secretaria_encontrada = 'não identificado'
                for secretaria_possivel in secretarias:
                    if secretaria_possivel in str(df.iloc[:]):
                        secretaria_encontrada = secretaria_possivel

        for index, row in df.iterrows():
            if row[0] in servicos:
                mes_competencia_encontrada = 'não identificado'
                for mes_competencia_possivel in mes_competencia:
                    if mes_competencia_possivel in str(df.iloc[0:0]):
                        mes_competencia_encontrada = mes_competencia_possivel

        for index, row in df.iterrows():
            if row[0] in servicos:
                ano_competencia_encontrada = 'não identificado'
                for ano_competencia_possivel in ano_competencia:
                    if ano_competencia_possivel in str(df.iloc[0:0]):
                        ano_competencia_encontrada = ano_competencia_possivel

        for index, row in df.iterrows():
            if row[0] in servicos:
                sei = 'não identificado'
                for texto in df.iloc[:]:
                    # Expressão regular para buscar a string no formato 00.00.000000000-0
                    match = re.search(r"\d{2}\.\d{2}\.\d{9}-\d", str(texto))
                    if match:
                        sei = match.group()
                        break     
                
          # Loop para buscar objeto na coluna 0
        for index, row in df.iterrows():
            if row[0] in servicos:
                resultados.append({'Nome do Arquivo Original': filename,
                                   'Secretaria': secretaria_encontrada,
                                   'SEI': sei,
                                   'Mês': mes_competencia_encontrada,
                                   'Ano': ano_competencia_encontrada,
                                   'Serviços': row[0],
                                   'Vlr Total': row[3],
                                   'Vlr Faturado': row[6]})

# Cria dataframe com resultados
df_resultados = pd.DataFrame(resultados)

# Exibe resultados
print(df_resultados)
df_resultados.to_excel('consolidado.xlsx', index=False)
