import pandas as pd
import os

def read_xlsx_files(path):
    result = []
    for filename in os.listdir(path):
        if filename.endswith('.xlsx'):
            df = pd.read_excel(os.path.join(path, filename), sheet_name=0, header=2)
            df["Arquivo"] = filename
            result.append(df)
    return result

def total_items(dfs, items):
    result = []
    for df in dfs:
        for item in items:
            filtered = df[df['Serviços'] == item]
            count = filtered['Serviços'].count()
            total = filtered['Valor'].sum()
            faturado = filtered['Valor Faturado'].sum()
            result.append({
                'Arquivo': df['Arquivo'].iloc[0],
                'Serviços': item,
                'Quantidade': count,
                'Vlr Total': total,
                'Vlr Faturado': faturado
            })
    return result

if __name__ == '__main__':
    path = '/workspaces/Planilhas/Relatórios'
    items = ['Suporte local a hardware e software - estação com garantia',
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
             'Administração e Manutenção de Redes Locais',
             'Disponibilização de Servidor Computacional',
             'Disponibilização de Servidor de Arquivos',
             'Hospedagem de Aplicação e Armazenamento de Dados - Sistemas',
             'Gestão e operação da Rede Digital de Telefonia Municipal (RDTM)',
             'Suporte técnico a evento sazonal',
             'Consultoria de Infraestrutura',
             'Outro (informar)']
    dfs = read_xlsx_files(path)
    result = total_items(dfs, items)
    df_result = pd.DataFrame(result)
    print(df_result)
