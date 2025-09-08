import pandas as pd
import os
import glob
from datetime import datetime

# --- Parte 1: Inteligência para Encontrar o Arquivo ---

def encontrar_relatorio_recente():
    """Encontra o arquivo 'Relatório*.xlsx' mais recente na pasta Downloads."""
    pasta_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    padrao_busca = os.path.join(pasta_downloads, 'Relatório*.xlsx')
    arquivos_encontrados = glob.glob(padrao_busca)
    
    if not arquivos_encontrados:
        raise FileNotFoundError(f"Nenhum arquivo 'Relatório*.xlsx' foi encontrado na pasta {pasta_downloads}")
    
    arquivo_mais_recente = max(arquivos_encontrados, key=os.path.getmtime)
    return arquivo_mais_recente

# --- Início do processamento principal ---

try:
    # ETAPA 1: Ler e fazer o tratamento inicial do arquivo
    arquivo_entrada = encontrar_relatorio_recente()
    print(f"ETAPA 1: Processando o arquivo '{os.path.basename(arquivo_entrada)}'")

    colunas = ['N° CT-e', 'Notas Fiscais', 'Cidade origem', 'CPF/CNPJ Remetente', 
               'Cidade destino', 'CPF/CNPJ Destinatário', 'Data Frete', 'Data Entrega']
    df = pd.read_excel(arquivo_entrada, usecols=colunas)

    df['Data Frete'] = pd.to_datetime(df['Data Frete'], dayfirst=True)
    df['Data Entrega'] = pd.to_datetime(df['Data Entrega'], dayfirst=True, errors='coerce')

    new_rows = []
    for index, row in df.iterrows():
        notas_fiscais = str(row['Notas Fiscais']).split(', ')
        for nota in notas_fiscais:
            new_row = row.copy()
            new_row['Notas Fiscais'] = nota
            new_rows.append(new_row)

    df_processado = pd.DataFrame(new_rows)

    df_processado['CPF/CNPJ Remetente'] = pd.to_numeric(df_processado['CPF/CNPJ Remetente'], errors='coerce')
    df_processado['CPF/CNPJ Destinatário'] = pd.to_numeric(df_processado['CPF/CNPJ Destinatário'], errors='coerce')

    print("✔ ETAPA 1: Tratamento inicial concluído.")

    # ETAPA 2: Aplicar a lógica de mapeamento (DE-PARA)
    print("\nETAPA 2: Aplicando mapeamento de CNPJ e Cidades para códigos...")

    map_cnpj = {
        '2012862022996': 'CML', '2012862002294': 'GRU', '2012862001050': 'SDU',
        '2012862002456': 'GIG', '2012862006109': 'MRO', '2012862000917': 'CGH'
    }
    map_cidades = {
        'ARACAJU': 'AJU', 'BELÉM': 'BEL', 'BOA VISTA': 'BVB', 'BRASÍLIA': 'BSB', 'CAMPO GRANDE': 'CGR',
        'CONFINS': 'CNF', 'CURITIBA': 'CWB', 'FLORIANÓPOLIS': 'FLN', 'FORTALEZA': 'FOR',
        'FOZ DO IGUACU': 'IGU', 'GOIANIA': 'GYN', 'ILHEUS': 'IOS', 'IMPERATRIZ': 'IMP',
        'JAGUARUNA': 'JJG', 'BAYEUX': 'JPA', 'JOINVILLE': 'JOI', 'LONDRINA': 'LDB', 'MACAPA': 'MCP',
        'MANAUS': 'MAO', 'MARABA': 'MAB', 'MARINGA': 'MGF', 'NAVEGANTES': 'NVT', 'PALMAS': 'PMW',
        'PORTO ALEGRE': 'POA', 'PORTO SEGURO': 'BPS', 'PORTO VELHO': 'PVH', 'RECIFE': 'REC',
        'RIBEIRAO PRETO': 'RAO', 'RIO BRANCO': 'RBR', 'RIO LARGO': 'MCZ',
        'SAO JOSE DO RIO PRETO': 'SJP', 'SALVADOR': 'SSA', 'SANTAREM': 'STM', 'SAO LUIS': 'SLZ',
        'TERESINA': 'THE', 'UBERLANDIA': 'UDI', 'VARZEA GRANDE': 'CGB', 'CAMPINAS': 'VCP',
        'VITORIA': 'VIX', 'CHAPECO': 'XAP', 'SINOP': 'OPS', 'AGUA VERMELHA (SAO CARLOS)': 'MRO',
        'SÃO GONÇALO DO AMARANTE': 'NAT', 'SÃO JOSÉ DOS PINHAIS': 'CWB', 'GUARULHOS': 'CML',
        'SÃO PAULO': 'CGH', 'SÃO CARLOS': 'MRO'
    }
    
    df_codificado = df_processado.copy()
    
    df_codificado['CPF/CNPJ Remetente'] = df_codificado['CPF/CNPJ Remetente'].astype('Int64').astype(str)
    df_codificado['CPF/CNPJ Destinatário'] = df_codificado['CPF/CNPJ Destinatário'].astype('Int64').astype(str)
    
    df_codificado['Origem_Codificada'] = df_codificado['CPF/CNPJ Remetente'].map(map_cnpj)
    df_codificado['Destino_Codificada'] = df_codificado['CPF/CNPJ Destinatário'].map(map_cnpj)

    map_cidade_origem = df_codificado['Cidade origem'].str.strip().str.upper().replace(map_cidades)
    map_cidade_destino = df_codificado['Cidade destino'].str.strip().str.upper().replace(map_cidades)
    
    df_codificado['Origem_Codificada'] = df_codificado['Origem_Codificada'].fillna(map_cidade_origem)
    df_codificado['Destino_Codificada'] = df_codificado['Destino_Codificada'].fillna(map_cidade_destino)
    
    print("✔ ETAPA 2: Mapeamento concluído.")

    # ETAPA 3: REORGANIZAR, RENOMEAR E SALVAR O ARQUIVO FINAL
    print("\nETAPA 3: Reorganizando colunas e salvando o arquivo final...")
    
    df_final = pd.DataFrame()

    df_final['N° OC'] = df_codificado['N° CT-e']
    df_final['Nft'] = df_codificado['Notas Fiscais']
    df_final['Origem'] = df_codificado['Origem_Codificada']
    df_final['CNPJ ORIGEM'] = df_codificado['CPF/CNPJ Remetente']
    df_final['Destino'] = df_codificado['Destino_Codificada']
    df_final['CNPJ DESTINO'] = df_codificado['CPF/CNPJ Destinatário']
    df_final['Data inclusão'] = df_codificado['Data Frete']
    df_final['Data expedida'] = df_codificado['Data Frete']
    df_final['Data chegada'] = df_codificado['Data Entrega']

    data_processamento = datetime.now().strftime('%d-%m-%Y_%Hh%M')
    arquivo_codificado = f'Relatório Tratado.xlsx'
    
    with pd.ExcelWriter(arquivo_codificado,
                        engine='xlsxwriter',
                        datetime_format='d/m/yy h:mm') as writer:
        
        df_final.to_excel(writer, sheet_name='Sheet1', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        formato_inteiro = workbook.add_format({'num_format': '0'})
        worksheet.set_column('D:D', 18, formato_inteiro)
        worksheet.set_column('F:F', 18, formato_inteiro)
        
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 10)
        worksheet.set_column('E:E', 10)
        worksheet.set_column('G:I', 15)

    print(f"✔ ETAPA 3: Arquivo final salvo como '{arquivo_codificado}'")
    print(f"\nProcesso completo finalizado com sucesso!")
    print(f"Caminho do arquivo final: {os.path.abspath(arquivo_codificado)}")

except FileNotFoundError as e:
    print(f"\n❌ ERRO: {e}")
except Exception as e:
    print(f"\n❌ Ocorreu um erro inesperado: {e}")