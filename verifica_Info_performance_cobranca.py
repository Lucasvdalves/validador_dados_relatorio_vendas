import pandas as pd
import os

def comparar_planilhas(caminho_abr, caminho_suelen, caminho_saida):
    df_abril = pd.read_excel(caminho_abr, sheet_name='Abril', usecols='A,H')
    df_abril.columns = ['NUMCONTRATO', 'VLR_RENEG_BAIXADO']
    
    df_suelen = pd.read_excel(caminho_suelen, sheet_name='1_BASE', usecols='B,W')
    df_suelen.columns = ['NroContrato', 'ValorPago']

    df_abril = df_abril.dropna(subset=['NUMCONTRATO'])
    df_suelen = df_suelen.dropna(subset=['NroContrato'])

    df_abril['NUMCONTRATO'] = df_abril['NUMCONTRATO'].astype(str).str.strip()
    df_suelen['NroContrato'] = df_suelen['NroContrato'].astype(str).str.strip()

    df_comum = pd.merge(df_abril, df_suelen, left_on='NUMCONTRATO', right_on='NroContrato', how='inner')

    # Seleciona colunas finais
    df_resultado = df_comum[['NUMCONTRATO', 'VLR_RENEG_BAIXADO', 'ValorPago']]

    # Salva a planilha resultado
    nome_saida = os.path.join(caminho_saida, 'Contratos_Identificados.xlsx')
    df_resultado.to_excel(nome_saida, index=False)

    print(f'✅ Planilha gerada com sucesso: {nome_saida}')


comparar_planilhas(
    r'C: caminho_1',
    r'C: caminho_2',
    r'C: caminho_1'
)
