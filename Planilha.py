import pandas as pd

caminho_planilha1 = 'C:/Users/camila.aguiar/Downloads/copia shopee.xlsx'
caminho_planilha2 = 'C:/Users/camila.aguiar/Downloads/Cópia de BID Nacional 2.0.xlsx'

aba_planilha1 = 'Rotas Bid - Agosto 2024'
aba_planilha2 = 'Plan1'

planilha1 = pd.read_excel(caminho_planilha1, sheet_name=aba_planilha1)
planilha2 = pd.read_excel(caminho_planilha2, sheet_name=aba_planilha2)

campos_para_transferir = {
    'Unnamed: 16': 'Unnamed: 16',
    'Unnamed: 17': 'Unnamed: 17',
    'Unnamed: 18': 'Unnamed: 18',
    'Unnamed: 19': 'Unnamed: 19',
    'Unnamed: 20': 'Unnamed: 20',
    'Unnamed: 21': 'Unnamed: 21',
    'Unnamed: 22': 'Unnamed: 22',
    'Unnamed: 23': 'Unnamed: 23',
    'Unnamed: 24': 'Unnamed: 24',
    'Unnamed: 25': 'Unnamed: 25',
    'Unnamed: 26': 'Unnamed: 26',
    'Unnamed: 27': 'Unnamed: 27'
}

colunas_para_juncao = list(campos_para_transferir.keys())
colunas_existentes = [col for col in colunas_para_juncao if col in planilha2.columns]
if len(colunas_existentes) != len(colunas_para_juncao):
    print("Aviso: Algumas colunas a serem transferidas não foram encontradas em planilha2.")
    print("Colunas encontradas:", colunas_existentes)

print("\nValores únicos da coluna de identificação na planilha1:")
print(planilha1['Unnamed: 1'].unique())

print("\nValores únicos da coluna de identificação na planilha2:")
print(planilha2['Unnamed: 1'].unique())

planilha1_atualizada = pd.merge(
    planilha1,
    planilha2[['Unnamed: 1'] + colunas_existentes],
    on='Unnamed: 1',
    how='left',
    suffixes=('', '_new')
)

for col_origem, col_destino in campos_para_transferir.items():
    col_origem_novo = col_origem + '_new'
    if col_origem_novo in planilha1_atualizada.columns:
        planilha1_atualizada[col_destino] = planilha1_atualizada[col_origem_novo]
        planilha1_atualizada.drop(columns=[col_origem_novo], inplace=True)

planilha1_atualizada.to_excel(caminho_planilha1, sheet_name=aba_planilha1, index=False)
print("Planilha atualizada salva com sucesso.")
