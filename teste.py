import pandas as pd
import openpyxl
import numpy as np
from modulo import dateVerification

"""# Define as opções de exibição
pd.set_option('display.max_columns', None)  # Exibe todas as colunas
pd.set_option('display.max_rows', None)  # Exibe todas as linhas
pd.set_option('display.width', 1000)  # Largura máxima da exibição

# Lê os DataFrames do arquivo Excel
df_lista_RDMarcas = pd.read_excel('storage/listaRDMARCAS.xlsx')
# df_lista_Perfumaria = pd.read_excel('storage/listaPERFUMARIA')
# df_lista_Dermo = pd.read_excel('storage/listaDERMO.xlsx')

# Soma os valores
df_lista_RDMarcas['Meta'] = df_lista_RDMarcas['Meta'].astype(str).str.replace('R$', '').replace(',', '.').astype(float)
metaAC = df_lista_RDMarcas['Meta'].sum().astype(float)


df_lista_RDMarcas['Venda'] = df_lista_RDMarcas['Venda'].astype(str).str.replace('R$', '').replace(',', '.').astype(float)
vendaAC = df_lista_RDMarcas['Venda'].sum().astype(float)

print(metaAC, vendaAC)

data_input = "10/05/2002"
metaDia_input = 12710.75
metaAC = metaAC + float(metaDia_input)
vendasDia_input = 49700.64
vendaAC = vendaAC + float(vendasDia_input)

if vendaAC < metaAC:
    sobras = metaAC - vendaAC
    devedor = '-'
elif metaAC < vendaAC:
    sobras = vendaAC - metaAC
else:
    sobras = 0

porcentagem = (vendaAC / metaAC) * 100

# Cria uma linha de novos dados
novoDado = pd.DataFrame([[data_input, metaDia_input, metaAC, vendasDia_input, vendaAC, sobras, porcentagem]],
                        columns=df_lista_RDMarcas.columns)


# Concatela com os dados anteriores da lista
df_lista_RDMarcas = pd.concat([df_lista_RDMarcas, novoDado], ignore_index=True)

# Tentativa de formatação
df_lista_RDMarcas['Meta'] = df_lista_RDMarcas['Meta'].map('{:.2f}'.format)
df_lista_RDMarcas['Meta.AC'] = df_lista_RDMarcas['Meta.AC'].map('{:.2f}'.format)
df_lista_RDMarcas['Venda'] = df_lista_RDMarcas['Venda'].map('{:.2f}'.format)
df_lista_RDMarcas['Venda.AC'] = df_lista_RDMarcas['Venda.AC'].map('{:.2f}'.format)
df_lista_RDMarcas['P'] = df_lista_RDMarcas['P'].apply(lambda x: '{:.2f}%'.format(x) if isinstance(x, (int, float)) else x)


df_lista_RDMarcas['Sobras'] = np.where(df_lista_RDMarcas['Venda.AC'] < df_lista_RDMarcas['Meta.AC'],
                                       "-" + df_lista_RDMarcas['Sobras'].apply('{:.2f}'.format),
                                       df_lista_RDMarcas['Sobras'].apply('{:.2f}'.format))



# Salva as alterações
df_lista_RDMarcas.to_excel('storage/listaRDMARCAS.xlsx', index=False)
"""


calc_lista_RDMarcas = pd.read_excel('storage/lista_calc_RDMarcas.xlsx')
max_lines = len(calc_lista_RDMarcas)
index_value = 1
"""metaDia = calc_lista_RDMarcas.at[index_value, 'Meta']"""
metaDia = calc_lista_RDMarcas.loc[5, 'Meta']
print(metaDia)
"""for _, linha in calc_lista_RDMarcas.iterrows():
    print(linha)"""

