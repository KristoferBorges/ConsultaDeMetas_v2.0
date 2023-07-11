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


"""df_lista_RDMarcas = pd.read_excel('storage/listaRDMarcas.xlsx')
busca = 5  # Filtra pela data
# busca = 3 # Filtra pela linha
if len(str(busca)) > 4:
    linha_filtrada = df_lista_RDMarcas[df_lista_RDMarcas['Data'] == busca]
else:
    busca = busca - 2
    linha_filtrada = df_lista_RDMarcas[df_lista_RDMarcas.index == busca]

index_value = linha_filtrada.index[0]
print(index_value)
max_lines = len(df_lista_RDMarcas)
print(max_lines)"""


"""class Pai:
    def __init__(self):
        self.x = None

    def filho(self):
        self.x = 10
        return self.x

    def filha(self):
        print(self.x)


pessoa = Pai()
pessoa.filho()
pessoa.filha()"""

df_lista_RDMarcas = pd.read_excel('storage/listaRDMarcas.xlsx')
busca = '40/80/8000'
linha_filtrada = df_lista_RDMarcas[df_lista_RDMarcas['Data'] == busca]
num_linhas_data = df_lista_RDMarcas.loc[linha_filtrada.index[0]:, 'Data'].shape[0]
print(f'Linha escolhida: {num_linhas_data}')

index_value = 4 - 2
num_linhas_index = df_lista_RDMarcas.loc[index_value:, 'Data'].shape[0]
print(f'Linhas restantes: {num_linhas_index}')

