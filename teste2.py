import pandas as pd
import openpyxl


# Inclusão de dados
try:
    pd.set_option('display.max_columns', None)  # Exibe todas as colunas
    pd.set_option('display.max_rows', None)  # Exibe todas as linhas
    pd.set_option('display.width', 1000)  # Largura máxima da exibição

    # Carrega o arquivo
    df_lista_RDMarcas = pd.read_excel('storage/listaRDMarcas.xlsx')
    calc_lista_RDMarcas = pd.read_excel('storage/calc_listas.xlsx')
    metaDia = 2000
    vendaDia = 6000

    novoCalc = {
        'Meta.RD': f'{metaDia}',
        'Venda.RD': f'{vendaDia}',
    }

    # Insere o dado na lista
    calc_lista_RDMarcas.loc[len(calc_lista_RDMarcas)] = novoCalc

    metaAC = calc_lista_RDMarcas['Meta.RD'].astype(float).sum()
    vendaAC = calc_lista_RDMarcas['Venda.RD'].astype(float).sum()

    if vendaAC < metaAC:
        sobras = (metaAC - vendaAC)
    elif metaAC < vendaAC:
        sobras = (vendaAC - metaAC)
    else:
        sobras = 0

    if vendaAC != 0 and metaAC != 0:
        porcentagem = (vendaAC / metaAC) * 100
    else:
        porcentagem = 'Error'

    # Cria uma linha de inserção de dados
    novoDado = {
        'Data': '2023-05-20',
        'Meta': f'{metaDia:.2f}',
        'Meta.AC': f'{metaAC:.2f}',
        'Venda': f'{vendaDia:.2f}',
        'Venda.AC': f'{vendaAC:.2f}',
        'Sobras': f'{sobras:.2f}',
        'P': f'{porcentagem:.2f}'
    }

    # Insere o dado na lista
    df_lista_RDMarcas.loc[len(df_lista_RDMarcas)] = novoDado
    print(df_lista_RDMarcas)
    print(f'Valor: {metaAC}')
    print(f'Valor: {vendaAC}')

    # Salva o arquivo com o novo dado
    df_lista_RDMarcas.to_excel('storage/listaRDMarcas.xlsx', index=False)
    calc_lista_RDMarcas.to_excel('storage/calc_listas.xlsx', index=False)
except ValueError as error:
    print(error)

"""# Carregar o arquivo (O arquivo de calculo e o arquivo da lista de visualização)
calc_lista_RDMarcas = pd.read_excel('storage/calc_listas.xlsx')
df_lista_RDMarcas = pd.read_excel('storage/listaRDMarcas.xlsx')

# Localizar a linha tendo que inserir 2 índices a menos para acertar a linha solicitada
index_value = (3 - 2)

# Cálculo dos dados
metaDia = 5000  # input
vendaDia = 5000  # input
metaAC = calc_lista_RDMarcas.loc[index_value:, 'Meta.RD'].astype(float).sum()
vendaAC = calc_lista_RDMarcas.loc[index_value:, 'Venda.RD'].astype(float).sum()
if vendaAC < metaAC:
    sobras = (metaAC - vendaAC)
elif metaAC < vendaAC:
    sobras = (vendaAC - metaAC)
else:
    sobras = 0

if vendaAC != 0 and metaAC != 0:
    porcentagem = (vendaAC / metaAC) * 100
else:
    porcentagem = 'Error'


# Input de dados
novoDado = {
        'Data': '2023-05-20',
        'Meta': f'{metaDia:.2f}',
        'Meta.AC': f'{metaAC:.2f}',
        'Venda': f'{vendaDia:.2f}',
        'Venda.AC': f'{vendaAC:.2f}',
        'Sobras': f'{sobras:.2f}',
        'P': f'{porcentagem:.2f}'
    }

# Modifica os valores da linha (MetaDia/VendaDia) | Modifica os dados da linha por completo com os devidos cálculos
calc_lista_RDMarcas.loc[index_value] = [metaDia, vendaDia]
df_lista_RDMarcas.loc[index_value] = novoDado

# Salva o arquivo
df_lista_RDMarcas.to_excel('storage/listaRDMarcas.xlsx', index=False)
calc_lista_RDMarcas.to_excel('storage/calc_listas.xlsx', index=False)"""
