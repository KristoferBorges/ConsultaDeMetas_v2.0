import pandas as pd
import openpyxl


"""# Inclusão de dados
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
    print(error)"""


#
# Fazer as alterações das variávels conforme o uso dos inputs com o front-end
#
# Primeira forma de editar um arquivo
# Carregar o arquivo (O arquivo de calculo e o arquivo da lista de visualização)
calc_lista_RDMarcas = pd.read_excel('storage/lista_calc_RDMarcas.xlsx')
df_lista_RDMarcas = pd.read_excel('storage/listaRDMarcas.xlsx')

# Localiza a linha com base no input do usuário
busca = 2  # Filtra pela data
# busca = 3 # Filtra pela linha
if len(str(busca)) > 4:
    linha_filtrada = df_lista_RDMarcas[df_lista_RDMarcas['Data'] == busca]
else:
    busca = busca - 2
    linha_filtrada = df_lista_RDMarcas[df_lista_RDMarcas.index == busca]

index_value = linha_filtrada.index[0]
print(index_value)

# Define a quantidade de repetições iniciais
qnt = 1

# Verifica a quantidade máxima de linhas dentro do arquivo
max_lines = len(calc_lista_RDMarcas)

# Cálculo dos dados
metaDia = 500  # input
vendaDia = 500  # input

try:
    for linha in calc_lista_RDMarcas.iterrows():
        if qnt == 1:
            qnt = qnt + 1
            # Insere os valores (MetaDia/VendaDia), logo em seguida é feito o cálculo já pegando o valor alterado
            calc_lista_RDMarcas.loc[index_value] = [metaDia, vendaDia]

            metaAC = calc_lista_RDMarcas.loc[:index_value, 'Meta'].astype(float).sum()
            vendaAC = calc_lista_RDMarcas.loc[:index_value, 'Venda'].astype(float).sum()
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
                    'Data': '05/06/2023',
                    'Meta': f'{metaDia:.2f}',
                    'Meta.AC': f'{metaAC:.2f}',
                    'Venda': f'{vendaDia:.2f}',
                    'Venda.AC': f'{vendaAC:.2f}',
                    'Sobras': f'{sobras:.2f}',
                    'P': f'{porcentagem:.2f}'
                }

            # Modifica os valores da linha (MetaDia/VendaDia) | Modifica os dados da
            # linha por completo com os devidos cálculos
            df_lista_RDMarcas.loc[index_value] = novoDado

            # Salva o arquivo
            df_lista_RDMarcas.to_excel('storage/listaRDMarcas.xlsx', index=False)
            calc_lista_RDMarcas.to_excel('storage/lista_calc_RDMarcas.xlsx', index=False)

            # print(df_lista_RDMarcas.loc[:index_value])
        elif max_lines >= index_value:
            index_value = index_value + 1

            metaDia = calc_lista_RDMarcas.at[index_value, 'Meta']
            vendaDia = calc_lista_RDMarcas.at[index_value, 'Venda']
            data = df_lista_RDMarcas.at[index_value, 'Data']
            calc_lista_RDMarcas.loc[index_value] = [metaDia, vendaDia]

            metaAC = calc_lista_RDMarcas.loc[:index_value, 'Meta'].astype(float).sum()
            vendaAC = calc_lista_RDMarcas.loc[:index_value, 'Venda'].astype(float).sum()

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
                    'Data': f'{data}',
                    'Meta': f'{metaDia:.2f}',
                    'Meta.AC': f'{metaAC:.2f}',
                    'Venda': f'{vendaDia:.2f}',
                    'Venda.AC': f'{vendaAC:.2f}',
                    'Sobras': f'{sobras:.2f}',
                    'P': f'{porcentagem:.2f}'
                }

            df_lista_RDMarcas.loc[index_value] = novoDado

            # Salva o arquivo
            df_lista_RDMarcas.to_excel('storage/listaRDMarcas.xlsx', index=False)
            calc_lista_RDMarcas.to_excel('storage/lista_calc_RDMarcas.xlsx', index=False)

        else:
            print('Houve um problema')
except KeyError as error:
    print("Fim")
except Exception as error:
    print(error)

