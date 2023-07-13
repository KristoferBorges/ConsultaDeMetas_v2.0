# Cores
import sys
import random
import datetime
from kivy.uix.boxlayout import BoxLayout
from time import sleep
import platform
import pandas as pd
from openpyxl import load_workbook
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.button import Button
from kivy.uix.label import Label


red = '\033[31m'
green = '\033[32m'
yellow = '\033[33m'
ciano = '\033[36m'
normal = '\033[m'
roxo = '\033[35m'
rosa = '\033[95m'


def dateVerification(data):
    """
    --> Pedido do cliente: 'Preciso que o sistema pegue o dia de ontem baseado na data atual'.
    --> Sistema: Ele pegará a data atual e fara uma formatação dos dois primeiros números, fazendo com que seja a data
    de ontem.
    Exemplo: 15/03/2021 = 14/03/2021
    Teste ON: Caso a variável {teste} for verdadeira, o sistema pegará aleatóriamente um número entre 1 e 28 para
    preencher o dia.
    :return: Retornará uma string com a data de ontem
    """
    if len(data) == 0:
        data = datetime.datetime.now()
        data = datetime.datetime.date(data)
        data = data.strftime("%d/%m/%Y")
        data_soma = int(data[:2])
        data_soma = data_soma - 1
        data_soma = str(data_soma)
        if len(data_soma) == 1:
            data_soma = '0' + data_soma
        data = str(data_soma) + data[2:]
        return data
        # Data de ontem formatada
    else:
        data = str(data)
        return data


def tryOption(option):
    """
    --> Função para determinar a opção que o usuário irá escolher no menu, atualmente temos 6 opções, de registro,
    consulta ou de sair do programa, caso o usuário digitar outro número ou letras fora do esperado o sistema
    apresentará uma exceção.
    :param option: Entrada do usuário.
    :return: Retorna a opção escolhida em caso de zero exceções.
    """
    try:
        option = int(option)
    except ValueError:
        if option != 'adm':
            print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)
    except Exception as error:
        print(red + f' [!] - ERRO DE {error.__class__}' + normal)
    else:
        if option in range(1, 7):
            return option
        elif option == 'adm':
            return option
        else:
            print(red + ' [!] - APENAS NÚMEROS DAS OPÇÕES LISTADAS!' + normal)


def tryOptionList(option):
    """
    --> Função para determinar a opção que o usuário irá escolher no menu, atualmente temos 3 opções, de registro,
    consulta ou de sair do programa, caso o usuário digitar outro número ou letras fora do esperado o sistema
    apresentará uma exceção.
    :param option: Entrada do usuário.
    :return: Retorna a opção escolhida em caso de zero exceções.
    """
    try:
        option = int(option)
    except ValueError:
        if option != 'adm':
            print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)
    except Exception as error:
        print(red + f' [!] - ERRO DE {error.__class__}' + normal)
    else:
        if option in range(1, 4):
            return option
        elif option == 'adm':
            return option
        else:
            print(red + ' [!] - APENAS NÚMEROS DAS OPÇÕES LISTADAS!' + normal)


def tryOptionConsult(option):
    """
    --> Função para determinar a opção que o usuário irá escolher no menu, atualmente temos 4 opções, de consulta,
    caso o usuário digitar outro número ou letras fora do esperado o sistema apresentará uma exceção.
    :param option: Entrada do usuário.
    :return: Retorna a opção escolhida em caso de zero exceções.
    """
    try:
        option = int(option)
    except ValueError:
        print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)
    except Exception as error:
        print(red + f' [!] - ERRO DE {error.__class__}' + normal)
    else:
        if option in range(1, 5):
            return option
        else:
            print(red + ' [!] - APENAS NÚMEROS INTEIROS DAS OPÇÕES LISTADAS!' + normal)


def tryOptionBackup(option):
    """
    --> Função para determinar a opção que o usuário irá escolher no menu, atualmente temos 4 opções, de Backup,
    caso o usuário digitar outro número ou letras fora do esperado o sistema apresentará uma exceção.
    :param option: Entrada do usuário.
    :return: Retorna a opção escolhida em caso de zero exceções.
    """
    try:
        option = int(option)
    except ValueError:
        print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)
    except Exception as error:
        print(red + f' [!] - ERRO DE {error.__class__}' + normal)
    else:
        if option in range(1, 5):
            return option
        else:
            print(red + ' [!] - APENAS NÚMEROS INTEIROS DAS OPÇÕES LISTADAS!' + normal)


def tryExclusion(option):
    """
    --> Função para determinar a opção que o usuário irá escolher no menu, atualmente temos 4 opções, poderá excluir as
    listas separadamente ou todas ao mesmo tempo, caso o usuário digitar outro número ou letras fora do esperado o
    sistema apresentará uma exceção.
    :param option: Entrada do usuário.
    :return: Retorna a opção escolhida em caso de zero exceções.
    """
    try:
        option = int(option)
    except ValueError:
        print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)
    except Exception as error:
        print(red + f' [!] - ERRO DE {error.__class__}' + normal)
    else:
        if option in range(1, 5):
            return option
        else:
            print(red + ' [!] - APENAS NÚMEROS INTEIROS DAS OPÇÕES LISTADAS!' + normal)


def verificar_numero(valor):
    """
    --> Sub-Função para verificar se o input do usuário foi um número.
    :param valor: input do arquivo main.py
    :return: Valor verdadeiro ou falso
    """
    try:
        numero = float(valor)
        return True
    except ValueError:
        return False


def capturaDeValoresMetaDia():
    """
    --> Função simples para identificar se o usuário digitou um número inteiro válido para as metas do dia.
    :return: Retorna o número em caso de digitar corretamente.
    """
    try:
        while True:
            metaDia = input(green + ' [?] - Qual a Meta do Dia R$ ' + normal)
            if verificar_numero(metaDia):
                metaDia = float(metaDia)
                return metaDia
            else:
                print(red + ' [!] - Valor inválido. Insira um número válido.' + normal)
    except ValueError:
        print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)


def capturaDeValoresVendaDia():
    """
    --> Função simples para identificar se o usuário digitou um número inteiro válido para as vendas do dia.
    :return: Retorna o número em caso de digitar corretamente.
    """
    try:
        while True:
            vendaDia = input(green + ' [?] - Qual a Venda do Dia R$ ' + normal)
            if verificar_numero(vendaDia):
                vendaDia = float(vendaDia)
                return vendaDia
            else:
                print(red + ' [!] - Valor inválido. Insira um número válido.' + normal)
    except ValueError:
        print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)


def capturaDeValoresPecaDia():
    """
        --> Função simples para identificar se o usuário digitou um número inteiro válido para as peças do dia.
        :return: Retorna o número em caso de digitar corretamente.
        """
    try:
        while True:
            pecaDia = input(green + ' [?] - Quantas peças Vendeu Hoje: ' + normal)
            if pecaDia.isnumeric():
                pecaDia = int(pecaDia)
                return pecaDia
            else:
                print(red + ' [!] - Valor inválido. Insira um número válido.' + normal)
    except ValueError:
        print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)


def tryIsNumber_pecas(valor):

    """
        --> Função simples para identificar se o usuário digitou um número inteiro válido no caso das peças.
        :param valor:
        :return: Retorna o número em caso de digitar corretamente.
        """
    try:
        pass
    except ValueError:
        print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)
    except Exception as error:
        print(red + f' [!] - ERRO DE {error.__class__}' + normal)
    else:
        if valor.isnumeric():
            return int(valor)
        else:
            print(red + ' [!] - DIGITE UM NÚMERO INTEIRO VÁLIDO!' + normal)
            sys.exit()


def abatimento(meta, vendas):
    """
    Função para verificar se as metas foram atingidas, caso não tenham sido, o sistema apresentará o sinal de menos "-".
    :param meta: Valor da meta
    :param vendas: Valor das vendas
    :return: Um sinal negativo para em casos de não atingir as metas.
    """
    if vendas >= meta:
        devedor = ""
    else:
        devedor = "-"
    return devedor


def popupError():
    content = BoxLayout(orientation='vertical', padding=10)
    label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
    close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

    content.add_widget(label)
    content.add_widget(close_button)

    popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                  size=(375, 200))
    close_button.bind(on_release=popup.dismiss)
    popup.open()


def popup_Confirmacao_Exclusao():
    content = BoxLayout(orientation='vertical', padding=10)
    label = Label(text="Exclusão Realizada com Sucesso!")
    confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

    content.add_widget(label)
    content.add_widget(confirm_button)

    popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))
    confirm_button.bind(on_release=popup.dismiss)
    popup.open()


def popup_Confirmacao_Backup():
    content = BoxLayout(orientation='vertical', padding=10)
    label = Label(text="Backup Realizado com Sucesso!")
    confirm_button = Button(text='OK', size_hint=(None, None), size=(313, 50))

    content.add_widget(label)
    content.add_widget(confirm_button)

    popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))
    confirm_button.bind(on_release=popup.dismiss)
    popup.open()
