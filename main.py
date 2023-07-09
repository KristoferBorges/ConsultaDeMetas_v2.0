from time import sleep
import platform
import pandas as pd
from modulo import dateVerification, abatimento
from openpyxl import load_workbook
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.button import Button
from kivy.uix.label import Label


# Verifica se o usuário está usando Windows
if platform.system() == "Windows":
    sistema_windows = True
else:
    sistema_windows = False

# Variável para testar inserções de dados
teste = False


class MenuPrincipal(Screen):
    """
    Menu com as opções principais, NovosRegistros, LimparDados, ConsultaDeListas,
    Criar Backup e FecharPrograma.
    """
    pass


class NovosRegistros(Screen):
    """
    Opção para inserir novos dados, porém se faz necessário escolher as
    lista que deseja inserir os dados,
    poderá escolher entre: RD Marcas, Perfumaria e Dermo.
    """
    pass


class RegistrosRDMarcas(Screen):
    """
    Opção do menu principal após clicar na opção de registros (RDMarcas).
    """

    def pega_input_rdmarcas(self):
        """
        --> Função para pegar os dados inseridos na opção 'REGISTROS' -> 'RDMARCAS'.
        :return: Retorna os dados devidamente formatados.
        """
        try:
            data = self.ids.data_input.text
            metaDia = self.ids.meta_input.text
            vendaDia = self.ids.venda_input.text

            metaDia = float(metaDia)
            vendaDia = float(vendaDia)
            data = dateVerification(data)

            # Exibição no terminal
            pd.set_option('display.max_columns', None)  # Exibe todas as colunas
            pd.set_option('display.max_rows', None)  # Exibe todas as linhas
            pd.set_option('display.width', 1000)  # Largura máxima da exibição

            # Carrega o arquivo
            df_lista_RDMarcas = pd.read_excel('storage/listaRDMarcas.xlsx')
            calc_lista_RDMarcas = pd.read_excel('storage/lista_calc_RDMarcas.xlsx')

            novoCalc = {
                'Meta': f'{metaDia}',
                'Venda': f'{vendaDia}',
            }

            # Insere o dado na lista
            calc_lista_RDMarcas.loc[len(calc_lista_RDMarcas)] = novoCalc

            metaAC = calc_lista_RDMarcas['Meta'].astype(float).sum()
            vendaAC = calc_lista_RDMarcas['Venda'].astype(float).sum()

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

            # Análise alcance de metas
            devedor = abatimento(metaAC, vendaAC)

            # Cria uma linha de inserção de dados
            novoDado = {
                'Data': f'{data}',
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
            print(f'Meta: {metaAC}')
            print(f'Venda: {vendaAC}')

            # Salva o arquivo com o novo dado
            df_lista_RDMarcas.to_excel('storage/listaRDMarcas.xlsx', index=False)
            calc_lista_RDMarcas.to_excel('storage/lista_calc_RDMarcas.xlsx', index=False)

            # Limpa os dados anteriormente informados (Somente teste = False)
            if not teste:
                self.ids.data_input.text = ""
                self.ids.meta_input.text = ""
                self.ids.venda_input.text = ""

                # Baseado na variável (devedor) o sistema passará a situação da meta/vendas no popup
                if devedor == '-':
                    situacao = "Metas não atingidas!"
                else:
                    situacao = "Metas atingitas!"

                # Popup de resumo
                content = BoxLayout(orientation='vertical', padding=10)
                label = Label(text=f'Resumo Acumulado (RD-Marcas)\n\n'
                                   f'Meta: R$ {metaAC:.2f}\nVendas: R$ {vendaAC:.2f}\n'
                                   f'Sobras: {devedor}R$ {sobras:.2f}\nSituação: {situacao}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except ValueError as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                          size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class RegistrosPerfumaria(Screen):
    """
    Opção do menu principal após clicar na opção de registros (Perfumaria).
    """

    def pega_input_perfumaria(self):
        """
        --> Função para pegar os dados inseridos na opção 'REGISTROS' -> 'PERFUMARIA'.
        :return: Retorna os dados devidamente formatados.
        """
        try:
            data = self.ids.data_input.text
            metaDia = self.ids.meta_input.text
            vendaDia = self.ids.venda_input.text

            metaDia = float(metaDia)
            vendaDia = float(vendaDia)
            data = dateVerification(data)

            # Exibição no terminal
            pd.set_option('display.max_columns', None)  # Exibe todas as colunas
            pd.set_option('display.max_rows', None)  # Exibe todas as linhas
            pd.set_option('display.width', 1000)  # Largura máxima da exibição

            # Carrega o arquivo
            df_lista_Perfumaria = pd.read_excel('storage/listaPerfumaria.xlsx')
            calc_lista_Perfumaria = pd.read_excel('storage/lista_calc_Perfumaria.xlsx')

            novoCalc = {
                'Meta': f'{metaDia}',
                'Venda': f'{vendaDia}',
            }

            # Insere o dado na lista
            calc_lista_Perfumaria.loc[len(calc_lista_Perfumaria)] = novoCalc

            metaAC = calc_lista_Perfumaria['Meta'].astype(float).sum()
            vendaAC = calc_lista_Perfumaria['Venda'].astype(float).sum()

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

            # Análise alcance de metas
            devedor = abatimento(metaAC, vendaAC)

            # Cria uma linha de inserção de dados
            novoDado = {
                'Data': f'{data}',
                'Meta': f'{metaDia:.2f}',
                'Meta.AC': f'{metaAC:.2f}',
                'Venda': f'{vendaDia:.2f}',
                'Venda.AC': f'{vendaAC:.2f}',
                'Sobras': f'{sobras:.2f}',
                'P': f'{porcentagem:.2f}'
            }

            # Insere o dado na lista
            df_lista_Perfumaria.loc[len(df_lista_Perfumaria)] = novoDado
            print(df_lista_Perfumaria)
            print(f'Meta: {metaAC}')
            print(f'Venda: {vendaAC}')

            # Salva o arquivo com o novo dado
            df_lista_Perfumaria.to_excel('storage/listaPerfumaria.xlsx', index=False)
            calc_lista_Perfumaria.to_excel('storage/lista_calc_Perfumaria.xlsx', index=False)

            # Limpa os dados anteriormente informados (Somente teste = False)
            if not teste:
                self.ids.data_input.text = ""
                self.ids.meta_input.text = ""
                self.ids.venda_input.text = ""

                # Baseado na variável (devedor) o sistema passará a situação da meta/vendas no popup
                if devedor == '-':
                    situacao = "Metas não atingidas!"
                else:
                    situacao = "Metas atingitas!"

                # Popup de resumo
                content = BoxLayout(orientation='vertical', padding=10)
                label = Label(text=f'Resumo Acumulado (PERFUMARIA)\n\n'
                                   f'Meta: R$ {metaAC:.2f}\nVendas: R$ {vendaAC:.2f}\n'
                                   f'Sobras: {devedor}R$ {sobras:.2f}\nSituação: {situacao}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except ValueError as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                          size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class RegistrosDermo(Screen):
    """
    Opção do menu principal após clicar na opção de registros (Dermo).
    """

    def pega_input_dermo(self):
        """
        --> Função para pegar os dados inseridos na opção 'REGISTROS' -> 'DERMO'.
        :return: Retorna os dados devidamente formatados.
        """
        try:
            data = self.ids.data_input.text
            metaDia = self.ids.meta_input.text
            vendaDia = self.ids.venda_input.text
            pecaDia = self.ids.peca_input.text

            metaDia = float(metaDia)
            vendaDia = float(vendaDia)
            data = dateVerification(data)
            pecaDia = int(pecaDia)

            # Exibição no terminal
            pd.set_option('display.max_columns', None)  # Exibe todas as colunas
            pd.set_option('display.max_rows', None)  # Exibe todas as linhas
            pd.set_option('display.width', 1000)  # Largura máxima da exibição

            # Carrega o arquivo
            df_lista_Dermo = pd.read_excel('storage/listaDermo.xlsx')
            calc_lista_Dermo = pd.read_excel('storage/lista_calc_Dermo.xlsx')

            novoCalc = {
                'Meta': f'{metaDia}',
                'Venda': f'{vendaDia}',
                'Pecas': f'{pecaDia}',
            }

            # Insere o dado na lista
            calc_lista_Dermo.loc[len(calc_lista_Dermo)] = novoCalc

            metaAC = calc_lista_Dermo['Meta'].astype(float).sum()
            vendaAC = calc_lista_Dermo['Venda'].astype(float).sum()
            pecaAC = calc_lista_Dermo['Pecas'].astype(float).sum()

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

            # Análise alcance de metas
            devedor = abatimento(metaAC, vendaAC)

            # Cria uma linha de inserção de dados
            novoDado = {
                'Data': f'{data}',
                'Meta': f'{metaDia:.2f}',
                'Meta.AC': f'{metaAC:.2f}',
                'Venda': f'{vendaDia:.2f}',
                'Venda.AC': f'{vendaAC:.2f}',
                'Pecas.AC': f'{pecaAC}',
                'Sobras': f'{sobras:.2f}',
                'P': f'{porcentagem:.2f}'
            }

            # Insere o dado na lista
            df_lista_Dermo.loc[len(df_lista_Dermo)] = novoDado
            print(df_lista_Dermo)
            print(f'Meta: {metaAC}')
            print(f'Venda: {vendaAC}')
            print(f'Peças: {pecaAC}')

            # Salva o arquivo com o novo dado
            df_lista_Dermo.to_excel('storage/listaDermo.xlsx', index=False)
            calc_lista_Dermo.to_excel('storage/lista_calc_Dermo.xlsx', index=False)

            # Limpa os dados anteriormente informados (Somente teste = False)
            if not teste:
                self.ids.data_input.text = ""
                self.ids.meta_input.text = ""
                self.ids.venda_input.text = ""
                self.ids.peca_input.text = ""

                # Baseado na variável (devedor) o sistema passará a situação da meta/vendas no popup
                if devedor == '-':
                    situacao = "Metas não atingidas!"
                else:
                    situacao = "Metas atingitas!"

                # Popup de resumo
                content = BoxLayout(orientation='vertical', padding=10)
                label = Label(text=f'Resumo Acumulado (DERMO)\n\n'
                                   f'Meta: R$ {metaAC:.2f}\nVendas: R$ {vendaAC:.2f}\nPeças: {pecaAC} Un\n'
                                   f'Sobras: {devedor}R$ {sobras:.2f}\nSituação: {situacao}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except ValueError as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                          size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class LimparDados(Screen):
    """
    Opção para limpar dos dados, porém se faz necessário escolher as
    lista que deseja limpar os dados,
    poderá escolher entre: RD Marcas, Perfumaria, Dermo ou todas ao mesmo tempo.
    """

    def apagarLista_popup(self):
        """
        """

        content = BoxLayout(orientation='vertical', padding=10)
        label = Label(text='Confirma a exclusão das Listas?')

        close_button = Button(text='Cancelar', size_hint=(None, None), size=(313, 50))
        confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

        content.add_widget(label)
        content.add_widget(close_button)
        content.add_widget(confirm_button)

        popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))

        close_button.bind(on_release=popup.dismiss)

        confirm_button.bind(on_release=lambda btn: self.apagarLista())
        confirm_button.bind(on_release=popup.dismiss)
        popup.open()

    def apagarLista(self):
        """
        """

        try:
            # Exclusão TODAS AS LISTAS
            with open("storage/listaRDMARCAS.txt", "w") as listaRDMARCAS:
                listaRDMARCAS.write("")
            with open("storage/metaAcumuladaRDMARCAS.txt", "w") as metaAcumuladaRDMARCAS:
                metaAcumuladaRDMARCAS.write("")
            with open("storage/vendaAcumuladaRDMARCAS.txt", "w") as vendaAcumuladaRDMARCAS:
                vendaAcumuladaRDMARCAS.write("")

            with open("storage/listaPERFUMARIA.txt", "w") as listaPERFUMARIA:
                listaPERFUMARIA.write("")
            with open("storage/metaAcumuladaPERFUMARIA.txt", "w") as metaAcumuladaPERFUMARIA:
                metaAcumuladaPERFUMARIA.write("")
            with open("storage/vendaAcumuladaPERFUMARIA.txt", "w") as vendaAcumuladaPERFUMARIA:
                vendaAcumuladaPERFUMARIA.write("")

            with open("storage/listaDERMO.txt", "w") as listaDERMO:
                listaDERMO.write("")
            with open("storage/metaAcumuladaDERMO.txt", "w") as metaAcumuladaDERMO:
                metaAcumuladaDERMO.write("")
            with open("storage/vendaAcumuladaDERMO.txt", "w") as vendaAcumuladaDERMO:
                vendaAcumuladaDERMO.write("")
            with open("storage/pecaAcumuladaDERMO.txt", "w") as pecaAcumuladaDERMO:
                pecaAcumuladaDERMO.write("")

            # POPUP DE FINALIZAÇÃO
            sleep(0.3)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text="Exclusão Realizada com Sucesso!")
            confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

            content.add_widget(label)
            content.add_widget(confirm_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))
            confirm_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                          size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class LimparRD(Screen):
    """
    """

    def apagarLista_popup_RDMarcas(self):
        """
        """

        content = BoxLayout(orientation='vertical', padding=10)
        label = Label(text='Confirma a exclusão da Lista RD Marcas?')

        close_button = Button(text='Cancelar', size_hint=(None, None), size=(313, 50))
        confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

        content.add_widget(label)
        content.add_widget(close_button)
        content.add_widget(confirm_button)

        popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))

        close_button.bind(on_release=popup.dismiss)

        confirm_button.bind(on_release=lambda btn: self.apagarLista_RDMarca())
        confirm_button.bind(on_release=popup.dismiss)
        popup.open()

    def apagarLista_RDMarca(self):
        """
        """

        try:
            # Exclusão RD MARCAS
            # Carrega o arquivo
            lista = 'storage/listaRDMarcas.xlsx'
            calculo = 'storage/lista_calc_RDMarcas.xlsx'
            bk_lista = load_workbook(lista)
            bk_calculo = load_workbook(calculo)

            # Pega a primeira planilha do arquivo de lista
            sheet_lista = bk_lista.active

            # Pega a primeira planilha do arquivo de cálculo
            sheet_calculo = bk_calculo.active

            # Exclui as linhas tirando a primeira (Nome das colunas/Key)
            sheet_lista.delete_rows(2, sheet_lista.max_row)
            sheet_calculo.delete_rows(2, sheet_calculo.max_row)

            # Salva as alterações
            bk_lista.save(lista)
            bk_calculo.save(calculo)

            # Passa as alterações para uma variável
            df_lista_RDMarcas = pd.read_excel(lista)
            calc_lista_RDMarcas = pd.read_excel(calculo)

            # Salva o arquivo com as alterações no DataFrame
            df_lista_RDMarcas.to_excel('storage/listaRDMarcas.xlsx', index=False)
            calc_lista_RDMarcas.to_excel('storage/lista_calc_RDMarcas.xlsx', index=False)

            # POPUP DE FINALIZAÇÃO
            sleep(0.3)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text="Exclusão Realizada com Sucesso!")
            confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

            content.add_widget(label)
            content.add_widget(confirm_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))
            confirm_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                          size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class LimparPERFUMARIA(Screen):
    """
    """

    def apagarLista_popup_Perfumaria(self):
        """
        """

        content = BoxLayout(orientation='vertical', padding=10)
        label = Label(text='Confirma a exclusão da Lista Perfumaria?')

        close_button = Button(text='Cancelar', size_hint=(None, None), size=(313, 50))
        confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

        content.add_widget(label)
        content.add_widget(close_button)
        content.add_widget(confirm_button)

        popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))

        close_button.bind(on_release=popup.dismiss)

        confirm_button.bind(on_release=lambda btn: self.apagarLista_Perfumaria())
        confirm_button.bind(on_release=popup.dismiss)
        popup.open()

    def apagarLista_Perfumaria(self):
        """
        """
        try:
            # Exclusão PERFUMARIA
            # Carrega o arquivo
            lista = 'storage/listaPerfumaria.xlsx'
            calculo = 'storage/lista_calc_Perfumaria.xlsx'
            bk_lista = load_workbook(lista)
            bk_calculo = load_workbook(calculo)

            # Pega a primeira planilha do arquivo de lista
            sheet_lista = bk_lista.active

            # Pega a primeira planilha do arquivo de cálculo
            sheet_calculo = bk_calculo.active

            # Exclui as linhas tirando a primeira (Nome das colunas/Key)
            sheet_lista.delete_rows(2, sheet_lista.max_row)
            sheet_calculo.delete_rows(2, sheet_calculo.max_row)

            # Salva as alterações
            bk_lista.save(lista)
            bk_calculo.save(calculo)

            # Passa as alterações para uma variável
            df_lista_Perfumaria = pd.read_excel(lista)
            calc_lista_Perfumaria = pd.read_excel(calculo)

            # Salva o arquivo com as alterações no DataFrame
            df_lista_Perfumaria.to_excel('storage/listaPerfumaria.xlsx', index=False)
            calc_lista_Perfumaria.to_excel('storage/lista_calc_Perfumaria.xlsx', index=False)

            # POPUP DE FINALIZAÇÃO
            sleep(0.3)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text="Exclusão Realizada com Sucesso!")
            confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

            content.add_widget(label)
            content.add_widget(confirm_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))
            confirm_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                          size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class LimparDERMO(Screen):
    """
    """

    def apagarLista_popup_Dermo(self):
        """
        """

        content = BoxLayout(orientation='vertical', padding=10)
        label = Label(text='Confirma a exclusão da Lista Dermo?')

        close_button = Button(text='Cancelar', size_hint=(None, None), size=(313, 50))
        confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

        content.add_widget(label)
        content.add_widget(close_button)
        content.add_widget(confirm_button)

        popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))

        close_button.bind(on_release=popup.dismiss)

        confirm_button.bind(on_release=lambda btn: self.apagarLista_Dermo())
        confirm_button.bind(on_release=popup.dismiss)
        popup.open()

    def apagarLista_Dermo(self):
        """
        """

        try:
            # Exclusão Dermo
            # Carrega o arquivo
            lista = 'storage/listaDermo.xlsx'
            calculo = 'storage/lista_calc_Dermo.xlsx'
            bk_lista = load_workbook(lista)
            bk_calculo = load_workbook(calculo)

            # Pega a primeira planilha do arquivo de lista
            sheet_lista = bk_lista.active

            # Pega a primeira planilha do arquivo de cálculo
            sheet_calculo = bk_calculo.active

            # Exclui as linhas tirando a primeira (Nome das colunas/Key)
            sheet_lista.delete_rows(2, sheet_lista.max_row)
            sheet_calculo.delete_rows(2, sheet_calculo.max_row)

            # Salva as alterações
            bk_lista.save(lista)
            bk_calculo.save(calculo)

            # Passa as alterações para uma variável
            df_lista_Dermo = pd.read_excel(lista)
            calc_lista_Dermo = pd.read_excel(calculo)

            # Salva o arquivo com as alterações no DataFrame
            df_lista_Dermo.to_excel('storage/listaDermo.xlsx', index=False)
            calc_lista_Dermo.to_excel('storage/lista_calc_Dermo.xlsx', index=False)

            # POPUP DE FINALIZAÇÃO
            sleep(0.3)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text="Exclusão Realizada com Sucesso!")
            confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

            content.add_widget(label)
            content.add_widget(confirm_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))
            confirm_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Erro de execução!\n(O procedimento pode não ter sido realizado).')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Exceção encontrada (Abra um chamado)>', content=content, size_hint=(None, None),
                          size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class LimparTodasAsListas(Screen):
    """
    """
    pass


class ConsultaDeListas(Screen):
    """
    Opção para consultar os dados existentes, porém se faz necessário escolher as
    lista que deseja consultar,
    poderá escolher entre: RD Marcas, Perfumaria, Dermo ou todas ao mesmo tempo.
    """
    pass


class CriarBackup(Screen):
    """
    Opção para fazer backup dos dados existentes, porém se faz necessário escolher as
    lista que deseja fazer o backup,
    poderá escolher entre: RD Marcas, Perfumaria, Dermo ou todas ao mesmo tempo.
    """
    pass


class FecharPrograma(Screen):
    """
    Opção simples para fechar o programa com segurança sem medo de perder os
    dados ou interromper no meio do processo.
    """
    pass


class Tela(App):

    def build(self):
        if sistema_windows:
            Window.size = (379, 810)
        self.title = 'ConsultaDeMetas_v2.0'
        adm = ScreenManager()
        return adm


if __name__ == '__main__':
    Tela().run()
