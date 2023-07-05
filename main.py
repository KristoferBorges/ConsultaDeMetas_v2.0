import platform
from modulo import dateVerification, abatimento
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
        global situacaoRD
        try:
            data = self.ids.data_input.text
            metaDia = self.ids.meta_input.text
            vendaDia = self.ids.venda_input.text

            metaDia = float(metaDia)
            vendaDia = float(vendaDia)
            data = dateVerification(data)

            metaAcRDMARCAS = 0
            vendaAcRDMARCAS = 0
            porcentagemRDMARCAS = 0

            # Cálculo de Metas acumuladas
            with open("storage/metaAcumuladaRDMARCAS.txt", "a") as metaAcumuladaRDMARCAS:
                metaAcumuladaRDMARCAS.write(f"{metaDia}\n")
            with open("storage/metaAcumuladaRDMARCAS.txt", "r") as metaAcumuladaRDMARCAS:
                linhas = metaAcumuladaRDMARCAS.readlines()

            for linha in linhas:
                metaAcRDMARCAS = metaAcRDMARCAS + float(linha.strip())

            # Cálculo de Vendas acumuladas
            with open("storage/vendaAcumuladaRDMARCAS.txt", "a") as vendaAcumuladaRDMARCAS:
                vendaAcumuladaRDMARCAS.write(f"{vendaDia}\n")
            with open("storage/vendaAcumuladaRDMARCAS.txt", "r") as vendaAcumuladaRDMARCAS:
                linhas2 = vendaAcumuladaRDMARCAS.readlines()

            for linha in linhas2:
                vendaAcRDMARCAS = vendaAcRDMARCAS + float(linha.strip())

            # Cálculo de porcentagem
            if vendaAcRDMARCAS < metaAcRDMARCAS:
                sobrasRD = (metaAcRDMARCAS - vendaAcRDMARCAS)
            elif metaAcRDMARCAS < vendaAcRDMARCAS:
                sobrasRD = (vendaAcRDMARCAS - metaAcRDMARCAS)
            else:
                sobrasRD = 0
            if vendaAcRDMARCAS != 0 and metaAcRDMARCAS != 0:
                porcentagemRDMARCAS = (vendaAcRDMARCAS / metaAcRDMARCAS) * 100

            # Análise alcance de metas
            devedor = abatimento(metaAcRDMARCAS, vendaAcRDMARCAS)
            # Inserção de dados
            with open("storage/listaRDMARCAS.txt", "a") as listaRDMARCAS:
                listaRDMARCAS.write(f"{data}|R${metaDia:.2f}|R${metaAcRDMARCAS:.2f}|R${vendaDia:.2f}|"
                                    f"R${vendaAcRDMARCAS:.2f}|"
                                    f"{devedor}R${sobrasRD:.2f}|"
                                    f"{porcentagemRDMARCAS:.2f}%\n")

            # Limpa os dados anteriormente informados (Somente teste = False)
            if not teste:
                self.ids.data_input.text = ""
                self.ids.meta_input.text = ""
                self.ids.venda_input.text = ""

                # Baseado na variável (devedor) o sistema passará a situação da meta/vendas no popup
                if devedor == '-':
                    situacaoRD = "Metas não atingidas!"
                elif devedor == '':
                    situacaoRD = "Metas atingitas!"

                # Popup de resumo
                content = BoxLayout(orientation='vertical', padding=10)
                label = Label(text=f'Resumo Acumulado (RD-Marcas)\n\n'
                                   f'Meta: R$ {metaAcRDMARCAS:.2f}\nVendas: R$ {vendaAcRDMARCAS:.2f}\n'
                                   f'Sobras: {devedor}R$ {sobrasRD:.2f}\nSituação: {situacaoRD}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(375, 200))
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
        global situacaoPERFUMARIA
        try:
            data = self.ids.data_input.text
            metaDia = self.ids.meta_input.text
            vendaDia = self.ids.venda_input.text

            metaDia = float(metaDia)
            vendaDia = float(vendaDia)
            data = dateVerification(data)

            metaAcPERFUMARIA = 0
            vendaAcPERFUMARIA = 0
            porcentagemPERFUMARIA = 0

            # Cálculo de Metas acumuladas
            with open("storage/metaAcumuladaPERFUMARIA.txt", "a") as metaAcumuladaPERFUMARIA:
                metaAcumuladaPERFUMARIA.write(f"{metaDia}\n")
            with open("storage/metaAcumuladaPERFUMARIA.txt", "r") as metaAcumuladaPERFUMARIA:
                linhas = metaAcumuladaPERFUMARIA.readlines()

            for linha in linhas:
                metaAcPERFUMARIA = metaAcPERFUMARIA + float(linha.strip())

            # Cálculo de Vendas acumuladas
            with open("storage/vendaAcumuladaPERFUMARIA.txt", "a") as vendaAcumuladaPERFUMARIA:
                vendaAcumuladaPERFUMARIA.write(f"{vendaDia}\n")
            with open("storage/vendaAcumuladaPERFUMARIA.txt", "r") as vendaAcumuladaPERFUMARIA:
                linhas2 = vendaAcumuladaPERFUMARIA.readlines()

            for linha in linhas2:
                vendaAcPERFUMARIA = vendaAcPERFUMARIA + float(linha.strip())

            # Cálculo de porcentagem
            if vendaAcPERFUMARIA < metaAcPERFUMARIA:
                sobrasPERFUMARIA = (metaAcPERFUMARIA - vendaAcPERFUMARIA)
            elif metaAcPERFUMARIA < vendaAcPERFUMARIA:
                sobrasPERFUMARIA = (vendaAcPERFUMARIA - metaAcPERFUMARIA)
            else:
                sobrasPERFUMARIA = 0
            if vendaAcPERFUMARIA != 0 and metaAcPERFUMARIA != 0:
                porcentagemPERFUMARIA = (vendaAcPERFUMARIA / metaAcPERFUMARIA) * 100

            # Análise alcance de metas
            devedor = abatimento(metaAcPERFUMARIA, vendaAcPERFUMARIA)
            # Inserção de dados
            with open("storage/listaPERFUMARIA.txt", "a") as listaPERFUMARIA:
                listaPERFUMARIA.write(f"{data}|R${metaDia:.2f}|R${metaAcPERFUMARIA:.2f}|R${vendaDia:.2f}|"
                                      f"R${vendaAcPERFUMARIA:.2f}|"
                                      f"{devedor}R${sobrasPERFUMARIA:.2f}|"
                                      f"{porcentagemPERFUMARIA:.2f}%\n")

            # Limpa os dados anteriormente informados (Somente teste = False)
            if not teste:
                self.ids.data_input.text = ""
                self.ids.meta_input.text = ""
                self.ids.venda_input.text = ""

                # Baseado na variável (devedor) o sistema passará a situação da meta/vendas no popup
                if devedor == '-':
                    situacaoPERFUMARIA = "Metas não atingidas!"
                elif devedor == '':
                    situacaoPERFUMARIA = "Metas atingitas!"

                # Popup de resumo
                content = BoxLayout(orientation='vertical', padding=10)
                label = Label(text=f'Resumo Acumulado (PERFUMARIA)\n\n'
                                   f'Meta: R$ {metaAcPERFUMARIA:.2f}\nVendas: R$ {vendaAcPERFUMARIA:.2f}\n'
                                   f'Sobras: {devedor}R$ {sobrasPERFUMARIA:.2f}\nSituação: {situacaoPERFUMARIA}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(375, 200))
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
        global situacaoDERMO
        try:
            data = self.ids.data_input.text
            metaDia = self.ids.meta_input.text
            vendaDia = self.ids.venda_input.text
            pecaDia = self.ids.peca_input.text

            metaDia = float(metaDia)
            vendaDia = float(vendaDia)
            data = dateVerification(data)
            pecaDia = int(pecaDia)

            metaAcDERMO = 0
            vendaAcDERMO = 0
            pecaAc = 0
            porcentagemDERMO = 0

            # Cálculo de Metas acumuladas
            with open("storage/metaAcumuladaDERMO.txt", "a") as metaAcumuladaDERMO:
                metaAcumuladaDERMO.write(f"{metaDia}\n")
            with open("storage/metaAcumuladaDERMO.txt", "r") as metaAcumuladaDERMO:
                linhas = metaAcumuladaDERMO.readlines()

            for linha in linhas:
                metaAcDERMO = metaAcDERMO + float(linha.strip())

            # Cálculo de Vendas acumuladas
            with open("storage/vendaAcumuladaDERMO.txt", "a") as vendaAcumuladaDERMO:
                vendaAcumuladaDERMO.write(f"{vendaDia}\n")
            with open("storage/vendaAcumuladaDERMO.txt", "r") as vendaAcumuladaDERMO:
                linhas2 = vendaAcumuladaDERMO.readlines()

            for linha in linhas2:
                vendaAcDERMO = vendaAcDERMO + float(linha.strip())

            # Cálculo de Peças acumuladas
            with open("storage/pecaAcumuladaDERMO.txt", "a") as pecaAcumuladaDERMO:
                pecaAcumuladaDERMO.write(f'{pecaDia}\n')
            with open("storage/pecaAcumuladaDERMO.txt", "r") as pecaAcumuladaDERMO:
                linhasPeca = pecaAcumuladaDERMO.readlines()

            for linha in linhasPeca:
                pecaAc = pecaAc + int(linha.strip())

            # Cálculo de porcentagem
            if vendaAcDERMO < metaAcDERMO:
                sobrasDERMO = (metaAcDERMO - vendaAcDERMO)
            elif metaAcDERMO < vendaAcDERMO:
                sobrasDERMO = (vendaAcDERMO - metaAcDERMO)
            else:
                sobrasDERMO = 0
            if vendaAcDERMO != 0 and metaAcDERMO != 0:
                porcentagemDERMO = (vendaAcDERMO / metaAcDERMO) * 100

            # Análise alcance de metas
            devedor = abatimento(metaAcDERMO, vendaAcDERMO)
            # Inserção de dados
            with open("storage/listaDERMO.txt", "a") as listaDERMO:
                listaDERMO.write(f"{data} | R${metaDia:.2f} | R${metaAcDERMO:.2f} | R${vendaDia:.2f} |"
                                 f" R${vendaAcDERMO:.2f} | "
                                 f" {pecaAc}Un | "
                                 f" {devedor}R${sobrasDERMO:.2f} | "
                                 f"{porcentagemDERMO:.2f}%\n")

            # Limpa os dados anteriormente informados (Somente teste = False)
            if not teste:
                self.ids.data_input.text = ""
                self.ids.meta_input.text = ""
                self.ids.venda_input.text = ""
                self.ids.peca_input.text = ""

                # Baseado na variável (devedor) o sistema passará a situação da meta/vendas no popup
                if devedor == '-':
                    situacaoDERMO = "Metas não atingidas!"
                elif devedor == '':
                    situacaoDERMO = "Metas atingitas!"

                # Popup de resumo
                content = BoxLayout(orientation='vertical', padding=10)
                label = Label(text=f'Resumo Acumulado (DERMO)\n\n'
                                   f'Meta: R$ {metaAcDERMO:.2f}\nVendas: R$ {vendaAcDERMO:.2f}\nPeças: {pecaAc} Un\n'
                                   f'Sobras: {devedor}R$ {sobrasDERMO:.2f}\nSituação: {situacaoDERMO}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except Exception as error:
            print(error)
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(375, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class LimparDados(Screen):
    """
    Opção para limpar dos dados, porém se faz necessário escolher as
    lista que deseja limpar os dados,
    poderá escolher entre: RD Marcas, Perfumaria, Dermo ou todas ao mesmo tempo.
    """

    def apagarLista_popup(self, button):
        lista = button.text

        content = BoxLayout(orientation='vertical', padding=10)
        if lista == "TODAS AS LISTAS":
            label = Label(text=f'Confirma a exclusão de "{lista}"?')
        else:
            label = Label(text=f'Confirma a excluão da lista "{lista}"?')

        close_button = Button(text='Cancelar', size_hint=(None, None), size=(313, 50))
        confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(313, 50))

        content.add_widget(label)
        content.add_widget(close_button)
        content.add_widget(confirm_button)

        popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(360, 280))

        close_button.bind(on_release=popup.dismiss)

        confirm_button.bind(on_release=lambda btn: self.apagarLista(lista))
        confirm_button.bind(on_release=popup.dismiss)
        popup.open()

    def apagarLista(self, lista):
        if lista == "RD MARCAS":
            # Exclusão RD MARCAS
            with open("storage/listaRDMARCAS.txt", "w") as listaRDMARCAS:
                listaRDMARCAS.write("")
            with open("storage/metaAcumuladaRDMARCAS.txt", "w") as metaAcumuladaRDMARCAS:
                metaAcumuladaRDMARCAS.write("")
            with open("storage/vendaAcumuladaRDMARCAS.txt", "w") as vendaAcumuladaRDMARCAS:
                vendaAcumuladaRDMARCAS.write("")

        elif lista == "PERFUMARIA":
            # Exclusão Perfumaria
            with open("storage/listaPERFUMARIA.txt", "w") as listaPERFUMARIA:
                listaPERFUMARIA.write("")
            with open("storage/metaAcumuladaPERFUMARIA.txt", "w") as metaAcumuladaPERFUMARIA:
                metaAcumuladaPERFUMARIA.write("")
            with open("storage/vendaAcumuladaPERFUMARIA.txt", "w") as vendaAcumuladaPERFUMARIA:
                vendaAcumuladaPERFUMARIA.write("")

        elif lista == "DERMO":
            # Exclusão Dermo
            with open("storage/listaDERMO.txt", "w") as listaDERMO:
                listaDERMO.write("")
            with open("storage/metaAcumuladaDERMO.txt", "w") as metaAcumuladaDERMO:
                metaAcumuladaDERMO.write("")
            with open("storage/vendaAcumuladaDERMO.txt", "w") as vendaAcumuladaDERMO:
                vendaAcumuladaDERMO.write("")
            with open("storage/pecaAcumuladaDERMO.txt", "w") as pecaAcumuladaDERMO:
                pecaAcumuladaDERMO.write("")

        elif lista == "TODAS AS LISTAS":
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


class LimparRD(Screen):
    """
    """
    pass


class LimparPERFUMARIA(Screen):
    """
    """
    pass


class LimparDERMO(Screen):
    """
    """
    pass


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
