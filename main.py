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

    def avisoInput(self):
        meta_input = self.ids.meta_input.text
        venda_input = self.ids.venda_input.text

        meta_input = str(meta_input)
        venda_input = str(venda_input)

        if meta_input == '' or venda_input == '':
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(400, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()

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
                label = Label(text=f'Resumo Acumulado (RD-Marcas)\n\nData: {data}\n'
                                   f'Meta: R$ {metaAcRDMARCAS}\nVendas: R$ {vendaAcRDMARCAS}\n\n'
                                   f'Sobras: {devedor}R$ {sobrasRD}\nSituação: {situacaoRD}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except Exception as error:
            print(error)


class RegistrosPerfumaria(Screen):
    """
    Opção do menu principal após clicar na opção de registros (Perfumaria).
    """

    def avisoInput(self):
        meta_input = self.ids.meta_input.text
        venda_input = self.ids.venda_input.text

        meta_input = str(meta_input)
        venda_input = str(venda_input)

        if meta_input == '' or venda_input == '':
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(400, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()

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
                label = Label(text=f'Resumo Acumulado (PERFUMARIA)\n\nData: {data}\n'
                                   f'Meta: R$ {metaAcPERFUMARIA}\nVendas: R$ {vendaAcPERFUMARIA}\n\n'
                                   f'Sobras: {devedor}R$ {sobrasPERFUMARIA}\nSituação: {situacaoPERFUMARIA}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except Exception as error:
            print(error)


class RegistrosDermo(Screen):
    """
    Opção do menu principal após clicar na opção de registros (Dermo).
    """

    def avisoInput(self):
        meta_input = self.ids.meta_input.text
        venda_input = self.ids.venda_input.text
        peca_input = self.ids.peca_input.text

        meta_input = str(meta_input)
        venda_input = str(venda_input)
        peca_input = str(peca_input)

        if meta_input == '' or venda_input == '' or peca_input == '':
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(400, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()

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
                label = Label(text=f'Resumo Acumulado (DERMO)\n\nData: {data}\n'
                                   f'Meta: R$ {metaAcDERMO}\nVendas: R$ {vendaAcDERMO}\nPeças: {pecaAc} Un\n'
                                   f'Sobras: {devedor}R$ {sobrasDERMO}\nSituação: {situacaoDERMO}\n')
                close_button = Button(text='Fechar', size_hint=(None, None), size=(313, 50))

                content.add_widget(label)
                content.add_widget(close_button)

                popup = Popup(title='Dados armazenados com Sucesso!', content=content, size_hint=(None, None),
                              size=(360, 280))
                close_button.bind(on_release=popup.dismiss)
                popup.open()

        except Exception as error:
            print(error)


class LimparDados(Screen):
    """
    Opção para limpar dos dados, porém se faz necessário escolher as
    lista que deseja limpar os dados,
    poderá escolher entre: RD Marcas, Perfumaria, Dermo ou todas ao mesmo tempo.
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
