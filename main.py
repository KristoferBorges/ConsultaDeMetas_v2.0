import datetime
import platform
from modulo import dateVerification
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
        data_input = self.ids.data_input.text
        meta_input = self.ids.meta_input.text
        venda_input = self.ids.venda_input.text

        data_input = str(data_input)
        meta_input = str(meta_input)
        venda_input = str(venda_input)

        if data_input == '' or meta_input == '' or venda_input == '':
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(400, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class RegistrosPerfumaria(Screen):
    """
    Opção do menu principal após clicar na opção de registros (Perfumaria).
    """

    def avisoInput(self):
        data_input = self.ids.data_input.text
        meta_input = self.ids.data_input.text
        venda_input = self.ids.venda_input.text

        data_input = str(data_input)
        meta_input = str(meta_input)
        venda_input = str(venda_input)

        if data_input == '' or meta_input == '' or venda_input == '':
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(400, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


class RegistrosDermo(Screen):
    """
    Opção do menu principal após clicar na opção de registros (Dermo).
    """
    def avisoInput(self):
        data_input = self.ids.data_input.text
        meta_input = self.ids.data_input.text
        venda_input = self.ids.venda_input.text
        peca_input = self.ids.peca_input.text

        data_input = str(data_input)
        meta_input = str(meta_input)
        venda_input = str(venda_input)
        peca_input = str(peca_input)

        if data_input == '' or meta_input == '' or venda_input == '' or peca_input == '':
            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Campos não preenchidos!')
            close_button = Button(text='Fechar', size_hint=(None, None), size=(100, 50))

            content.add_widget(label)
            content.add_widget(close_button)

            popup = Popup(title='Aviso', content=content, size_hint=(None, None), size=(400, 200))
            close_button.bind(on_release=popup.dismiss)
            popup.open()


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

    @staticmethod
    def pega_input_rdmarcas(data, meta, vendas):
        """
        --> Função para pegar os dados inseridos na opção 'REGISTROS' -> 'RDMARCAS'.
        :param data: Valor de Data inserida na interface
        :param meta: Valor de Meta inserida na interface
        :param vendas: Valor de Venda inserida na interface
        :return: Retorna os dados devidamente formatados.
        """
        try:
            meta = float(meta)
            vendas = float(vendas)
            data = dateVerification(data)
            if vendas >= meta:
                result = "META ATINGIDA"
            else:
                result = "META NÃO ATINGIDA"
            return print(f"DATA: {data}\nMETA: {meta}\nVENDAS: {vendas}\nRESULTADO: {result}")

        except Exception as error:
            print(error)


if __name__ == '__main__':
    Tela().run()
