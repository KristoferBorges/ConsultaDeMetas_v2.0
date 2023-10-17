import platform
from datetime import datetime
import pandas as pd
from kivy.properties import NumericProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import Screen
from app import db_rdmarcas, db_perfumaria, db_dermo, popupError, formataLista, popup_Confirmacao_Backup
from app import click_button, back_button

# Variável para testar inserções de dados
teste = False

# Verifica se o usuário está usando Windows
if platform.system() == "Windows":
    sistema_windows = True
    font_column = 18
    font_row = 16
    font_button = 35
    font_text = 35
    font_text_menu = 48
    font_title = 60

else:
    sistema_windows = False
    font_column = 20
    font_row = 19
    font_button = 55
    font_text = 40
    font_text_menu = 60
    font_title = 80


class CriarBackup(Screen):
    """
    Opção para fazer backup dos dados existentes, porém se faz necessário escolher as
    lista que deseja fazer o backup,
    poderá escolher entre: RD Marcas, Perfumaria, dermo ou todas ao mesmo tempo.
    """

    font_column = NumericProperty(font_column)
    font_row = NumericProperty(font_row)
    font_button = NumericProperty(font_button)
    font_text = NumericProperty(font_text)
    font_text_menu = NumericProperty(font_text_menu)
    font_title = NumericProperty(font_title)

    def pressButton(self):
        click_button.play()

    def pressBackButton(self):
        back_button.play()

    def __init__(self, **kw):
        super().__init__()
        self.button = None

    def fazerBackup_popup(self, button):
        """
        --> Função que mostra um Popup de confirmação antes de prosseguir com a exclusão da lista.
        """
        try:
            self.button = button

            content = BoxLayout(orientation='vertical', padding=10)
            label = Label(text='Confirma o Backup?')

            close_button = Button(text='Cancelar', size_hint=(1.0, 0.2))
            confirm_button = Button(text='Confirmar', size_hint=(1.0, 0.2))

            content.add_widget(label)
            content.add_widget(close_button)
            content.add_widget(confirm_button)

            popup = Popup(title='Aviso', content=content, size_hint=(0.5, 0.4))

            close_button.bind(on_release=popup.dismiss)

            if self.button == 'RD MARCAS' or self.button == 'PERFUMARIA' or self.button == 'DERMO':
                confirm_button.bind(on_release=lambda btn: self.realizarBackup())
            else:
                confirm_button.bind(on_release=lambda btn: self.realizarBackup_All())

            confirm_button.bind(on_release=popup.dismiss)
            popup.open()

        except Exception as error:
            print(error)
            popupError()

    def realizarBackup(self):
        try:
            # Pega a data formatada no dia atual
            hora = datetime.now()
            date = datetime.now()
            date = datetime.date(date)
            datahoje = date.strftime("%d-%m-%Y")
            horahoje = hora.strftime("%H;%M;%S")

            if self.button == 'RD MARCAS':
                nomeArquivoRD = f"BackupRDMARCAS-{datahoje}-{horahoje}"
                df_lista_RDMarcas = pd.read_excel(db_rdmarcas)
                df_lista_RDMarcas = formataLista(df_lista_RDMarcas, self.button)
                df_lista_RDMarcas.to_excel(f'app/data/backup/RDMarcas/{nomeArquivoRD}.xlsx', index=False)

            elif self.button == 'PERFUMARIA':
                nomeArquivoPERFUMARIA = f"BackupPERFUMARIA-{datahoje}-{horahoje}"
                df_lista_Perfumaria = pd.read_excel(db_perfumaria)
                df_lista_Perfumaria = formataLista(df_lista_Perfumaria, self.button)
                df_lista_Perfumaria.to_excel(f'app/data/backup/Perfumaria/{nomeArquivoPERFUMARIA}.xlsx', index=False)

            elif self.button == 'DERMO':
                nomeArquivoDERMO = f"BackupDERMO-{datahoje}-{horahoje}"
                df_lista_Dermo = pd.read_excel(db_dermo)
                df_lista_Dermo = formataLista(df_lista_Dermo, self.button)
                df_lista_Dermo.to_excel(f'app/data/backup/dermo/{nomeArquivoDERMO}.xlsx', index=False)

            popup_Confirmacao_Backup()

        except Exception as error:
            print(error)
            popupError()

    @staticmethod
    def realizarBackup_All():
        try:
            # Pega a data formatada no dia atual
            hora = datetime.now()
            date = datetime.now()
            date = datetime.date(date)
            datahoje = date.strftime("%d-%m-%Y")
            horahoje = hora.strftime("%H;%M;%S")

            nomeArquivoRD = f"BackupRDMARCAS-{datahoje}-{horahoje}"
            nomeArquivoPERFUMARIA = f"BackupPERFUMARIA-{datahoje}-{horahoje}"
            nomeArquivoDERMO = f"BackupDERMO-{datahoje}-{horahoje}"

            df_lista_RDMarcas = pd.read_excel(db_rdmarcas)
            df_lista_Perfumaria = pd.read_excel(db_perfumaria)
            df_lista_Dermo = pd.read_excel(db_dermo)

            df_lista_RDMarcas = formataLista(df_lista_RDMarcas, button='RD MARCAS')
            df_lista_Perfumaria = formataLista(df_lista_Perfumaria, button='PERFUMARIA')
            df_lista_Dermo = formataLista(df_lista_Dermo, button='DERMO')

            df_lista_RDMarcas.to_excel(f'app/data/backup/RDMarcas/{nomeArquivoRD}.xlsx', index=False)
            df_lista_Perfumaria.to_excel(f'app/data/backup/Perfumaria/{nomeArquivoPERFUMARIA}.xlsx', index=False)
            df_lista_Dermo.to_excel(f'app/data/backup/dermo/{nomeArquivoDERMO}.xlsx', index=False)
            popup_Confirmacao_Backup()

        except Exception as error:
            print(error)
            popupError()
