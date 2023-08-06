# app.py

import platform
from kivy.core.window import Window
from app.screens.telas import MenuPrincipal, NovosRegistros
from app.actions.rdmarcas import RegistrosRDMarcas, LimparRD, ConsultaRDMarcas
from app.actions.perfumaria import RegistrosPerfumaria, LimparPerfumaria, ConsultaPerfumaria
from app.actions.dermo import RegistrosDermo, LimparDermo, ConsultaDermo
from app.functions.limpar_lista import LimparDados, LimparTodasAsListas
from app.functions.consultar_listas import ConsultaDeListas
from app.functions.criar_backups import CriarBackup
from app.functions.fechar_programa import FecharPrograma


class Setup:
    def __init__(self):
        # Verifica se o usuário está usando Windows
        if platform.system() == "Windows":
            sistema_windows = True
            font_column = 18
            font_row = 16
            font_button = 35
            font_text = 35
            font_text_menu = 48
            font_title = 60
            Window.size = (1130, 810)

        else:
            sistema_windows = False
            font_column = 20
            font_row = 19
            font_button = 55
            font_text = 40
            font_text_menu = 60
            font_title = 80
        pass

    def run(self):
        # Coloque aqui o código que deseja executar na função run.
        menu_principal = MenuPrincipal()
        novos_registros = NovosRegistros()
        registros_rdmarcas = RegistrosRDMarcas()
        limpar_rd = LimparRD()
        consulta_rdmarcas = ConsultaRDMarcas()
        registros_perfumaria = RegistrosPerfumaria()
        limpar_perfumaria = LimparPerfumaria()
        consulta_perfumaria = ConsultaPerfumaria()
        registros_dermo = RegistrosDermo()
        limpar_dermo = LimparDermo()
        consulta_dermo = ConsultaDermo()
        consultar_listas = ConsultaDeListas()
        limpar_dados = LimparDados()
        limpar_todas_as_listas = LimparTodasAsListas()
        criar_backup = CriarBackup()
        fechar_programa = FecharPrograma()

        # Coloque aqui qualquer código adicional que deseja executar na função run.
