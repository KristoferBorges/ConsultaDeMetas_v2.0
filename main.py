import platform
from app.functions.consultar_listas import ConsultaDeListas
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager
from kivy.core.window import Window
from app.telas.telas import MenuPrincipal
from app.telas.telas import NovosRegistros
from app.produtos.rdmarcas import RegistrosRDMarcas
from app.produtos.rdmarcas import LimparRD
from app.produtos.rdmarcas import ConsultaRDMarcas
from app.produtos.perfumaria import RegistrosPerfumaria
from app.produtos.perfumaria import LimparPerfumaria
from app.produtos.perfumaria import ConsultaPerfumaria
from app.produtos.dermo import RegistrosDermo
from app.produtos.dermo import LimparDermo
from app.produtos.dermo import ConsultaDermo
from app.functions.limpar_lista import LimparDados, LimparTodasAsListas
from app.functions.criar_backups import CriarBackup
from app.functions.fechar_programa import FecharPrograma

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

# Variável para testar inserções de dados
teste = False


MenuPrincipal()
NovosRegistros()
RegistrosRDMarcas()
RegistrosPerfumaria()
RegistrosDermo()
LimparDados()
LimparRD()
LimparPerfumaria()
LimparDermo()
LimparTodasAsListas()
ConsultaDeListas()
ConsultaRDMarcas()
ConsultaPerfumaria()
ConsultaDermo()
CriarBackup()
FecharPrograma()


class Tela(App):

    def build(self):
        if sistema_windows:
            Window.size = (1130, 810)
        self.title = 'ConsultaDeMetas_v2.0'
        adm = ScreenManager()
        return adm


if __name__ == '__main__':
    Tela().run()
