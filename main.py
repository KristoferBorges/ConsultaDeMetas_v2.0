from kivy.app import App
from kivy.uix.screenmanager import ScreenManager
from app.support.setup import Setup


class Tela(App):

    def build(self):
        self.title = 'Consulta De Metas'
        Setup()
        adm = ScreenManager()
        return adm


if __name__ == '__main__':
    Tela().run()
