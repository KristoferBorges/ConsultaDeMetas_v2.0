from app.support.setup import Setup
from app import *

class Tela(App):

    def build(self):
        self.title = 'Consulta De Metas'
        Setup()
        adm = ScreenManager()
        return adm

if __name__ == '__main__':
    Tela().run()
