from tkinter import ttk, Menu, Tk

from interface.windows.config_window import ConfigWindow
from interface.windows.config_window_plan import ConfigWindowPlan

class ConfigMenu:

    def __init__(self, parent: Menu):
        self.parent = parent

    def render(self) -> Menu:
        menu = Menu(self.parent, tearoff=False)
        
        menu.add_command(label='Configurações do Gerador', command=self.openGenConfigFrame, state='disabled')
        menu.add_command(label='Configuração da Planilha', command=self.openPlanConfigFrame)

        return menu

    def openGenConfigFrame(self):
        print("Abrindo configurações do Gerador")
        
        window = ConfigWindow(title='Configurações do Gerador')
        window.show()

    def openPlanConfigFrame(self):
        print("Abrindo configurações da Planilha")
        
        ConfigWindowPlan(title='Configurações da Planilha').show()