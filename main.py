import sys
import os

from tkinter import *
from tkinter import ttk

from interface.file_select import FileSelect
from interface.gen_button import GenButton
from interface.clinic_select import ClinicSelect
from interface.psicho_select import PsychoSelect
from interface.date_select import DateSelect
from interface.gen_odt_button import GenOdtButton
from interface.gen_odt_xlsm_button import GenOdtXlsmButton

from interface.menu.config_menu import ConfigMenu
from interface.tabs.gerar_arquivos import GerarArquivosTab
from interface.tabs.imprimir_arquivos import ImprimirArquivosTab
from interface.tabs.transformar_pdf import PdfizarArquivosTab

def main():

    # Configurações do windows para não ficar embaçado em 125% de zoom
    # Esse bloco é específico para Windows
    if sys.platform == "win32":
        try:
        # Para Windows 8.1 ou superior
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)  # ou 2 para Per Monitor DPI awareness pode ser uma opção a depender do cenário
        except Exception:
            try:
            # Em versões mais antigas do Windows pode ser necessária a chamada abaixo:
                from ctypes import windll
                windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    # TkInter
    root = Tk()

    # Icone
    if getattr(sys, 'frozen', False):
        # Executável
        icon_path = os.path.join(sys._MEIPASS, 'Nomina.ico')
    else:
        # Modo desenvolvimento
        icon_path = os.path.join(os.path.dirname(__file__), 'common/Nomina.ico')

    # Configurações da Janela
    root.title('Nomina v1.4.0')
    root.iconbitmap(icon_path)
    # root.geometry("400x150")
    root.maxsize(550, 550)
    root.minsize(225, 100)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # Widgets da Janela
    menu_master = Menu(root)
    
    # Tabs
    tab_control = ttk.Notebook(root)
    tab_control.grid(row=0, column=0, sticky='nsew')
    
    tab1 = GerarArquivosTab(root=root, master=tab_control).get_frame()
    tab2 = ImprimirArquivosTab(root, tab_control).get_frame()
    tab3 = PdfizarArquivosTab(root, tab_control).get_frame()
    
    tab_control.add(tab1, text="Gerar")
    tab_control.add(tab2, text="Imprimir")
    tab_control.add(tab3, text="Transformar PDF")

    # Configs dos menus
    menu_master.add_cascade(label='Configurações', menu=ConfigMenu(menu_master).render())
    
    # Configs da root
    root.config(menu=menu_master)
    root.mainloop()

if __name__ == '__main__':
    main()