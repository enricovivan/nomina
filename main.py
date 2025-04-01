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
    frame = ttk.Frame(root, padding=10)
    frame.grid(sticky='nsew')

    # Icone
    if getattr(sys, 'frozen', False):
        # Executável
        icon_path = os.path.join(sys._MEIPASS, 'Nomina.ico')
    else:
        # Modo desenvolvimento
        icon_path = os.path.join(os.path.dirname(__file__), 'common/Nomina.ico')

    # Configurações da Janela
    root.title('Nomina v1.3.0')
    root.iconbitmap(icon_path)
    # root.geometry("400x150")
    root.maxsize(550, 500)
    root.minsize(225, 100)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # Widgets da Janela
    menu_master = Menu(root)

    file_select = FileSelect(frame)
    clinic_radios = ClinicSelect(frame)
    psy_radios = PsychoSelect(frame)
    date_select = DateSelect(frame)

    gen_button = GenButton(frame, file_select, clinic_radios, psy_radios, date_select)
    odt_gen_button = GenOdtButton(frame, file_select, clinic_radios, psy_radios, date_select)
    plan_odt_button = GenOdtXlsmButton(root, frame, file_select, clinic_radios, psy_radios, date_select, gen_button, odt_gen_button)

    # Renders
    file_select.render()
    clinic_radios.render()
    psy_radios.render()
    date_select.render()

    ttk.Separator(frame, orient='horizontal').grid(row=98, sticky='ew', columnspan=2, padx=5, pady=5)

    gen_button.render()
    odt_gen_button.render()
    plan_odt_button.render()

    # Configs dos menus
    menu_master.add_cascade(label='Configurações', menu=ConfigMenu(menu_master).render())
    
    # Configs da root
    root.config(menu=menu_master)
    root.mainloop()

if __name__ == '__main__':
    main()