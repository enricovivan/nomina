from tkinter import ttk, Tk

from interface.clinic_select import ClinicSelect
from interface.date_select import DateSelect
from interface.file_select import FileSelect
from interface.gen_button import GenButton
from interface.gen_odt_button import GenOdtButton
from interface.gen_odt_xlsm_button import GenOdtXlsmButton
from interface.pdf_format_select import PdfFormatSelect
from interface.psicho_select import PsychoSelect

class GerarArquivosTab:
    
    def __init__(self, root: Tk, master = None):        
        self.root = root
        self.master = master
        
        self.frame = ttk.Frame(self.master)
        
    def render(self):
        pdf_select = PdfFormatSelect(self.frame)
        file_select = FileSelect(self.frame)
        clinic_radios = ClinicSelect(self.frame)
        psy_radios = PsychoSelect(self.frame)
        date_select = DateSelect(self.frame)

        gen_button = GenButton(self.frame, pdf_select, file_select, clinic_radios, psy_radios, date_select)
        odt_gen_button = GenOdtButton(self.frame, file_select, clinic_radios, psy_radios, date_select)
        plan_odt_button = GenOdtXlsmButton(self.root, self.frame, file_select, clinic_radios, psy_radios, date_select, gen_button, odt_gen_button)

        # Renders
        pdf_select.render()
        file_select.render()
        clinic_radios.render()
        psy_radios.render()
        date_select.render()

        ttk.Separator(self.frame, orient='horizontal').grid(row=98, sticky='ew', columnspan=2, padx=5, pady=5)

        gen_button.render()
        odt_gen_button.render()
        plan_odt_button.render()
        
    def get_frame(self) -> ttk.Frame:
        
        self.render()
        
        return self.frame