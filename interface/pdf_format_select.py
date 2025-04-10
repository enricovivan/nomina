from tkinter import StringVar, ttk


class PdfFormatSelect:
    def __init__(self, frame: ttk.Frame):
        self.frame = frame
        self.var_pdf = StringVar(value='normal')
    
    def render(self):
        
        label = ttk.Label(
            self.frame,
            text="Selecione o tipo do Documento PDF",
            anchor='center'
        )
        radio_normal = ttk.Radiobutton(
            self.frame, 
            text='Normal', 
            variable=self.var_pdf,
            value='normal'
        )
        radio_nome = ttk.Radiobutton(
            self.frame, 
            text='Com Nome', 
            variable=self.var_pdf,
            value='nome'
        )

        # Packs
        label.grid(
            row=2,
            column=0,
            columnspan=2,
            sticky='ew',
            padx=5,
            pady=5
        )

        radio_normal.grid(
            row=3,
            column=0,
            sticky='w',
            padx=5,
            pady=5
        )

        radio_nome.grid(
            row=3,
            column=1,
            sticky='e',
            padx=5,
            pady=5
        )
    
    def get_selected(self) -> str:
        return self.var_pdf.get()
    