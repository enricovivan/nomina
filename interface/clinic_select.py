from tkinter import ttk, StringVar

class ClinicSelect:
    def __init__(self, frame: ttk.Frame):
        self.frame = frame

        self.radio_value = StringVar(value='transitar')

    def render(self):

        clinica_label = ttk.Label(
            self.frame, 
            text='Selecione a Clínica', 
            anchor='center'
        )

        radio_transitar = ttk.Radiobutton(
            self.frame, 
            text='Transitar', 
            variable=self.radio_value,
            value='transitar',
        )
        radio_dmtran = ttk.Radiobutton(
            self.frame, 
            text='DMTRAN',
            variable=self.radio_value,
            value='dmtran'
        )

        # Packs
        clinica_label.grid(
            row=2,
            column=0,
            columnspan=2,
            sticky='ew',
            padx=5,
            pady=5
        )
        radio_transitar.grid(
            row=3,
            column=0,
            sticky='w',
            padx=5,
            pady=5
        )
        radio_dmtran.grid(
            row=3,
            column=1,
            sticky='e',
            padx=5,
            pady=5
        )

    def get_selected(self) -> str:
        return self.radio_value.get()