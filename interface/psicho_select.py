from tkinter import ttk, StringVar

class PsychoSelect:
    def __init__(self, frame: ttk.Frame):
        self.frame = frame

        self.psycho_value = StringVar(value='meire')

    def render(self):
        
        label = ttk.Label(
            self.frame,
            text="Selecione a Psicóloga",
            anchor='center'
        )
        radio_meire = ttk.Radiobutton(
            self.frame, 
            text='Mire', 
            variable=self.psycho_value,
            value='meire',
        )
        radio_onilce = ttk.Radiobutton(
            self.frame, 
            text='Onilce', 
            variable=self.psycho_value,
            value='onilce',
        )

        # Packs
        label.grid(
            row=4,
            column=0,
            columnspan=2,
            sticky='ew',
            padx=5,
            pady=5
        )

        radio_meire.grid(
            row=5,
            column=0,
            sticky='w',
            padx=5,
            pady=5
        )

        radio_onilce.grid(
            row=5,
            column=1,
            sticky='e',
            padx=5,
            pady=5
        )

    def get_selected(self):
        return self.psycho_value.get()