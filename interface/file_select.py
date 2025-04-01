from tkinter import filedialog
from tkinter import ttk, StringVar

class FileSelect:
    def __init__(self, frame: ttk.Frame):

        # Variáveis Globais
        self.frame = frame
        self.file_path: str = ''
        self.file_path_entry = StringVar()

    
    def render(self):
        # Elementos
        entry_label = ttk.Label(
            self.frame,
            text='Selecione a Lista de Candidatos',
            anchor='center'
        )

        path_entry = ttk.Entry(
            self.frame,
            textvariable=self.file_path_entry,
            state='disabled'
        )

        button_select_file = ttk.Button(
            self.frame, 
            command=self.open_file,
            text='📁'
        )

        # Packs
        entry_label.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky='ew',
            padx=5,
            pady=5
        )

        path_entry.grid(
            row=1,
            column=0,
            padx=5,
            pady=5,
            sticky='ew'
        )

        button_select_file.grid(
            row=1,
            column=1,
            padx=5,
            pady=5
        )

        # Configs
        self.frame.rowconfigure(0, weight=1)
        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(1, weight=0)

    def get_file_path_entry(self):       
        return self.file_path_entry.get()
    
    def get_file_path(self):
        return self.file_path
    
    def open_file(self):
        file = filedialog.askopenfilename(
            title='Selecione o arquivo PDF',
            filetypes=[
                ('Arquivos PDF', '*.pdf')
            ]
        )

        if file:
            self.file_path = file
            self.file_path_entry.set(file)