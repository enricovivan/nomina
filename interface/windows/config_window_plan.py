from tkinter import messagebox, ttk, Tk, Toplevel, StringVar

from modules.load_config import LoadConfig

import json

class ConfigWindowPlan:
    def __init__(self, title: str):
        self.title = title
        
        # Configs
        self.config_handler = LoadConfig()
        self.config_data = self.config_handler.load()
        
        # Entries
        self.pos_entrada_entry = StringVar(value=self.config_data['planilha']['entrada']['posicao'])
        
        self.cell_nome_entry = StringVar(value=self.config_data['planilha']['entrada']['celula_nome'])
        self.cell_processo_entry = StringVar(value=self.config_data['planilha']['entrada']['celula_processo'])
        self.cell_dia_entry = StringVar(value=self.config_data['planilha']['entrada']['celula_dia'])
        self.name_mes_entry = StringVar(value=self.config_data['planilha']['entrada']['nome_combobox_mes'])
        self.cell_ano_entry = StringVar(value=self.config_data['planilha']['entrada']['celula_ano'])
        
        self.pos_cadastro_entry = StringVar(value=self.config_data['planilha']['cadastro']['posicao'])
        
        self.cell_x_dmtran_entry = StringVar(value=self.config_data['planilha']['cadastro']['celula_x_dmtran'])
        self.cell_x_transitar_entry = StringVar(value=self.config_data['planilha']['cadastro']['celula_x_transitar'])
        self.cell_x_meire_entry = StringVar(value=self.config_data['planilha']['cadastro']['celula_x_meire'])
        self.cell_x_onilce_entry = StringVar(value=self.config_data['planilha']['cadastro']['celula_x_onilce'])
        
        # Widgets
        self.window = Toplevel()
        self.frame = ttk.Frame(self.window, padding=10)
        self.frame.grid(sticky='nsew')    
        
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
    def show(self):
        self.window.title(self.title)
        # self.window.geometry('300x200')

        self.render_elements()
        
    def render_elements(self):

        # Widgets
        label_title = ttk.Label(self.frame, text='Configurações da Planilha')
        label_celulas = ttk.Label(self.frame, text='Células')
        label_celulas2 = ttk.Label(self.frame, text='Células')
        
        label_posicoes_entrada = ttk.Label(self.frame, text='Pasta Entrada')
        
        label_posicao_entrada = ttk.Label(self.frame, text='Posição Entrada')
        entry_posicao_entrada = ttk.Entry(self.frame, textvariable=self.pos_entrada_entry)
        
        label_cell_nome = ttk.Label(self.frame, text='Nome do Candidato')
        entry_cell_nome = ttk.Entry(self.frame, textvariable=self.cell_nome_entry)
        
        label_cell_processo = ttk.Label(self.frame, text='Processo')
        entry_cell_processo  = ttk.Entry(self.frame, textvariable=self.cell_processo_entry)
        
        label_cell_dia = ttk.Label(self.frame, text='Célula dia')
        entry_cell_dia =  ttk.Entry(self.frame, textvariable=self.cell_dia_entry)
        
        label_name_mes_dia = ttk.Label(self.frame, text='Nome da Caixa do Mês')
        entry_name_mes_dia = ttk.Entry(self.frame, textvariable=self.name_mes_entry)

        label_cell_ano = ttk.Label(self.frame, text='Célula ano')
        entry_cell_ano = ttk.Entry(self.frame, textvariable=self.cell_ano_entry)
        
        # Divider
        
        label_posicoes_cadastro = ttk.Label(self.frame, text='Pasta Cadastro')
        
        label_posicao_cadastro = ttk.Label(self.frame, text='Posição Cadastro')
        entry_posicao_cadastro = ttk.Entry(self.frame, textvariable=self.pos_cadastro_entry)
        
        # Células
        label_cell_dmtran = ttk.Label(self.frame, text='"x" DMTRAN')
        entry_cell_dmtran = ttk.Entry(self.frame, textvariable=self.cell_x_dmtran_entry)
        
        label_cell_transitar = ttk.Label(self.frame, text='"x" Transitar')
        entry_cell_transitar = ttk.Entry(self.frame, textvariable=self.cell_x_transitar_entry)
        
        label_cell_meire = ttk.Label(self.frame, text='"x" Meire')
        entry_cell_meire = ttk.Entry(self.frame, textvariable=self.cell_x_meire_entry)
        
        label_cell_onilce = ttk.Label(self.frame, text='"x" Onilce')
        entry_cell_onilce = ttk.Entry(self.frame, textvariable=self.cell_x_onilce_entry)
        
        # Buttons
        default_config_button = ttk.Button(self.frame, text='Caregar Config. Padrão', command=self.load_default_config)
        close_button = ttk.Button(self.frame, text='Cancelar', command=self.close_button_click)
        save_button = ttk.Button(self.frame, text='Salvar', command=self.save_button_click)
        
        # Renders
        label_title.grid(
            row=0,
            column=0,
            columnspan=5
        )
        
        ttk.Separator(self.frame, orient='horizontal').grid(row=1, sticky='ew', columnspan=5, padx=5, pady=5)
        
        label_posicoes_entrada.grid(
            row=2,
            column=0,
            pady=5
        )
        
        label_posicao_entrada.grid(
            row=3,
            column=0,
            sticky='nw'
        )
        entry_posicao_entrada.grid(
            row=4,
            column=0,
            sticky='ew',
            columnspan=2,
            pady=5
        )
        
        label_celulas.grid(
            row=5,
            column=0,
            pady=10
        )
        
        label_cell_nome.grid(
            row=6,
            column=0,
            sticky='nw'
        )
        entry_cell_nome.grid(
            row=7,
            column=0,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_cell_processo.grid(
            row=8,
            column=0,
            sticky='nw'
        )
        entry_cell_processo.grid(
            row=9,
            column=0,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_cell_dia.grid(
            row=10,
            column=0,
            sticky='nw'
        )
        entry_cell_dia.grid(
            row=11,
            column=0,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_name_mes_dia.grid(
            row=12,
            column=0,
            sticky='nw'
        )
        entry_name_mes_dia.grid(
            row=13,
            column=0,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_cell_ano.grid(
            row=14,
            column=0,
            sticky='nw'
        )
        entry_cell_ano.grid(
            row=15,
            column=0,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        ttk.Separator(self.frame, orient='vertical').grid(row=1, column=2, rowspan=16, sticky='ns', padx=10, pady=10)
        
        label_posicoes_cadastro.grid(
            row=2,
            column=3,
            columnspan=2,
            pady=5
        )
        
        label_posicao_cadastro.grid(
            row=3,
            column=3,
            sticky='nw'
        )
        entry_posicao_cadastro.grid(
            row=4,
            column=3,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_celulas2.grid(
            row=5,
            column=3,
            pady=10
        )
        
        label_cell_dmtran.grid(
            row=6,
            column=3,
            sticky='nw'
        )
        entry_cell_dmtran.grid(
            row=7,
            column=3,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_cell_transitar.grid(
            row=8,
            column=3,
            sticky='nw'
        )
        entry_cell_transitar.grid(
            row=9,
            column=3,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_cell_meire.grid(
            row=10,
            column=3,
            sticky='nw'
        )
        entry_cell_meire.grid(
            row=11,
            column=3,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        label_cell_onilce.grid(
            row=12,
            column=3,
            sticky='nw'
        )
        entry_cell_onilce.grid(
            row=13,
            column=3,
            columnspan=2,
            sticky='ew',
            pady=5
        )
        
        ttk.Separator(self.frame, orient='horizontal').grid(row=16, sticky='ew', columnspan=5, padx=5, pady=5)
        
        # Buttons
        default_config_button.grid(
            row=17,
            column=0,
            columnspan=5,
            sticky='nsew'
        )
        
        close_button.grid(
            row=18,
            column=0,
            columnspan=2,
            sticky='nsew'
        )
        
        save_button.grid(
            row=18,
            column=3,
            columnspan=2,
            sticky='nsew'
        )
    
        # Config
        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(2, weight=1)
        self.frame.columnconfigure(3, weight=1)
        self.frame.columnconfigure(17, weight=1)
        
        self.frame.rowconfigure(17, weight=1)
        self.frame.rowconfigure(18, weight=1)
        
    def close_button_click(self):
        self.window.destroy()
    
    def save_button_click(self):
        
        # Verifica os valores
        if self.pos_entrada_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_nome_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_processo_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_dia_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.name_mes_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_ano_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.pos_cadastro_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_x_dmtran_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_x_transitar_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_x_meire_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        if self.cell_x_onilce_entry.get() == '':
            messagebox.showwarning(message='Insira pelo menos um valor...')
            return
        
        self.config_data['planilha']['entrada']['posicao'] = int(self.pos_entrada_entry.get())
        self.config_data['planilha']['entrada']['celula_nome'] = self.cell_nome_entry.get()
        self.config_data['planilha']['entrada']['celula_processo'] = self.cell_processo_entry.get()
        self.config_data['planilha']['entrada']['celula_dia'] = self.cell_dia_entry.get()
        self.config_data['planilha']['entrada']['nome_combobox_mes'] = self.name_mes_entry.get()
        self.config_data['planilha']['entrada']['celula_ano'] = self.cell_ano_entry.get()
        self.config_data['planilha']['cadastro']['posicao'] = int(self.pos_cadastro_entry.get())
        self.config_data['planilha']['cadastro']['celula_x_dmtran'] = self.cell_x_dmtran_entry.get()
        self.config_data['planilha']['cadastro']['celula_x_transitar'] = self.cell_x_transitar_entry.get()
        self.config_data['planilha']['cadastro']['celula_x_meire'] = self.cell_x_meire_entry.get()
        self.config_data['planilha']['cadastro']['celula_x_onilce'] = self.cell_x_onilce_entry.get()
        
        self.config_handler.save(self.config_data)

        messagebox.showinfo(title='Sucesso', message='Configurações Atualizadas!')
        self.window.destroy()
        
    def load_default_config(self):
        
        config = LoadConfig(filename='config.default.json')
        config_data = config.load()
        
        self.pos_entrada_entry.set(config_data['planilha']['entrada']['posicao'])
        
        self.cell_nome_entry.set(config_data['planilha']['entrada']['celula_nome'])
        self.cell_processo_entry.set(config_data['planilha']['entrada']['celula_processo'])
        self.cell_dia_entry.set(config_data['planilha']['entrada']['celula_dia'])
        self.name_mes_entry.set(config_data['planilha']['entrada']['nome_combobox_mes'])
        self.cell_ano_entry.set(config_data['planilha']['entrada']['celula_ano'])
        
        self.pos_cadastro_entry.set(config_data['planilha']['cadastro']['posicao'])
        
        self.cell_x_dmtran_entry.set(config_data['planilha']['cadastro']['celula_x_dmtran'])
        self.cell_x_transitar_entry.set(config_data['planilha']['cadastro']['celula_x_transitar'])
        self.cell_x_meire_entry.set(config_data['planilha']['cadastro']['celula_x_meire'])
        self.cell_x_onilce_entry.set(config_data['planilha']['cadastro']['celula_x_onilce'])