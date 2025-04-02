from tkinter import END, BooleanVar, Tk, Variable, ttk, Listbox, filedialog, messagebox

import win32com.client as win32
import win32print
import os

class ImprimirArquivosTab:
    def __init__(self, root: Tk, master = None):
        self.root = root
        self.master = master
        
        # Variables
        self.available_printers = win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
        self.file_list: list[dict[str, str]] = []
        self.var_file_list = Variable(value=[item['nome'] for item in self.file_list])
        
        self.var_laudo1 = BooleanVar(value=True)
        self.var_laudo2 = BooleanVar(value=False)
        self.var_laudo3 = BooleanVar(value=False)
        
        self.var_entrevista = BooleanVar(value=True)
        self.var_atestado_coletivo = BooleanVar(value=True)
        
        self.var_atestado_rt1 = BooleanVar(value=False)
        self.var_atestado_rt2 = BooleanVar(value=False)
        
        self.var_resultado_palo = BooleanVar(value=True)
        
        self.var_processo_de_avaliacao = BooleanVar(value=True)
        
        # Widgets
        self.frame = ttk.Frame(self.master)
        
        self.list_files = Listbox(self.frame, listvariable=self.var_file_list)
        
        self.combo_impressora = ttk.Combobox(self.frame, values=self.get_printer_name(), state='readonly')
        
        # Configures
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(99, weight=1)
        # self.frame.columnconfigure(1, weight=1)
        # self.frame.columnconfigure(2, weight=1)
    
    def render(self):
        
        # Declaração
        select_button = ttk.Button(self.frame, text='Selecionar Arquivos', command=self.selecionar_arquivos_btn)

        scroll_bar_list = ttk.Scrollbar(self.frame)
        

        ## Options
        label_options = ttk.Label(self.frame, text='Selecione as Pastas para imprimir')
        
        sel_frame = ttk.Frame(self.frame)
        
        sel_laudo1 = ttk.Checkbutton(sel_frame, variable=self.var_laudo1, text='Laudo 1')
        sel_laudo2 = ttk.Checkbutton(sel_frame, variable=self.var_laudo2, text='Laudo 2')
        sel_laudo3 = ttk.Checkbutton(sel_frame, variable=self.var_laudo3, text='Laudo 3')
        
        sel_entrevista = ttk.Checkbutton(sel_frame, variable=self.var_entrevista, text='Entrevista')
        sel_atestado_coletivo = ttk.Checkbutton(sel_frame, variable=self.var_atestado_coletivo, text='Atestado Coletivo')
        
        sel_atestado_rt1 = ttk.Checkbutton(sel_frame, variable=self.var_atestado_rt1, text='Atest. Reteste 1')
        sel_atestado_rt2 = ttk.Checkbutton(sel_frame, variable=self.var_atestado_rt2, text='Atest. Reteste 2')
        
        sel_resultado_palo = ttk.Checkbutton(sel_frame, variable=self.var_resultado_palo, text='Resultado do Palográfico')
        
        sel_processo_de_avaliacao = ttk.Checkbutton(sel_frame, variable=self.var_processo_de_avaliacao, text='Processo de Avaliação Psicológica')
        
        ## Butão
        btn_imprimir = ttk.Button(self.frame, text='Imprimir Planilhas', command=self.imprimir_planilhas)

        # Packs
        select_button.grid(row=0, column=0, pady=5, padx=5, sticky='ew', columnspan=4)
        self.combo_impressora.grid(row=1, column=0, padx=5, pady=5, sticky='ew', columnspan=4)
        
        self.list_files.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        scroll_bar_list.grid(row=2, column=3, sticky='nse')
        
        ttk.Separator(self.frame, orient='horizontal').grid(row=3, sticky='ew', columnspan=4, padx=5, pady=5)
        
        label_options.grid(row=4, column=0, padx=5, pady=5, columnspan=4)
        
        sel_frame.grid(sticky='nsew', padx=5, pady=5)
        
        sel_laudo1.grid(row=0, column=0, sticky='w')
        sel_laudo2.grid(row=1, column=0, sticky='w')
        sel_laudo3.grid(row=2, column=0, sticky='w')
        sel_entrevista.grid(row=3, column=0, sticky='w')
        
        sel_atestado_coletivo.grid(row=0, column=1, sticky='w')
        sel_atestado_rt1.grid(row=1, column=1, sticky='w')
        sel_atestado_rt2.grid(row=2, column=1, sticky='w')
        sel_resultado_palo.grid(row=3, column=1, sticky='w')
        
        sel_processo_de_avaliacao.grid(row=5, column=0, columnspan=2, sticky='w')
        
        btn_imprimir.grid(row=99, columnspan=4, sticky='nsew', padx=10, pady=10)
        
        # Configures
        self.list_files.config(yscrollcommand=scroll_bar_list.set)
        scroll_bar_list.config(command=self.list_files.yview)
        
        sel_frame.columnconfigure(0, weight=1)
        sel_frame.columnconfigure(1, weight=1)
        
        self.combo_impressora.set(win32print.GetDefaultPrinter())
        
    def get_printer_name(self):
        return [printer[2] for printer in self.available_printers] if self.available_printers else []
        
    def selecionar_arquivos_btn(self):
        
        self.file_list.clear()
        
        files = filedialog.askopenfilenames(
            title="Selecione todas as planilhas", 
            filetypes=[
                ('Arquivos Excel Macro', '*.xlsm')
            ])
        
        for item in files:
            self.file_list.append({
                'nome': os.path.basename(item),
                'path': item
            })
            print(item)
            
        self.var_file_list.set([new_item['nome'] for new_item in self.file_list])
        
    def imprimir_planilhas(self):
        print('Imprimindo planilhas...')
        print(f'Usando impressora: {self.combo_impressora.get()}')
        
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        
        # Olha os dados
        for item in self.file_list:
            
            # Abre o arquivo excel
            wb = excel.Workbooks.Open(item['path'])
            
            # Verifica o que foi marcado
            ## Processo de Avaliação
            if self.var_processo_de_avaliacao.get():
                proc = wb.Sheets('Proc')
                proc.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
            
            ## Laudos
            if self.var_laudo1.get():
                laudo1 = wb.Sheets('1Laudo')
                laudo1.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
            if self.var_laudo2.get():
                laudo2 = wb.Sheets('2Laudos')
                laudo2.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
            if self.var_laudo3.get():
                laudo3 = wb.Sheets('3Laudos')
                laudo3.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
                
            ## Atestados
            if self.var_atestado_coletivo.get():
                atest1 = wb.Sheets('AtestColet')
                atest1.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
            if self.var_atestado_rt1.get():
                atest2 = wb.Sheets('AtestRet1')
                atest2.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
            if self.var_atestado_rt2.get():
                atest3 = wb.Sheets('AtestRet2')
                atest3.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
            
            ## Entrevista
            if self.var_entrevista.get():
                entrv = wb.Sheets('Entrev')
                entrv.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )

            ## Resultado Palográfico
            if self.var_resultado_palo.get():
                palo = wb.Sheets('ResPalo')
                palo.PrintOut(
                    ActivePrinter=self.combo_impressora.get()
                )
            
            # Fecha a pasta
            wb.Close(
                SaveChanges=False
            )

        excel.Quit()
        messagebox.showinfo(title='Sucesso', message='Planilhas enviadas para impressão!')
    
    def get_frame(self) -> ttk.Frame:

        self.render()
        
        return self.frame