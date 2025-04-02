import os
from tkinter import BooleanVar, Listbox, Tk, Variable, filedialog, ttk, messagebox

import win32com.client as win32
import pythoncom
import time

class PdfizarArquivosTab:
    def __init__(self, root: Tk, master = None):
        self.root = root
        self.master = master
        
        # Variáveis
        self.exporting_ws: list[str] = []
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
        
        self.btn_ger_pdf = ttk.Button(self.frame, text="Gerar PDFs", command=self.gerar_pdfs)
        
        
        # Configures
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(99, weight=1)
        
    def render(self):
        
        # Declarações
        select_button = ttk.Button(self.frame, text='Selecionar Arquivos', command=self.selecionar_arquivos_btn)
        scroll_bar_list = ttk.Scrollbar(self.frame)
        
        ## Options
        label_options = ttk.Label(self.frame, text='Selecione as Pastas para transformar')
        
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
        
        # Packs
        select_button.grid(row=0, column=0, pady=5, padx=5, sticky='ew', columnspan=4)
        
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
        
        self.btn_ger_pdf.grid(row=99, column=0, columnspan=4, sticky='nsew', padx=10, pady=10)
        
        # Configures
        self.list_files.config(yscrollcommand=scroll_bar_list.set)
        scroll_bar_list.config(command=self.list_files.yview)

    def gerar_pdfs(self):
        print("Gerando os arquivos PDF...")
        
        self.btn_ger_pdf.config(
            state='disabled',
            text='Gerando...'
        )
        
        ger_path = filedialog.askdirectory(
            title='Selecione um Local Para salvar os PDFs gerados...'
        )
        
        if not ger_path:
            self.btn_ger_pdf.config(
                state='normal',
                text='Gerar PDFs'
            )
            return
        
        print(f"Caminho selecionado: {ger_path}")
        
        pythoncom.CoInitialize()
        excel = None
        try:
            # Abre o excel
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
        
            # Cria a pasta de salvamento se não existir
            save_folder_geral = os.path.join(ger_path, 'PDFs')
            os.makedirs(save_folder_geral, exist_ok=True)
            
            # Pega os dados das planilhas selecionadas
            for item in self.file_list:
                           
                # Cria a pasta para o camarada
                candidato_folder = os.path.join(save_folder_geral, item['nome'][:-17])
                os.makedirs(candidato_folder, exist_ok=True)
                
                wb = None
                
                try:           
                    # Abre o workbook dele
                    wb = excel.Workbooks.Open(item['path'])
                    
                    # Verifica o que foi marcado
                    ## Processo de Avaliação
                    if self.var_processo_de_avaliacao.get():
                        self.export_sheets_as_pdf(wb, 'Proc', candidato_folder, f'Processo de Avaliação Psicológica - {item["nome"][:-17]}')
                    
                    ## Laudos
                    if self.var_laudo1.get():
                        self.export_sheets_as_pdf(wb, '1Laudo', candidato_folder, f'Laudo 1 - {item["nome"][:-17]}')
                    if self.var_laudo2.get():
                        self.export_sheets_as_pdf(wb, '2Laudos', candidato_folder, f'Laudo 2 - {item["nome"][:-17]}')
                    if self.var_laudo3.get():
                        self.export_sheets_as_pdf(wb, '3Laudos', candidato_folder, f'Laudo 3 - {item["nome"][:-17]}')
                        
                    ## Atestados
                    if self.var_atestado_coletivo.get():
                        self.export_sheets_as_pdf(wb, 'AtestColet', candidato_folder, f'Atestado Coletivo - {item["nome"][:-17]}')
                    if self.var_atestado_rt1.get():
                        self.export_sheets_as_pdf(wb, 'AtestRet1', candidato_folder, f'Atestado Reteste 1 - {item["nome"][:-17]}')
                    if self.var_atestado_rt2.get():
                        self.export_sheets_as_pdf(wb, 'AtestRet2', candidato_folder, f'Atestado Reteste 2 - {item["nome"][:-17]}')
                    
                    ## Entrevista
                    if self.var_entrevista.get():
                        self.export_sheets_as_pdf(wb, 'Entrev', candidato_folder, f'Entrevista - {item["nome"][:-17]}')

                    ## Resultado Palográfico
                    if self.var_resultado_palo.get():
                        self.export_sheets_as_pdf(wb, 'ResPalo', candidato_folder, f'Resultado Palográfico - {item["nome"][:-17]}')
                        
                    # Exporta as selecionadas
                    # print('Exportando planilhas: ')
                    # print(self.exporting_ws)
                    
                    
                except Exception as e:
                    print(f'Erro ao gerar pdfs: {e}')
                finally:
                    if wb:
                        wb.Close(SaveChanges=False)
                        
            messagebox.showinfo(message="PDFs Gerados com Sucesso.")
                    
        except Exception as e:
            print(f'Erro geral: {e}')
            messagebox.showerror(message="Erro ao gerar PDFs")
        finally:
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()

            self.btn_ger_pdf.config(
                state='normal',
                text='Gerar PDFs'
            )
            
    def export_sheets_as_pdf(self, wb, sheet_name, folder, filename):
        try:
            ws = wb.Worksheets(sheet_name)
            full_path = os.path.join(folder, f'{filename}.pdf')
            
            # ws.ExportAsFixedFormat(
            #     Type=0,
            #     Filename=full_path,
            #     IgnorePrintAreas=False
            # )
            ws.PrintOut(
                Copies=1,
                Preview=False,
                ActivePrinter="Microsoft Print to PDF",
                PrintToFile=True,
                PrToFileName=full_path,
                Collate=True
            )
            print(f"Arquivo Salvo: {full_path}")
            
        except Exception as e:
            print(f"Erro ao gerar o PDF: {e}")

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
    
    def get_frame(self) -> ttk.Frame:
        self.render()
        return self.frame