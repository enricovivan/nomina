from tkinter import ttk, filedialog, messagebox, Tk

from modules.load_config import LoadConfig
from modules.pdf_data import PdfData
from modules.get_fist_file_from_folder import GetFirstFileFromFolder

from interface.file_select import FileSelect
from interface.clinic_select import ClinicSelect
from interface.psicho_select import PsychoSelect
from interface.date_select import DateSelect

import os
import shutil

import win32com.client

class GenButton:
    def __init__(
        self,
        frame: ttk.Frame, 
        file_select_instance: FileSelect, 
        radio_clinic_instance: ClinicSelect, 
        radio_psycho_instance: PsychoSelect,
        date_select_instance: DateSelect 
    ):
        self.frame = frame
        self.file_select = file_select_instance
        self.clinic_select = radio_clinic_instance
        self.psy_select = radio_psycho_instance
        self.date_select = date_select_instance

        self.plan_model_path = GetFirstFileFromFolder('./planilha', 'xlsm').getFilePath()

        self.gen_button: ttk.Button

        if self.plan_model_path is None:
            messagebox.showerror(
                title='Erro',
                message='Não foi encontrado nenhuma planilha modelo com extensão ".xlsm" dentro da pasta "planilha"\n\nPor favor, insira um arquivo modelo na pasta e abra o programa.'
            )
            raise Exception('Arquivo xlsm não encontrado na pasta "planilha"')


    def render(self):

        # Elements
        self.gen_button = ttk.Button(
            self.frame,
            text='Gerar Planilhas',
            command=self.gen_files,
            state='disabled'
        )

        # Packs
        self.gen_button.grid(
            row=99,
            column=0,
            padx=5,
            pady=5,
            sticky='nsew',
            columnspan=2
        )

        # Configs
        self.frame.rowconfigure(99, weight=1)       

        if self.file_select.get_file_path_entry() is not None or not '':
            self.gen_button.configure(
                state='normal'
            )

    def gen_files(self, show_message = True):

        print(f'Arquivo Selecionado: {self.file_select.get_file_path_entry()}')
        print(f'Data selecionada: {self.date_select.get_day()}')
        
        # Load config.json
        self.config_handler = LoadConfig()
        self.config_data = self.config_handler.load()

        # Verifica se a lista de candidatos foi selecionada
        if self.file_select.get_file_path_entry() == '' or None:
            self.gen_button.config(text='Gerar Planilhas', state='normal')
            messagebox.showerror(title='Erro', message='Por favor, selecione uma lista de candidatos!')
            return

        self.gen_button.config(text='Gerando...', state='disabled')

        save_folder = filedialog.askdirectory(
            title='Selecione a Pasta para salvar as planilhas'
        )

        # Caso não seja selecionado nenhum caminho de destino, retorna o botão ao estado original
        if save_folder == '' or None:
            self.gen_button.config(text='Gerar Planilhas', state='normal')
            return

        # Pega os dados do pdf
        doc_data = PdfData(self.file_select.get_file_path_entry())
        print(doc_data.get_pdf_data())

        print(f'Gerando planilhas no caminho: {save_folder}')

        # Verifica se o caminho de destino existe
        os.makedirs(save_folder, exist_ok=True)

        # Conecta o Excel na porta COM
        excel = win32com.client.Dispatch("Excel.Application")


        # Duplica os arquivos da planilha, cada um com seu nome
        for i in range(0, len(doc_data.get_pdf_data())):
            new_filename = f'{doc_data.get_row(i)['Nome']} {doc_data.get_row(i)['Processo']}.xlsm'
            new_filepath = os.path.join(save_folder, new_filename)

            print(new_filepath)

            shutil.copy(self.plan_model_path, new_filepath)

            workbook = excel.Workbooks.Open(new_filepath)

            # Altera Nome e Processo
            sheet = workbook.Sheets(self.config_data['planilha']['entrada']['posicao'])

            # sheet.Cells(6, 6).Value = doc_data.get_row(i)['Nome']
            # sheet.Cells(6, 24).Value = doc_data.get_row(i)['Processo']
            sheet.Range(self.config_data['planilha']['entrada']['celula_nome']).Value = doc_data.get_row(i)['Nome']
            sheet.Range(self.config_data['planilha']['entrada']['celula_processo']).Value = doc_data.get_row(i)['Processo']

            # Altera Clinica
            cadastros = workbook.Sheets(self.config_data['planilha']['cadastro']['posicao'])

            # Zera os valores de ambos as seleções
            # cadastros.Cells(9, 2).Value = ''
            # cadastros.Cells(11, 2).Value = ''
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_dmtran']).Value = ''
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_transitar']).Value = ''

            # Coloca o valor informado pela instância do ClinicSelect
            if self.clinic_select.get_selected() == 'transitar':
                # cadastros.Cells(11, 2).Value = 'x'
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_transitar']).Value = 'x'
            else:
                # cadastros.Cells(9, 2).Value = 'x'
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_dmtran']).Value = 'x'

            # Altera Psicóloga
            # Zera valores das psicologas
            # cadastros.Cells(23, 2).Value = ''
            # cadastros.Cells(25, 2).Value = ''
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_meire']).Value = ''
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_onilce']).Value = ''

            # Coloca o valor do radio
            if self.psy_select.get_selected() == 'meire':
                # cadastros.Cells(23, 2).Value = 'x'
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_meire']).Value = 'x'
            else:
                # cadastros.Cells(25, 2).Value = 'x'
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_onilce']).Value = 'x'

            # Altera Data
            # sheet.Cells(43, 9).Value = self.date_select.get_day()
            sheet.Range(self.config_data['planilha']['entrada']['celula_dia']).Value = self.date_select.get_day()

            combo_box = sheet.Shapes(self.config_data['planilha']['entrada']['nome_combobox_mes']).OLEFormat.Object
            combo_box.Value = self.date_select.get_month()

            # sheet.Cells(43, 18).Value = self.date_select.get_year()
            sheet.Range(self.config_data['planilha']['entrada']['celula_ano']).Value = self.date_select.get_year()

            # Salva e fecha
            workbook.Save()
            workbook.Close()


        excel.Quit()
        self.gen_button.config(text='Gerar Planilhas', state='normal')

        if show_message:
            messagebox.showinfo(
                title='Sucesso',
                message='Planilhas geradas com sucesso!'
            )
        
