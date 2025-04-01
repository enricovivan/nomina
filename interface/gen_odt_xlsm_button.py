from tkinter import ttk, filedialog, messagebox, WORD, Tk
from docxtpl import DocxTemplate

from modules.load_config import LoadConfig
from modules.pdf_data import PdfData
from modules.get_fist_file_from_folder import GetFirstFileFromFolder

from interface.file_select import FileSelect
from interface.clinic_select import ClinicSelect
from interface.psicho_select import PsychoSelect
from interface.date_select import DateSelect
from interface.gen_button import GenButton
from interface.gen_odt_button import GenOdtButton

from classes.clinica import Clinica
from classes.psicologo import Psicologo

import os
import shutil

import win32com.client

class GenOdtXlsmButton:
    
    def __init__(
        self,
        root: Tk,
        frame: ttk.Frame,
        file_select_instance: FileSelect, 
        radio_clinic_instance: ClinicSelect, 
        radio_psycho_instance: PsychoSelect,
        date_select_instance: DateSelect,
        gen_planilhas_button_instance: GenButton,
        gen_odts_button_instance: GenOdtButton
    ):
        self.root = root
        self.frame = frame
        self.frame = frame
        self.file_select = file_select_instance
        self.clinic_select = radio_clinic_instance
        self.psy_select = radio_psycho_instance
        self.date_select = date_select_instance
        self.gen_plan = gen_planilhas_button_instance
        self.gen_odt = gen_odts_button_instance

        self.odt_model_path = GetFirstFileFromFolder('./words', 'docx').getFilePath()
        self.plan_model_path = GetFirstFileFromFolder('./planilha', 'xlsm').getFilePath()

        self.button: ttk.Button

    def render(self):

        button_style = ttk.Style(self.root)

        custom_style_name = 'Custom.GenOdtXlsmButtonStyle'

        layout = button_style.layout('TButton')
        button_style.layout(custom_style_name, layout)

        button_style.configure(custom_style_name, anchor="center", padding=5, justify='center')
        
        self.button = ttk.Button(
            self.frame,
            text='Gerar Planilhas e\nProcessos de Avaliação Psicológica',
            command=self.gen_all_files,
            style=custom_style_name
        )

        # Packs
        self.button.grid(
            row=101,
            column=0,
            padx=5,
            pady=5,
            sticky='nsew',
            columnspan=2
        )

        # Configs
        self.frame.rowconfigure(101, weight=1)


    def gen_all_files(self):
        print('Gerando todos os arquivos...')
        
        # Load config.json
        self.config_handler = LoadConfig()
        self.config_data = self.config_handler.load()

        # Checa se um arquivo PDF foi selecionado
        if self.file_select.get_file_path_entry() == '' or None:
            self.button.config(text='Gerar Planilhas e\nProcessos de Avaliação Psicológica', state='normal')
            messagebox.showerror(title='Erro', message='Por favor, selecione uma lista de candidatos!')
            return

        self.button.config(text='Gerando Planilhas e\nProcessos...', state='disabled')

        save_folder = filedialog.askdirectory(
            title='Selecione a pasta para salvar as planilhas e os processos'
        )

        # Caso não seja selecionado nenhum caminho de destino, retorna o botão ao estado original
        if save_folder == '' or None:
            self.button.config(text='Gerar Planilhas e\nProcessos de Avaliação Psicológica', state='normal')
            # messagebox.showwarning(title='Atenção', message='Por favor, selecione uma pasta para gerar os arquivos')
            return
        
        print(f'Pasta selecionada: {save_folder}')

        # Cria a pasta de planilhas se não existir
        os.makedirs(save_folder, exist_ok=True)

        # Cria a pasta de processos se não existir
        save_folder_process = os.path.join(save_folder, 'Processos de Avaliação Psicológica')
        os.makedirs(save_folder_process, exist_ok=True)

        # Pega os dados do pdf
        doc_data = PdfData(self.file_select.get_file_path_entry())
        print(doc_data.get_pdf_data())

        ###################
        #       ODT       #
        ###################
        # Define o nome e endereça da clinica selecionada
        clinica: Clinica

        if self.clinic_select.get_selected() == 'dmtran':
            clinica = Clinica(
                nome='DMTRAN - Clínica de Avaliação Médica e Psicológica para o Trânsito',
                endereco='Rua Mato Grosso, 638 - Centro - CEP: 86870-000 Fone/ Fax (43) 3472-1590 - Ivaiporã - Pr.'
            )
        else :
            clinica = Clinica(
                nome='TRANSITAR - Clínica de Avaliação Médica e Psicológica para o Trânsito',
                endereco='Rua Placídio Miranda, 445 - Centro - CEP: 86870-000 Fone/ Fax (43) 3472-9962 - Ivaiporã - PR.'
            )

        # Define o psicologo responsável
        psicologo: Psicologo

        if self.psy_select.get_selected() == 'meire':
            psicologo = Psicologo(
                nome='Meire Regiane Lourenço Nunes',
                crp='08/09673'
            )
        else:
            psicologo = Psicologo(
                nome='Onilce Célia Pricinato Estrada',
                crp='08/1675'
            )

        # Pega a data selecionada e converte para texto comum
        data_selecionada = f'{self.date_select.get_day()}/{self.date_select.get_month()}/{self.date_select.get_year()}'

        # Lê o arquivo ODT
        doc = DocxTemplate(self.odt_model_path)

        context = {
            'clinica_nome': clinica.getNome(),
            'clinica_endereco': clinica.getEndereco(),
            'psicologo_nome': psicologo.getNome(),
            'psicologo_crp': psicologo.getCrp(),
            'data_exame': data_selecionada
        }

        ####################
        #     Planilha     #
        ####################
        # Conecta o Excel na porta COM
        excel = win32com.client.Dispatch("Excel.Application")

        # Loop para iterar a lista de candidatos e gerar os arquivos
        # começando pelos processos (pois é mais rápido)
        for i in range(0, len(doc_data.get_pdf_data())):
            
            ###########
            #   ODT   #
            ###########
            # Adiciona o Nome e o Processo do candidato no contexto
            context['candidato_nome'] = doc_data.get_row(i)['Nome']
            context['candidato_processo'] = doc_data.get_row(i)['Processo']

            doc.render(context)

            new_filepath = os.path.join(save_folder_process, f'{doc_data.get_row(i)['Nome']} {doc_data.get_row(i)['Processo']}.docx')

            doc.save(new_filepath)

        for i in range(0, len(doc_data.get_pdf_data())):

            ####################
            #     Planilha     #
            ####################
            new_filename = f'{doc_data.get_row(i)['Nome']} {doc_data.get_row(i)['Processo']}.xlsm'
            new_filepath = os.path.join(save_folder, new_filename)

            print(new_filepath)

            shutil.copy(self.plan_model_path, new_filepath)

            workbook = excel.Workbooks.Open(new_filepath)

            # Altera Nome e Processo
            sheet = workbook.Sheets(self.config_data['planilha']['entrada']['posicao'])

            sheet.Range(self.config_data['planilha']['entrada']['celula_nome']).Value = doc_data.get_row(i)['Nome']
            sheet.Range(self.config_data['planilha']['entrada']['celula_processo']).Value = doc_data.get_row(i)['Processo']

            # Altera Clinica
            cadastros = workbook.Sheets(self.config_data['planilha']['cadastro']['posicao'])

            # Zera os valores de ambos as seleções
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_dmtran']).Value = ''
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_transitar']).Value = ''

            # Coloca o valor informado pela instância do ClinicSelect
            if self.clinic_select.get_selected() == 'transitar':
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_transitar']).Value = 'x'
            else:
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_dmtran']).Value = 'x'

            # Altera Psicóloga
            # Zera valores das psicologas
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_meire']).Value = ''
            cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_onilce']).Value = ''

            # Coloca o valor do radio
            if self.psy_select.get_selected() == 'meire':
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_meire']).Value = 'x'
            else:
                cadastros.Range(self.config_data['planilha']['cadastro']['celula_x_onilce']).Value = 'x'

            # Altera Data
            sheet.Range(self.config_data['planilha']['entrada']['celula_dia']).Value = self.date_select.get_day()

            combo_box = sheet.Shapes(self.config_data['planilha']['entrada']['nome_combobox_mes']).OLEFormat.Object
            combo_box.Value = self.date_select.get_month()

            sheet.Range(self.config_data['planilha']['entrada']['celula_ano']).Value = self.date_select.get_year()

            # Salva e fecha
            workbook.Save()
            workbook.Close()

        excel.Quit()
        
        # Retorna o botão ao estado normal
        self.button.config(text='Gerar Planilhas e\nProcessos de Avaliação Psicológica', state='normal')

        # Finaliza com uma mensagem que deu tudo certo :)
        messagebox.showinfo(
            title='Sucesso',
            message='Planilhas e Processos de Avaliação Psicológica gerados com sucesso!'
        )
        