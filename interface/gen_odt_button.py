from tkinter import ttk, filedialog, messagebox, Tk
from docxtpl import DocxTemplate

from modules.pdf_data import PdfData
from modules.get_fist_file_from_folder import GetFirstFileFromFolder

from interface.file_select import FileSelect
from interface.clinic_select import ClinicSelect
from interface.psicho_select import PsychoSelect
from interface.date_select import DateSelect

from classes.clinica import Clinica
from classes.psicologo import Psicologo

import os
import shutil

class GenOdtButton:
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

        self.odt_model_path = GetFirstFileFromFolder('./words', 'docx').getFilePath()

        self.button: ttk.Button

        if self.odt_model_path is None:
            messagebox.showerror(
                title='Erro',
                message='Não foi encontrado nenhum documento modelo com extensão ".docx" dentro da pasta "words".\n\nPor favor, insira um arquivo modelo na pasta e abra o programa.'
            )
            raise Exception('Arquivo docx não encontrado na pasta "words"')
            

    def render(self):

        self.button = ttk.Button(
            self.frame,
            text='Gerar Processos de Avaliação',
            command=self.gen_files
        )

        # Packs
        self.button.grid(
            row=100,
            column=0,
            padx=5,
            pady=5,
            sticky='nsew',
            columnspan=2
        )

        # Configs
        self.frame.rowconfigure(100, weight=1)

    def gen_files(self, show_message = True):
        print(f'Gerando os arquivos de Processo de Avaliação Psicológica do arquivo: {self.file_select.get_file_path_entry()}')
        print(f'Data selecionada: {self.date_select.get_day()}/{self.date_select.get_month()}/{self.date_select.get_year()}')

        # Checa se um arquivo PDF foi selecionado
        if self.file_select.get_file_path_entry() == '' or None:
            self.button.config(text='Gerar Processos de Avaliação', state='normal')
            messagebox.showerror(title='Erro', message='Por favor, selecione uma lista de candidatos!')
            return

        self.button.config(text='Gerando Processos...', state='disabled')

        save_folder = filedialog.askdirectory(
            title='Selecione a pasta para salvar os processos'
        )
        
        # Caso não seja selecionado nenhum caminho de destino, retorna o botão ao estado original
        if save_folder == '' or None:
            self.button.config(text='Gerar Processos de Avaliação', state='normal')
            # messagebox.showwarning(title='Atenção', message='Por favor, selecione uma pasta para gerar os arquivos')
            return

        save_folder = os.path.join(save_folder, 'Processos de Avaliação Psicológica')

        print(f'Pasta selecionada: {save_folder}')

        
        # Cria a pasta caso ela não exista
        os.makedirs(save_folder, exist_ok=True)

        # Pega os dados do pdf
        doc_data = PdfData(self.file_select.get_file_path_entry())
        print(doc_data.get_pdf_data())

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

        # Loop para criar cópias e modificar os arquivos .odt
        for i in range(0, len(doc_data.get_pdf_data())):
            
            # Adiciona o Nome e o Processo do candidato no contexto
            context['candidato_nome'] = doc_data.get_row(i)['Nome']
            context['candidato_processo'] = doc_data.get_row(i)['Processo']

            doc.render(context)

            new_filepath = os.path.join(save_folder, f'{doc_data.get_row(i)['Nome']} {doc_data.get_row(i)['Processo']}.docx')

            doc.save(new_filepath)

        # Retorna o botão ao estado normal
        self.button.config(text='Gerar Processos de Avaliação', state='normal')

        if show_message:
            messagebox.showinfo(
                title='Sucesso',
                message='Processos de Avaliação Psicológica geradas com sucesso!'
            )




        

