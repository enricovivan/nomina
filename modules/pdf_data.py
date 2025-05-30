import tabula as tb
import pandas as pd

import sys
import os

from pandas import DataFrame

from interface.pdf_format_select import PdfFormatSelect

class PdfData:
    
    def __init__(
        self, 
        pdf_path: str,
        pdf_type: str = 'normal'
    ):
        self.pdf_data: DataFrame
        self.pdf_type = pdf_type

        if getattr(sys, 'frozen', False):
            # Executável
            jar_path = os.path.join(sys._MEIPASS, 'tabula', 'tabula-1.0.5-jar-with-dependencies.jar')
            print(f'Caminho do jar: {jar_path}')
        else:
            # Modo desenvolvimento
            jar_path = os.path.join(os.path.dirname(__file__), 'common/deps/tabula-1.0.5-jar-with-dependencies.jar')

        # Configurações do Tabula
        tb.io._jar = jar_path

        # table_header = ['Hora', 'Coletivo', 'Nome', 'Processo', 'Pauta']
        
        # [y1, x1, y2, x2]
        tabula_area: list[list[float]]

        # Caso o modo seja normal (sem cabeçalho)
        if self.pdf_type == 'normal':
            print('Modo normal selecionado')
            tabula_area = [
                [155.82629280090333, 13.171284103393555, 768.6762928009033, 53.333784103393555],
                [155.82629280090333, 54.82128410339356, 767.9325428009033, 90.52128410339355],
                [155.82629280090333, 90.00878410339355, 767.1887928009033, 367.19628410339357],
                [155.82629280090333, 369.42753410339355, 766.4450428009034, 454.2150341033936],
                [155.82629280090333, 455.7025341033936, 765.7012928009034, 578.4212841033935]
            ]
        else:
            print('Modo com nome selecionado')
            tabula_area = [
                [180.37004280090332, 13.171284103393555, 768.6762928009033, 53.333784103393555],
                [180.37004280090332, 54.82128410339356, 767.9325428009033, 90.52128410339355],
                [180.37004280090332, 90.00878410339355, 767.1887928009033, 367.19628410339357],
                [180.37004280090332, 369.42753410339355, 766.4450428009034, 454.2150341033936],
                [180.37004280090332, 455.7025341033936, 765.7012928009034, 578.4212841033935]
            ]

        tables = tb.read_pdf(
            pdf_path,
            pages=1, 
            multiple_tables=True, 
            silent=True, 
            guess=False,
            stream=True,
            area=tabula_area
        )

        # print(tables)

        df = pd.concat(tables, axis=1)

        # print(df)

        # df = tables[0]
        
        # correção de dados
        # primeira header (primeiro dado)
        # first_data = df.columns

        # df.columns = table_header

        # df.loc[len(df)] = first_data
        # df.loc[-1] = first_data
        # df.index = df.index + 1
        # df = df.sort_index()

        # Define o data como o dataframe
        self.pdf_data = df

    def get_pdf_data(self) -> DataFrame:
        return self.pdf_data
    
    def get_row(self, index: int):
        return self.pdf_data.iloc[index]