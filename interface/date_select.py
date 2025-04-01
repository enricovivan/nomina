from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime

class DateSelect:
    def __init__(self, frame: ttk.Frame):
        self.frame = frame
        self.date: DateEntry

    def render(self):
        label = ttk.Label(
            self.frame,
            text="Data da Entrevista",
            anchor='center'
        )

        self.date = DateEntry(
            self.frame,
            date_pattern='dd/mm/yyyy',
            locale='pt_BR'
        )

        # Packs
        label.grid(
            row=6,
            column=0,
            columnspan=2,
            sticky='ew',
            padx=5,
            pady=5
        )

        self.date.grid(
            row=7,
            columnspan=2,
            sticky='ew',
            padx=5,
            pady=5
        )

    def get_day(self):
        return str(self.date.get_date().day).zfill(2)
    
    def get_month(self):
        return str(self.date.get_date().month).zfill(2)

    def get_year(self):
        return str(self.date.get_date().year).zfill(2)