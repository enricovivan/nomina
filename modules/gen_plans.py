from tkinter import ttk

from interface.file_select import FileSelect
from interface.date_select import DateSelect

class GenPlans:

    def __init__(
        self, 
        file_select: FileSelect, 
        date_select: DateSelect,
        gen_button: ttk.Button,
        
    ):
        self.file_select = file_select