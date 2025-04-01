from tkinter import ttk, Tk, Toplevel

class ConfigWindow:
    def __init__(self, title: str):
        self.title = title
    
    def show(self):
        window = Toplevel()
        window.title(self.title)
        window.geometry('300x200')