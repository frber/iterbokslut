from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
import pandas as pd
import openpyxl
import os
import threading

# Egna klasser
from tab1 import *
from tab3 import *

class Gui:
    def __init__(self, master):
        self.master = master
        master.title("iterb")
        master.geometry("1200x800")
        # Skapa tabs
        self.tabcontrol = Notebook(master)
        tab1 = Frame(self.tabcontrol)
        self.tabcontrol.add(tab1, text="1. Start")
        self.tab2 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab2, text="2. Lägg till/Ta bort Finansiär")
        tab3 = Frame(self.tabcontrol)
        self.tabcontrol.add(tab3, text="3. Lägg till/Ta bort Projekt")
        self.tab4 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab4, text="3. Skapa periodiseringsförslag")
        self.tabcontrol.pack(expand=1, fill="both")

        tab1 = Tab1(tab1)

        tab3 = Tab3(tab3, tab1)


def main():
    root = ThemedTk(theme="black")
    gui = Gui(root)
    root.mainloop()

if __name__ == "__main__":
    main()






