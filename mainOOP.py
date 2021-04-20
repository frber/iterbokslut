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
from tab2 import *
from tab3 import *

class Gui:
    def __init__(self, master):
        self.master = master
        master.title("iterb")
        master.geometry("1200x800")
        # Skapa tabs
        tabcontrol = Notebook(master)
        tab1 = Frame(tabcontrol)
        tabcontrol.add(tab1, text="1. Start")
        tab2 = Frame(tabcontrol)
        tabcontrol.add(tab2, text="2. Lägg till/Ta bort Finansiär")
        tab3 = Frame(tabcontrol)
        tabcontrol.add(tab3, text="3. Lägg till/Ta bort Projekt")
        #self.tab4 = Frame(self.tabcontrol)
        #self.tabcontrol.add(self.tab4, text="3. Skapa periodiseringsförslag")
        tabcontrol.pack(expand=1, fill="both")

        tab1 = Tab1(tab1)
        tab2 = Tab2(tab2)
        tab3 = Tab3(tab3, tab1)


def main():
    root = ThemedTk(theme="black")
    gui = Gui(root)
    root.mainloop()

if __name__ == "__main__":
    main()






