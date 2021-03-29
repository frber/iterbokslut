from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd



def pop_droplist(root, finansiarer):
    clicked = StringVar()
    clicked.set(finansiarer[0])
    drop = OptionMenu(root, clicked, *finansiarer)
    drop.pack()

def hamta_fin():
    df = pd.read_excel(r'Docs\Data.xlsx')
    fin = df['FINANSIÄR'].tolist()
    return fin

def main():
    root = Tk()
    root.geometry("680x350")
    finansiarer = hamta_fin()
    pop_droplist(root, finansiarer)
    root.mainloop()

if __name__ == "__main__":
    main()
