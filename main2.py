from tkinter import *
from tkinter.ttk import *
#import tkinter.ttk as ttk
#import tkinter as tk
#import tkinter.ttk as ttk
import pandas as pd
import openpyxl




def spara_till_db(projnum, projnamn, fin, fingrad):

    projeknummer = projnum.get()
    projnum.delete(0, END)
    print(projeknummer)

    projektnamn = projnamn.get()
    projnamn.delete(0, END)
    print(projektnamn)

    finansiar = fin.get()
    fin.set("")
    print(finansiar)

    finansieringsgrad = fingrad.get()
    fingrad.delete(0, END)
    print(finansieringsgrad)


def skapa_falt(root, finansiarer):
    tabcontrol = Notebook(root)
    tab1 = Frame(tabcontrol)
    tabcontrol.add(tab1, text="Kontroll")
    #tabcontrol.pack(expand=1, fill="both")
    #ttk.Style().configure("TNotebook", bg="black")
    tab2 = Frame(tabcontrol)
    tabcontrol.add(tab2, text="Hantera dokument")
    tabcontrol.pack(expand = 1, fill ="both")



    projnum = Entry(tab2, width=30)
    projnum.grid(row=1, column=1)
    projnum_label = Label(tab2, text="Projektnummer:").grid(row=1, column=0)


    projnamn = Entry(tab2, width=30)
    projnamn.grid(row=2, column=1)
    projnamn_label = Label(tab2, text="Projektnamn:").grid(row=2, column=0)

    fingrad = Entry(tab2, width=30)
    fingrad.grid(row=4, column=1)
    fingrad_label = Label(tab2, text="Finansieringsgrad:").grid(row=4, column=0)




    fin = StringVar()
    drop = OptionMenu(tab2, fin, "", *finansiarer).grid(row=3, column=1, sticky="W")
    fin_label = Label(tab2, text="Finansiär:").grid(row=3, column=0)





    knapp = Button(tab2, text="Lägg till", command=lambda: spara_till_db(projnum, projnamn, fin, fingrad))
    knapp.grid(row=5, column=1)


def hamta_fin():
    df = pd.read_excel(r'Docs\Data.xlsx')
    fin = df['FINANSIÄR'].tolist()
    return fin

def main():
    root = Tk()
    root.geometry("680x350")

    finansiarer = hamta_fin()

    skapa_falt(root, finansiarer)

    root.mainloop()

if __name__ == "__main__":
    main()