from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import openpyxl



def get_check(finansiarer, finansiarer_get):
    #print("hej")
    for x, y in zip(finansiarer_get, finansiarer):
        print(x.get(), y)
       # plats = x[0]
        #print(plats.get())


def pop_droplist(root, finansiarer):
    clicked = StringVar()
    clicked.set(finansiarer[0])
    drop = OptionMenu(root, clicked, *finansiarer)
    drop.pack()


def hamta_projekt()

def hamta_fin():
    df = pd.read_excel(r'Docs\Data.xlsx')
    fin = df['FINANSIÄR'].tolist()
    return fin

def main():
    root = Tk()
    root.geometry("680x350")
    finansiarer = hamta_fin()
    projekt = hamta_projekt()
    finansiarer_get = list(enumerate(finansiarer))
    #pop_droplist(root, finansiarer)

    print(finansiarer)
    for x in finansiarer_get:
        plats = x[0]
        fin = x[1]
        finansiarer_get[plats] = IntVar()
        l = Checkbutton(root, text=fin, variable=finansiarer_get[plats])
        l.pack()



    B = Button(root, text="get_check", command=lambda: get_check(finansiarer, finansiarer_get))
    B.pack()




    root.mainloop()

if __name__ == "__main__":
    main()

