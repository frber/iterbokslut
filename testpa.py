from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import openpyxl



def get_check(lista_x, lista_y):
    for x, y in zip(lista_x, lista_y):
        print(x.get(), y.get())


def skapa():
    dyn = ["hej", "arne", "ja"]
    fin = ["Vin", "Inter", "VR"]
    lista_x = []
    lista_y = []
    rad = 0
    for x in dyn:
        x = IntVar()
        y = StringVar()

        l = Checkbutton(root, text="proj", variable=x)
        l.grid(row=rad, column=0, sticky="W")

        drop = OptionMenu(root, y, *fin)
        drop.grid(row=rad, column=1, sticky="W")

        rad += 1

        lista_x.append(x)
        lista_y.append(y)

    B = Button(root, text="get_check", command=lambda: get_check(lista_x, lista_y))
    B.grid(row=rad + 1, column=0)












    #dyn.append(var)
    #print(dyn)
    #return dyn



root = Tk()
root.geometry("680x350")
skapa()









root.mainloop()