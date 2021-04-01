from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import openpyxl



def hamta_varden(lista_box, lista_drop):
    for box, drop in zip(lista_box, lista_drop):
        print(box.get(), drop.get())


def skapa_box_drop(root, numnamn, finansiarer):
    lista_box = []
    lista_drop = []
    rad = 0
    for x in numnamn:
        box = IntVar()
        drop = StringVar()
        checkbox = Checkbutton(root, text=x, variable=box)
        checkbox.grid(row=rad, column=0, sticky="W")
        droplist = OptionMenu(root, drop, *finansiarer)
        droplist.grid(row=rad, column=1, sticky="W")
        lista_box.append(box)
        lista_drop.append(drop)
        rad += 1

    knapp = Button(root, text="get_check", command=lambda: hamta_varden(lista_box, lista_drop))
    knapp.grid(row=rad + 1, column=0)

def hamta_projekt():
    wb = openpyxl.load_workbook('Docs/Proj.xlsx', data_only=True)
    ws = wb["Proj"]
    projnummer = []
    projnamn = []
    numnamn = []
    for row in ws['A1:A100']:
        for cell in row:
            konto = cell.value
            #print(konto)
            proj = cell.offset(column = 2).value
            #print(proj)
            projt = cell.offset(column = 3).value
            belopp = cell.offset(column = 7).value
            if konto == None and proj != None and projt != None and belopp != None:
                projnummer.append(proj)
                projnamn.append(projt)
                numnamn.append(str(proj)+ " "+projt)

    return projnummer, projnamn, numnamn

def hamta_fin():
    df = pd.read_excel(r'Docs\Data.xlsx')
    fin = df['FINANSIÄR'].tolist()
    return fin

def main():
    root = Tk()
    root.geometry("680x350")

    finansiarer = hamta_fin()
    projektnummer = hamta_projekt()[0]
    projektnamn = hamta_projekt()[1]
    numnamn = hamta_projekt()[2]

    skapa_box_drop(root, numnamn, finansiarer)

    root.mainloop()

if __name__ == "__main__":
    main()

#Ge dynamiskt periodiseringsförslag efter val av fin och fingrad?
#Sätta in förslag i ny eller tidgare använd berper (beroende på namn etc).
#Skapa rätt berpers med rätt namn utifrån kontering i berper
#Flytta över till bokföringsmall med/utan rev från berpers
#Kontrollera fel i berpers


#Lägga in data - ex kontrollera vht och göra färdigt en bokföringsorder med rättningsförslag. Kontrollera andra saker
#Visualisera kontering av personal
