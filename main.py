from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import openpyxl



def get_check(numnamn_get, projektnummer):
    #print("hej")
    for x, y in zip(numnamn_get, projektnummer):
        print(x.get(), y)
       # plats = x[0]
        #print(plats.get())


def pop_droplist(root, finansiarer):
    clicked = StringVar()
    clicked.set(finansiarer[0])
    drop = OptionMenu(root, clicked, *finansiarer)
    drop.pack()


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
                print(type(proj))
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
    #print(numnamn)
    finansiarer_get = list(enumerate(finansiarer))
    #pop_droplist(root, finansiarer)

    numnamn_get = list(enumerate(numnamn))


    rad = 0
    for x in numnamn_get:
        plats = x[0]
        proj = x[1]
        numnamn_get[plats] = IntVar()
        l = Checkbutton(root, text=proj, variable=numnamn_get[plats])
        l.grid(row=rad, column=0, sticky="W")

        #numnamn_get[plats] = StringVar()
        #numnamn_get[plats].set("Välj finansiär")
       # drop = OptionMenu(root, numnamn_get[plats], *finansiarer)
        #drop.grid(row=rad, column=1, sticky="W")

        rad += 1

    B = Button(root, text="get_check", command=lambda: get_check(numnamn_get, projektnummer))
    B.grid(row=rad + 1, column=0)






    #rad = 0
    #for x in finansiarer_get:
        #plats = x[0]
        #fin = x[1]
        #finansiarer_get[plats] = IntVar()
        #l = Checkbutton(root, text=fin, variable=finansiarer_get[plats])
        #l.grid(row=rad, column=0)
        #drop = OptionMenu(root, finansiarer_get[plats], *finansiarer_get)
        #drop.grid(row=rad, column=1)
        #rad += 1





    #B = Button(root, text="get_check", command=lambda: get_check(finansiarer, finansiarer_get))
    #B.grid(row=rad+1, column=0)




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
