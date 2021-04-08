from tkinter import *
from tkinter.ttk import *
import pandas as pd
import openpyxl



class Gui:
    def __init__(self, master):
        self.master = master
        master.title("iterb")
        master.geometry("680x350")

        # Skapa tabs
        self.tabcontrol = Notebook(master)
        self.tab1 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab1, text="Hem")
        self.tab2 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab2, text="Hantera projekt")
        self.tabcontrol.pack(expand=1, fill="both")

        # Projektnummer -tab2
        self.projnum = Entry(self.tab2, width=30)
        self.projnum.grid(row=1, column=1)
        self.projnum_label = Label(self.tab2, text="Projektnummer:").grid(row=1, column=0)

        # Projektnamn -tab2
        self.projnamn = Entry(self.tab2, width=30)
        self.projnamn.grid(row=2, column=1)
        self.projnamn_label = Label(self.tab2, text="Projektnamn:")
        self.projnamn_label.grid(row=2, column=0)

        # Finansieringsgrad -tab2
        self.fingrad = Entry(self.tab2, width=30)
        self.fingrad.grid(row=4, column=1)
        self.fingrad_label = Label(self.tab2, text="Finansieringsgrad:")
        self.fingrad_label.grid(row=4, column=0)

        # Droplist finansiärer -tab2
        self.finansiarer = self.hamta_fin()
        self.fin = StringVar()
        self.drop = OptionMenu(self.tab2, self.fin, "", *self.finansiarer)
        self.drop.grid(row=3, column=1, sticky="W")
        self.fin_label = Label(self.tab2, text="Finansiär:")
        self.fin_label.grid(row=3, column=0)

        # Droplist projekt från db -tab2
        self.lista_projekt_i_db = self.hamta_projekt()
        self.proj_db = StringVar()
        self.drop2 = OptionMenu(self.tab2, self.proj_db, "", *self.lista_projekt_i_db)
        self.drop2.grid(row=6, column=1, sticky="W")

        # Knapp Lägg till -tab2
        self.knapp_lagg_till = Button(self.tab2, text="Lägg till", command=self.spara_till_db)
        self.knapp_lagg_till.grid(row=5, column=1)

        # Knapp Ta bort -tab2
        self.knapp_ta_bort = Button(self.tab2, text="Ta bort", command=self.ta_bort_fran_db)
        self.knapp_ta_bort.grid(row=7, column=1)

    def hamta_fin(self):
        df = pd.read_excel(r'Docs\Data.xlsx')
        fin = df['FINANSIÄR'].tolist()
        return fin

    def hamta_projekt(self):
        wb = openpyxl.load_workbook('Docs/Projekt.xlsx', data_only=True)
        ws = wb["Projekt"]
        lista_projekt_i_db = []
        for row in ws['A2:A1000']:
            for cell in row:
                projnum = cell.value
                projnamn = cell.offset(column=1).value
                if projnum != None:
                    if projnamn != None:
                        lista_projekt_i_db.append(str(projnum) + " " + str(projnamn))
                    else:
                        lista_projekt_i_db.append(str(projnum))
        wb.close()
        return lista_projekt_i_db

    def spara_till_db(self):
        projektnummer = self.projnum.get()
        self.projnum.delete(0, END)
        projektnamn = self.projnamn.get()
        self.projnamn.delete(0, END)
        finansiar = self.fin.get()
        self.fin.set("")
        finansieringsgrad = self.fingrad.get()
        self.fingrad.delete(0, END)
        wb = openpyxl.load_workbook('Docs/Projekt.xlsx', data_only=True)
        ws = wb["Projekt"]
        ws.cell(row=ws.max_row + 1, column=1).value = projektnummer
        ws.cell(row=ws.max_row, column=2).value = projektnamn
        ws.cell(row=ws.max_row, column=3).value = finansiar
        ws.cell(row=ws.max_row, column=4).value = finansieringsgrad
        wb.save('Docs/Projekt.xlsx')
        wb.close()
        self.uppdatera_droplist_projekt()

    def uppdatera_droplist_projekt(self):
        self.lista_projekt_i_db = self.hamta_projekt()
        self.drop2.set_menu("", *self.lista_projekt_i_db)
        self.proj_db.set("")

    def ta_bort_fran_db(self):
        wb = openpyxl.load_workbook('Docs/Projekt.xlsx', data_only=True)
        ws = wb["Projekt"]
        projnr = self.proj_db.get().split()[0]
        c = 1
        for row in ws['A2:A1000']:
            for cell in row:
                c+=1
                if cell.value == projnr:
                    ws.delete_rows(c)
        wb.save('Docs/Projekt.xlsx')
        self.uppdatera_droplist_projekt()

def main():
    root = Tk()
    gui = Gui(root)
    root.mainloop()

if __name__ == "__main__":
    main()






