from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
import pandas as pd
import openpyxl
import os
import threading

# Egna klasser
from skapaperforslag import *

class Gui:
    def __init__(self, master):
        self.master = master
        master.title("iterb")
        master.geometry("600x600")

        # Skapa tabs
        self.tabcontrol = Notebook(master)
        self.tab1 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab1, text="1. Start")
        self.tab2 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab2, text="2. Lägg till/Ta bort projekt")
        self.tab3 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab3, text="3. Skapa periodiseringsförslag")
        self.tabcontrol.pack(expand=1, fill="both")

        # Knapp Leta gamla berpers i - tab1
        self.knapp_gamla_berpers = Button(self.tab1, text="Leta gamla berpers i", command=self.valj_gamla_berpers)
        self.knapp_gamla_berpers.grid(row=0, column=0)

        # Label filväg gamla berpers -tab1
        self.label_gamla_berpers = Label(self.tab1, text="Ingen filväg vald")
        self.label_gamla_berpers.grid(row=0, column=1)

        # Knapp Spara nya berpers i -tab1
        self.knapp_spara_berpers = Button(self.tab1, text="Spara nya berpers i", command=self.valj_spara_berpers)
        self.knapp_spara_berpers.grid(row=1, column=0)

        # Label filväg spara berpers -tab1
        self.label_spara_berpers = Label(self.tab1, text="Ingen filväg vald")
        self.label_spara_berpers.grid(row=1, column=1)


        # Knapp Finansiärer
        self.knapp_finans = Button(self.tab1, text="Finansiärer", command=self.doc_finans)
        self.knapp_finans.grid(row=2, column=0)

        # Knapp Lägg in agressodata -tab1
        self.knapp_agressodata = Button(self.tab1, text="Lägg in agressodata", command=self.doc_agressodata)
        self.knapp_agressodata.grid(row=3, column=0)

        # Label Projektnummer -tab2
        self.projnum = Entry(self.tab2, width=30)
        self.projnum.grid(row=1, column=1)
        self.projnum_label = Label(self.tab2, text="Projektnummer:").grid(row=1, column=0)

        # Label Projektnamn -tab2
        self.projnamn = Entry(self.tab2, width=30)
        self.projnamn.grid(row=2, column=1)
        self.projnamn_label = Label(self.tab2, text="Projektnamn:")
        self.projnamn_label.grid(row=2, column=0)

        # Label Finansieringsgrad -tab2
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

        #self.s = ttk.Style()
        #self.s.theme_use("winnative")
        #self.s.configure("blue.Horizontal.TProgressbar", foreground='navy', background='navy')
        self.prog_bar = Progressbar(self.tab3, style="blue.Horizontal.TProgressbar", orient=HORIZONTAL, length=100,
                                        maximum=100, mode='determinate')
        self.prog_bar.grid(row=1, column=2)


        self.boxlist = []
        self.uppdatera_boxlista()


        # Knapp Ladda in projekt -tab3
        #self.knapp_ladda_projekt = Button(self.tab3, text="Ladda in projekt", command=self.skapa_perforslag)
        #self.knapp_ladda_projekt.grid(row=7, column=1)


        # Knapp Skapa förslag -tab3
        self.knapp_skapa_forslag = Button(self.tab3, text="Skapa periodiseringsförslag", command=self.thread)
        self.knapp_skapa_forslag.grid(row=8, column=4)

    def valj_gamla_berpers(self):
        self.filvag_gamla_berpers = StringVar()
        vald_filvag = filedialog.askdirectory()
        self.filvag_gamla_berpers.set(vald_filvag)
        self.label_gamla_berpers["text"] = vald_filvag

    def valj_spara_berpers(self):
        self.filvag_spara_berpers = StringVar()
        vald_filvag = filedialog.askdirectory()
        self.filvag_spara_berpers.set(vald_filvag)
        self.label_spara_berpers["text"] = vald_filvag

    def doc_finans(self):
        os.startfile('Docs\\Finansiarer.xlsx')

    def doc_agressodata(self):
        self.uppdatera_droplist_finans()
        os.startfile('Docs\\Agressodata.xlsx')

    def uppdatera_droplist_finans(self):
        self.finansiarer = self.hamta_fin()
        self.drop.set_menu("", *self.finansiarer)
        self.fin.set("")

    def hamta_fin(self):
        df = pd.read_excel(r'Docs\Finansiarer.xlsx')
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
        self.uppdatera_boxlista()

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
        wb.close()
        self.uppdatera_droplist_projekt()
        self.ta_bort_boxar()

    def uppdatera_boxlista(self):
        wb = openpyxl.load_workbook('Docs/Projekt.xlsx', data_only=True)
        ws = wb["Projekt"]
        lista_proj = []
        for row in ws['A2:A1000']:
            for cell in row:
                projnr = cell.value
                projnamn = cell.offset(column = 1).value
                if projnr != None:
                    if projnamn != None:
                        lista_proj.append(str(projnr)+" "+str(projnamn))
                    else:
                        lista_proj.append(str(projnr))
        self.boxlist_utfall = []

        if lista_proj:
            rad = 1
            rad2 = 1
            rad3 = 1
            for x in lista_proj:
                box = IntVar()
                checkbox = Checkbutton(self.tab3, text=x, variable=box)
                checkbox.grid(row=rad, column=0, sticky="W")
                rad += 1
                if rad > 20:
                    checkbox.grid(row=rad2, column=1, sticky="W")
                    rad2 += 1
                if rad2 > 20:
                    checkbox.grid(row=rad3, column=2, sticky="W")
                    rad3 += 1

                self.boxlist_utfall.append([box, x])
                self.boxlist.append(checkbox)
        wb.close()

    def ta_bort_boxar(self):
        for x in self.boxlist:
            x.destroy()
        self.uppdatera_boxlista()


    def thread(self):
        #Använder en annan thread så att gränssnittet inte fryser medans huvudprogrammet körs.
        #Startar lokalt i en egen metod eftersom detta måste instansieras på nytt, annars: RuntimeError: threads can only be started once.
        t = threading.Thread(target=self.skapa_perforslag, daemon=True)
        t.start()

    def skapa_perforslag(self):
        self.prog_bar.start(5)
        filvag_gamla_berpers = r'C:\Users\berfre\Desktop\gamla berper'
        filvag_spara_berpers = r'C:\Users\berfre\Desktop\testspara'
        # Lägg till för dynamiskt sen
        #filvag_gamla_berpers = self.filvag_gamla_berpers.get()
        #filvag_spara_berpers = self.filvag_spara_berpers.get()
        wb = openpyxl.load_workbook('Docs/Projekt.xlsx', data_only=True)
        ws = wb["Projekt"]
        for x in self.boxlist_utfall:
            if x[0].get() == 1:
                projnr_box = x[1].split()[0]
                for row in ws['A2:A1000']:
                    for cell in row:
                        projnr_db = cell.value
                        projnamn_db = cell.offset(column=1).value
                        finansiar_db = cell.offset(column=2).value
                        fingrad_db = cell.offset(column=3).value
                        if projnr_box == projnr_db:
                            skapa_per_forslag = SkapaPerForslag(projnr_db, projnamn_db, finansiar_db, fingrad_db, filvag_gamla_berpers, filvag_spara_berpers)
        wb.close()
        self.prog_bar.stop()

def main():
    root = ThemedTk(theme="radiance")
    gui = Gui(root)
    root.mainloop()

if __name__ == "__main__":
    main()






