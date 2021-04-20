from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
import pandas as pd
import openpyxl
import os
import threading


from tab1 import*
from skapaperforslag import *


class Tab3:

    def __init__(self, tab3, tab1):

        self.tab3 = tab3
        self.tab1 = tab1

        # Label Projektnummer -tab3
        self.projnum = Entry(self.tab3, width=30)
        self.projnum.grid(row=1, column=1)
        self.projnum_label = Label(self.tab3, text="Projektnummer:").grid(row=1, column=0)

        # Label Projektnamn -tab3
        self.projnamn = Entry(self.tab3, width=30)
        self.projnamn.grid(row=2, column=1)
        self.projnamn_label = Label(self.tab3, text="Projektnamn:")
        self.projnamn_label.grid(row=2, column=0)

        # Label Finansieringsgrad -tab3
        self.fingrad = Entry(self.tab3, width=30)
        self.fingrad.grid(row=4, column=1)
        self.fingrad_label = Label(self.tab3, text="Finansieringsgrad:")
        self.fingrad_label.grid(row=4, column=0)

        # Droplist finansiärer -tab3
        self.finansiarer = self.hamta_fin()
        self.fin = StringVar()
        self.drop = OptionMenu(self.tab3, self.fin, "", *self.finansiarer)
        self.drop.grid(row=3, column=1, sticky="W")
        self.fin_label = Label(self.tab3, text="Finansiär:")
        self.fin_label.grid(row=3, column=0)

        # Droplist projekt från db -tab3
        # self.lista_projekt_i_db = self.hamta_projekt()
        # self.proj_db = StringVar()
        # self.drop2 = OptionMenu(self.tab3, self.proj_db, "", *self.lista_projekt_i_db)
        # self.drop2.grid(row=6, column=1, sticky="W")

        # Knapp Lägg till -tab3
        self.knapp_lagg_till = Button(self.tab3, text="Lägg till", command=self.spara_till_db)
        self.knapp_lagg_till.grid(row=5, column=1)

        # Knapp Ta bort -tab3
        self.knapp_ta_bort = Button(self.tab3, text="Ta bort", command=self.ta_bort_fran_db)
        self.knapp_ta_bort.grid(row=8, column=1)

        # Träd -tab3
        self.tree = Treeview(self.tab3)
        self.tree['columns'] = ("Projektnummer", "Projektnamn", "Finansiär", "% Finansering")
        self.tree.column("#0", width=0, stretch=NO)
        self.tree.column("Projektnummer", anchor=W)
        self.tree.column("Projektnamn", anchor=W)
        self.tree.column("Finansiär", anchor=W)
        self.tree.column("% Finansering", anchor=W)

        self.tree.heading("#0", text="", anchor=W)
        self.tree.heading("Projektnummer", text="Projektnummer", anchor=W)
        self.tree.heading("Projektnamn", text="Projektnamn", anchor=W)
        self.tree.heading("Finansiär", text="Finansiär", anchor=W)
        self.tree.heading("% Finansering", text="% Finansering", anchor=W)

        self.tree.grid(row=7, column=1)
        self.initiera_trad()

        # Prognar -tab3
        # self.s = ttk.Style()
        # self.s.theme_use("winnative")
        # self.s.configure("blue.Horizontal.TProgressbar", foreground='navy', background='navy')
        self.prog_bar = Progressbar(self.tab3, orient=HORIZONTAL, length=100, maximum=100, mode='indeterminate')
        self.prog_bar.grid(row=10, column=5)

        self.boxlist = []
        self.uppdatera_boxlista()

        # Knapp Skapa perförslag -tab3
        self.knapp_skapa_forslag = Button(self.tab3, text="Skapa Periodiseringsförslag", command=self.thread)
        self.knapp_skapa_forslag.grid(row=10, column=4)

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
        # self.uppdatera_droplist_projekt()
        self.uppdatera_boxlista()
        self.uppdatera_trad_projekt()

    def uppdatera_droplist_projekt(self):
        # Används ej just nu
        self.lista_projekt_i_db = self.hamta_projekt()
        self.drop2.set_menu("", *self.lista_projekt_i_db)
        self.proj_db.set("")

    def ta_bort_fran_db(self):
        wb = openpyxl.load_workbook('Docs/Projekt.xlsx', data_only=True)
        ws = wb["Projekt"]

        # Hämta värden projektnummer från val i Treeview
        i = self.tree.focus()
        d = self.tree.item(i)
        val = d['values']
        projnr = val[0]

        c = 1
        for row in ws['A2:A1000']:
            for cell in row:
                c += 1
                if str(cell.value) == str(projnr):
                    ws.delete_rows(c)
        wb.save('Docs/Projekt.xlsx')
        wb.close()

        # self.uppdatera_droplist_projekt()
        self.ta_bort_boxar()
        self.uppdatera_trad_projekt()

    def initiera_trad(self):
        wb = openpyxl.load_workbook('Docs\\Projekt.xlsx', data_only=True)
        ws = wb['Projekt']
        self.lista_trad = []
        for row in ws['A2:A1000']:
            for cell in row:
                projnr = cell.value
                projnamn = cell.offset(column=1).value
                fin = cell.offset(column=2).value
                fingrad = cell.offset(column=3).value
                if projnr != None:
                    self.lista_trad.append([projnr, projnamn, fin, fingrad])

        wb.close()

        c = 0
        for x in self.lista_trad:
            self.tree.insert(parent='', index='end', iid=c, values=(x[0], x[1], x[2], x[3]))
            c += 1

    def uppdatera_trad_projekt(self):

        # Cleara Treeview
        c = 0
        for x in self.lista_trad:
            self.tree.delete(c)
            c += 1

        # Populera ny lista från databas
        wb = openpyxl.load_workbook('Docs\\Projekt.xlsx', data_only=True)
        ws = wb['Projekt']
        self.lista_trad = []
        for row in ws['A2:A1000']:
            for cell in row:
                projnr = cell.value
                projnamn = cell.offset(column=1).value
                fin = cell.offset(column=2).value
                fingrad = cell.offset(column=3).value
                if projnr != None:
                    self.lista_trad.append([projnr, projnamn, fin, fingrad])

        wb.close()

        # Sätt in ny lista i Treeview
        c2 = 0
        for x in self.lista_trad:
            self.tree.insert(parent='', index='end', iid=c2, values=(x[0], x[1], x[2], x[3]))
            c2 += 1

    def uppdatera_boxlista(self):

        wb_agressodata = openpyxl.load_workbook('Docs/Agressodata.xlsx', data_only=True)
        ws_agresso = wb_agressodata['Agressodata']
        lista_proj_agresso = []
        for row in ws_agresso['C1:C1000']:
            for cell in row:
                if cell.value != None:
                    lista_proj_agresso.append(str(cell.value))

        wb_agressodata.close()

        wb = openpyxl.load_workbook('Docs/Projekt.xlsx', data_only=True)
        ws = wb["Projekt"]
        lista_proj = []
        for row in ws['A2:A1000']:
            for cell in row:
                projnr = cell.value
                projnamn = cell.offset(column=1).value
                if projnr != None:
                    if projnamn != None:
                        lista_proj.append(str(projnr) + " " + str(projnamn))
                    else:
                        lista_proj.append(str(projnr))

        self.boxlist_utfall = []
        if lista_proj_agresso and lista_proj:
            lista_proj_agresso = set(lista_proj_agresso)
            rad = 9
            rad2 = 9
            rad3 = 9
            for x in lista_proj:
                if str(x.split()[0]) in lista_proj_agresso:
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
        # Använder en annan thread så att gränssnittet inte fryser medans huvudprogrammet körs.
        # Startar lokalt i en egen metod eftersom detta måste instansieras på nytt, annars: RuntimeError: threads can only be started once.
        t = threading.Thread(target=self.skapa_perforslag, daemon=True)
        t.start()

    def skapa_perforslag(self):
        self.prog_bar.start(4)
        # filvag_gamla_berpers = r'C:\Users\berfre\Desktop\gamla berper'
        # filvag_spara_berpers = r'C:\Users\berfre\Desktop\testspara'
        # Lägg till för dynamiskt sen
        filvag_gamla_berpers = self.tab1.get_fivlag_gamla_berper()
        filvag_spara_berpers = self.tab1.get_filvag_spara_berpers()
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
                            skapa_per_forslag = SkapaPerForslag(projnr_db, projnamn_db, finansiar_db, fingrad_db,
                                                                filvag_gamla_berpers, filvag_spara_berpers)
        wb.close()
        self.prog_bar.stop()





