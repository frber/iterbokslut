from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
import pandas as pd
import openpyxl
import os
import threading






class Tab2:
    def __init__(self, tab2):

        self.tab2 = tab2

        #  Namn
        self.finnamn = Entry(self.tab2, width=30)
        self.finnamn.grid(row=1, column=1)
        self.finnamn_label = Label(self.tab2, text="Namn:")
        self.finnamn_label.grid(row=1, column=0)

        # Motpart
        self.motp = Entry(self.tab2, width=30)
        self.motp.grid(row=2, column=1)
        self.motp_label = Label(self.tab2, text="Motpart:")
        self.motp_label.grid(row=2, column=0)

        # Perkonto fordran
        self.perf = Entry(self.tab2, width=30)
        self.perf.grid(row=3, column=1)
        self.perf_label = Label(self.tab2, text="Perkonto fordran:")
        self.perf_label.grid(row=3, column=0)

        # Perkonto skuld
        self.pers = Entry(self.tab2, width=30)
        self.pers.grid(row=4, column=1)
        self.pers_label = Label(self.tab2, text="Perkonto skuld:")
        self.pers_label.grid(row=4, column=0)

        # Radio OH
        self.ent_c = 0
        self.lista_ent = []
        self.lista_ent_label = []
        self.radio_var = IntVar()
        self.radio_ja = Radiobutton(self.tab2, text="Godkänner all OH", variable=self.radio_var, value=1, command=self.ta_bort_entry)
        self.radio_ja.grid(row=5, column=1, sticky="W")
        self.radio_dir = Radiobutton(self.tab2, text="Godkänner % på totala kostnader", variable=self.radio_var, value=2, command=self.skapa_entry)
        self.radio_dir.grid(row=6, column=1, sticky="W")
        self.radio_dir = Radiobutton(self.tab2, text="Godkänner % på lönekostnader", variable=self.radio_var, value=3, command=self.skapa_entry)
        self.radio_dir.grid(row=7, column=1, sticky="W")
        self.radio_ingen = Radiobutton(self.tab2, text="Godkänner ingen OH", variable=self.radio_var, value=4, command=self.ta_bort_entry)
        self.radio_ingen.grid(row=8, column=1, sticky="W")

        # Listbox in
        self.lista_valda_kostnader = []
        self.test_k = ["Lönekostnader", "5612", "4928"]
        self.listbox_in = Listbox(self.tab2, height='10')
        self.listbox_in.grid(row=10, column=0)

        for a in self.test_k:
            self.listbox_in.insert(END, a)

        # Knapp för över listbox
        self.knapp_lagg_till = Button(self.tab2, text="--->", command=self.for_over_listbox)
        self.knapp_lagg_till.grid(row=10, column=1)

        # Knapp ta bort listbox
        self.knapp_ta_bort = Button(self.tab2, text="<---", command=self.ta_bort_listbox)
        self.knapp_ta_bort.grid(row=11, column=1)

        # Listbox ut
        self.listbox_ut = Listbox(self.tab2, height='10')
        self.listbox_ut.grid(row=10, column=2)

        # Knapp Spara i db
        self.knapp_spara_db = Button(self.tab2, text="Spara till databas", command=self.spara_till_db)
        self.knapp_spara_db.grid(row=12, column=0)

        # Träd -tab3
        self.tree = Treeview(self.tab2)
        self.tree['columns'] = ("Namn", "Motpart", "Perkonto fordran", "Perkonto skuld", "OH", "Ej godk kostnader")
        self.tree.column("#0", width=0, stretch=NO)
        self.tree.column("Namn", anchor=W)
        self.tree.column("Motpart", anchor=W, width=100)
        self.tree.column("Perkonto fordran", anchor=W, width=100)
        self.tree.column("Perkonto skuld", anchor=W, width=100)
        self.tree.column("OH", anchor=W)
        self.tree.column("Ej godk kostnader", anchor=W, width=600)

        self.tree.heading("#0", text="", anchor=W)
        self.tree.heading("Namn", text="Namn", anchor=W)
        self.tree.heading("Motpart", text="Motpart", anchor=W)
        self.tree.heading("Perkonto fordran", text="Perkonto fordran", anchor=W)
        self.tree.heading("Perkonto skuld", text="Perkonto skuld", anchor=W)
        self.tree.heading("OH", text="OH", anchor=W)
        self.tree.heading("Ej godk kostnader", text="Ej godk kostnader", anchor=W)

        self.tree.grid(row=13, column=1)
        self.initiera_trad()

    def skapa_entry(self):
        self.ent_c = 1
        self.oh_ent = Entry(self.tab2, width=30)
        self.oh_ent.grid(row=9, column=1)
        self.oh_ent_label = Label(self.tab2, text="Ange %")
        self.oh_ent_label.grid(row=9, column=0)
        self.lista_ent.append(self.oh_ent)
        self.lista_ent_label.append(self.oh_ent_label)

    def ta_bort_entry(self):
        if self.ent_c > 0:
            for x in self.lista_ent:
                x.destroy()
            for y in self.lista_ent_label:
                y.destroy()
        self.ent_c = 0


    def for_over_listbox(self):
        self.listbox_ut.insert(END, self.listbox_in.get(ANCHOR))
        self.lista_valda_kostnader.append(self.listbox_in.get(ANCHOR))

    def ta_bort_listbox(self):
        self.lista_valda_kostnader.remove(self.listbox_ut.get(ANCHOR))
        self.listbox_ut.delete(ANCHOR)

    def spara_till_db(self):
        self.namn = self.finnamn.get()
        self.finnamn.delete(0, END)
        self.motpart = self.motp.get()
        self.motp.delete(0, END)
        self.per_f = self.perf.get()
        self.perf.delete(0, END)
        self.per_s = self.pers.get()
        self.pers.delete(0, END)
        self.radio_varde = self.radio_var.get()
        self.radio_var.set(None)
        self.ej_godk_kostnader = self.lista_valda_kostnader

        if self.ent_c > 0:
            self.procent_oh = self.oh_ent.get()
            self.ta_bort_entry()

        self.listbox_ut.delete(0, END)

        wb = openpyxl.load_workbook('Docs/Finansiarer.xlsx', data_only=True)
        ws = wb["Data"]

        ws.cell(row=ws.max_row + 1, column=1).value = self.namn
        ws.cell(row=ws.max_row, column=2).value = self.motpart
        ws.cell(row=ws.max_row, column=3).value = self.per_f
        ws.cell(row=ws.max_row, column=4).value = self.per_s

        if self.radio_varde == 0 or self.radio_varde == None:
            pass
        if self.radio_varde == 1:
            ws.cell(row=ws.max_row, column=5).value = "All OH godkänd"
        if self.radio_varde == 2:
            ws.cell(row=ws.max_row, column=5).value = "OH på tot. kostnader"
            ws.cell(row=ws.max_row, column=6).value = self.procent_oh
        if self.radio_varde == 3:
            ws.cell(row=ws.max_row, column=5).value = "OH på lön"
            ws.cell(row=ws.max_row, column=6).value = self.procent_oh
        if self.radio_varde == 4:
            ws.cell(row=ws.max_row, column=5).value = "Ingen OH"

        col = 7
        if self.ej_godk_kostnader:
            for x in self.ej_godk_kostnader:
                ws.cell(row=ws.max_row, column=col).value = x
                col+=1


        wb.save('Docs/Finansiarer.xlsx')
        wb.close()
        self.uppdatera_trad()

    def initiera_trad(self):
        wb = openpyxl.load_workbook('Docs\\Finansiarer.xlsx', data_only=True)
        ws = wb['Data']
        self.lista_trad = []


        for row in ws['A2:A1000']:

            for cell in row:
                namn = cell.value
                motp = cell.offset(column=1).value
                per_f = cell.offset(column=2).value
                per_s = cell.offset(column=3).value
                oh = cell.offset(column=4).value
                oh_proc = cell.offset(column=5).value




                #for x in range(1, 200):
                    #ej_godk = cell.offset(column=col).value
                    #if ej_godk != None:
                       # print(ej_godk)
                if namn != None:
                    lista_ej_godk = []
                    col = 6
                    while col < 20:
                        ej_godk = cell.offset(column=col).value
                        col += 1
                        if ej_godk != None:
                            lista_ej_godk.append(ej_godk)
                    ej_g = ', '.join(lista_ej_godk)
                    self.lista_trad.append([namn, motp, per_f, per_s, oh, oh_proc, ej_g])






        wb.close()

        c = 0
        for x in self.lista_trad:
            if x[5] != None:
                self.tree.insert(parent='', index='end', iid=c, values=(x[0], x[1], x[2], x[3], x[4]+", "+x[5]+"%", x[6]))
                c += 1
            else:
                self.tree.insert(parent='', index='end', iid=c, values=(x[0], x[1], x[2], x[3], x[4], "", x[6]))
                c += 1

    def uppdatera_trad(self):
        # Cleara Treeview
        c = 0
        for x in self.lista_trad:
            self.tree.delete(c)
            c += 1
        # Populera ny lista från databas
        self.initiera_trad()
















