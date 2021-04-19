from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
import os





class Tab1:

    def __init__(self, tab1):
        self.tab1 = tab1

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

        # Knapp Finansiärer -tab1
        self.knapp_finans = Button(self.tab1, text="Finansiärer", command=self.doc_finans)
        self.knapp_finans.grid(row=2, column=0)

        # Knapp Lägg in agressodata -tab1
        self.knapp_agressodata = Button(self.tab1, text="Lägg in agressodata", command=self.doc_agressodata)
        self.knapp_agressodata.grid(row=3, column=0)

    def valj_gamla_berpers(self):
        self.filvag_gamla_berpers = StringVar()
        self.vald_filvag_gamla_berper = filedialog.askdirectory()
        self.filvag_gamla_berpers.set(self.vald_filvag_gamla_berper)
        self.label_gamla_berpers["text"] = self.vald_filvag_gamla_berper

    def valj_spara_berpers(self):
        self.filvag_spara_berpers = StringVar()
        self.vald_filvag_spara_berpers = filedialog.askdirectory()
        self.filvag_spara_berpers.set(self.vald_filvag_spara_berpers)
        self.label_spara_berpers["text"] = self.vald_filvag_spara_berpers

    def get_fivlag_gamla_berper(self):
        return self.vald_filvag_gamla_berper

    def get_filvag_spara_berpers(self):
        return self.vald_filvag_spara_berpers

    def doc_finans(self):
        os.startfile('Docs\\Finansiarer.xlsx')

    def doc_agressodata(self):
        #self.uppdatera_droplist_finans()
        os.startfile('Docs\\Agressodata.xlsx')