import openpyxl
import os
import shutil
from datetime import datetime








class SkapaPerForslag:

    def __init__(self, projnr_db, projnamn_db, finansiar_db, fingrad_db, filvag_gamla_berpers, filvag_spara_berpers):
        self.projnr_db = projnr_db
        self.projnamn_db = projnamn_db
        self.finansiar_db = finansiar_db
        self.fingrad_db = fingrad_db
        self.filvag_gamla_berpers = filvag_gamla_berpers
        self.filvag_spara_berpers = filvag_spara_berpers

        self.avgor_vilket_bokslut()
        self.hamta_agressodata()

    def avgor_vilket_bokslut(self):
        self.idag = datetime.now()
        self.ar = self.idag.strftime("%Y")
        self.manad = self.idag.strftime("%m")
        self.manad = int(self.manad.split("0")[1])

        if self.manad > 2 and self.manad < 5:
            self.bokslutperiod = "T1"
        if self.manad > 5 and self.manad < 12:
            self.bokslutperiod = "T2"
        if self.manad < 2 or self.manad == 12:
            self.bokslutperiod = "T3"

    def hamta_agressodata(self):
        wb_agresso = openpyxl.load_workbook('Docs/Agressodata.xlsx', data_only=True)
        ws_agresso = wb_agresso['Agressodata']
        self.agressodata = []
        c = 0
        for row in ws_agresso['C1:C1000']:
            for cell in row:
                projektnr = cell.value
                konto = cell.offset(column=-2).value
                konto_text = cell.offset(column=-1).value
                #projektnr = cell.offset(column=2).value
                belopp = cell.offset(column=5).value
                if str(projektnr) == str(self.projnr_db):
                    c += 1
                    self.agressodata.append([konto, konto_text, projektnr, belopp])
        wb_agresso.close()
        if c > 0:
            if not self.leta_gamal_berper():
                self.skapa_berper()

    def leta_gamal_berper(self):
        for root, dirs, files in os.walk(self.filvag_gamla_berpers):
            for fil in files:
                filvag_gamal_berper = os.path.join(root, fil).replace("\\", "/")
                if filvag_gamal_berper.endswith(".xlsx") or filvag_gamal_berper.endswith(".xlsm"):
                    storlek = os.path.getsize(filvag_gamal_berper)
                    if storlek < 700000:
                        try:
                            wb_gammal = openpyxl.load_workbook(filvag_gamal_berper, data_only=True)
                        except:
                            self.skapa_berper()
                            continue
                        for sheet in wb_gammal.worksheets:
                            projnummer_gammal = sheet.cell(32, 8).value
                            #print(projnummer_gammal)
                            if str(projnummer_gammal) == str(self.projnr_db):
                                ny_berper = self.filvag_spara_berpers + "\\" + str(self.projnr_db) + ".xlsx"
                                shutil.copy(filvag_gamal_berper, ny_berper)
                                wb = openpyxl.load_workbook(ny_berper, data_only=True)
                                ny_sheet = 'Periodiseringsförslag '+str(self.bokslutperiod)+ " "+str(self.ar)
                                wb.create_sheet(ny_sheet)
                                wb.save(ny_berper)
                                self.for_over_agressodata(wb, ny_berper, ny_sheet)
                                return True
                    else:
                        self.skapa_berper()
                        continue

                # NAMNALTERNATIV----------------
                #filvag_gamal_berper = os.path.join(root, fil).replace("\\","/")
                #bara_filnamn = str(fil.split(".")[0])
                #projnr_filnamn = bara_filnamn.split(" ")[0]
                #if projnr_filnamn == self.projnr_db:
                    #ny_berper = self.filvag_spara_berpers+"\\"+str(self.projnr_db)+".xlsx"
                    #shutil.copy(filvag_gamal_berper, ny_berper)
                    #try:
                        #wb = openpyxl.load_workbook(ny_berper, data_only=True)
                    #except:
                        #self.skapa_berper()
                        #continue
                    #ny_sheet = 'Periodiseringsförslag '+str(self.bokslutperiod)+ " "+str(self.ar)
                    #print(ny_sheet)
                    #wb.create_sheet(ny_sheet)
                    #wb.save(ny_berper)
                    #self.for_over_agressodata(wb, ny_berper, ny_sheet)
                    #return True

    def skapa_berper(self):
        orginal_berper = 'Docs\\Berper.xlsx'
        ny_berper =  self.filvag_spara_berpers+"\\"+str(self.projnr_db)+".xlsx"
        shutil.copy(orginal_berper, ny_berper)
        wb = openpyxl.load_workbook(ny_berper, data_only=True)
        ny_sheet = 'Periodiseringsförslag ' + str(self.bokslutperiod) + " " + str(self.ar)
        wb.create_sheet(ny_sheet)
        wb.save(ny_berper)
        self.for_over_agressodata(wb, ny_berper, ny_sheet)


    def for_over_agressodata(self, wb, ny_berper, ny_sheet):
        ws = wb[ny_sheet]
        for x in self.agressodata:
            konto = x[0]
            kontotext = x[1]
            projektnr = x[2]
            belopp = x[3]
            ws.cell(row=ws.max_row+1, column=1).value = konto
            ws.cell(row=ws.max_row, column=2).value = kontotext
            ws.cell(row=ws.max_row, column=3).value = projektnr
            ws.cell(row=ws.max_row, column=4).value = belopp
            wb.save(ny_berper)
            wb.close()

















