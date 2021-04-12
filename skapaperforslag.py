import openpyxl
import os
import shutil








class SkapaPerForslag:

    def __init__(self, projnr_db, projnamn_db, finansiar_db, fingrad_db, filvag_gamla_berpers, filvag_spara_berpers):
        self.projnr_db = projnr_db
        self.projnamn_db = projnamn_db
        self.finansiar_db = finansiar_db
        self.fingrad_db = fingrad_db
        self.filvag_gamla_berpers = filvag_gamla_berpers
        self.filvag_spara_berpers = filvag_spara_berpers

        self.agressodata = self.hamta_agressodata()





    def hamta_agressodata(self):
        #print(self.projnr_db)
        wb_agresso = openpyxl.load_workbook('Docs/Agressodata.xlsx', data_only=True)
        ws_agresso = wb_agresso['Agressodata']
        agressodata = []
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
                    agressodata.append([konto, konto_text, projektnr, belopp])
        if c > 0:
            self.leta_gamal_berper()
            #self.skapa_berper()
        wb_agresso.close()
        return agressodata

    def leta_gamal_berper(self):
        for root, dirs, files in os.walk(self.filvag_gamla_berpers):
            for fil in files:
                filvag_gamal_berper = os.path.join(root, fil).replace("\\","/")
                bara_filnamn = str(fil.split(".")[0])
                projnr_filnamn = bara_filnamn.split(" ")[0]
                if projnr_filnamn == self.projnr_db:
                    ny_berper = self.filvag_spara_berpers+"\\"+str(self.projnr_db)+".xlsx"
                    shutil.copy(filvag_gamal_berper, ny_berper)


    def skapa_berper(self):
        orginal_berper = 'Docs\\Berper.xlsx'
        ny_berper =  self.filvag_spara_berpers+"\\"+str(self.projnr_db)+".xlsx"
        shutil.copy(orginal_berper, ny_berper)










