import openpyxl
import os






class SkapaPerForslag:

    def __init__(self, projnr_db, projnamn_db, finansiar_db, fingrad_db, filvag_gamla_berpers, filvag_spara_berpers):
        self.projnr_db = projnr_db
        self.projnamn_db = projnamn_db
        self.finansiar_db = finansiar_db
        self.fingrad_db = fingrad_db
        self.filvag_gamla_berpers = filvag_gamla_berpers
        self.filvag_spara_berpers = filvag_spara_berpers

        self.agressodata = self.hamta_agressodata()  
        self.leta_gamal_berper()

    def hamta_agressodata(self):
        #print(self.projnr_db)
        wb_agresso = openpyxl.load_workbook('Docs/Agressodata.xlsx', data_only=True)
        ws_agresso = wb_agresso['Agressodata']
        agressodata = []
        for row in ws_agresso['C1:C1000']:
            for cell in row:
                projektnr = cell.value
                konto = cell.offset(column=-2).value
                konto_text = cell.offset(column=-1).value
                #projektnr = cell.offset(column=2).value
                belopp = cell.offset(column=5).value
                if projektnr == int(self.projnr_db):
                    agressodata.append([konto, konto_text, projektnr, belopp])
        wb_agresso.close()
        return agressodata

    def leta_gamal_berper(self):
        pass








