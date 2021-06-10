import os
import openpyxl
import shutil
from tkinter import messagebox

from utvarderafil import *
from kontrolleraberper import *
from hanteraoutput import *
from hanteralistor import *


class ItereraMappFil:


    def __init__(self, filvag_ebokslut):

        self.filvag_ebokslut = filvag_ebokslut
        self.filvag_ebokslut = self.filvag_ebokslut.get()
        self.iterera()



    def nollstall_berperdata(self):
        try:
            os.remove('Docs\\Berperdata.xlsx')
            shutil.copy('Docs\\Orginal\\Berperdata.xlsx', 'Docs\\Berperdata.xlsx')
        except PermissionError:
            return True



    def iterera(self):

        if self.nollstall_berperdata():
             messagebox.showerror("OBS!", "Du har filen för Berperdata öppen, stäng ned den och börja om.")
        else:
            berperdata = 'Docs\\Berperdata.xlsx'
            wb_berperdata = openpyxl.load_workbook(berperdata, data_only=True)


            # Listor som används för att samla in information om filer på olika ställen
            # Listorna används sedan i klassen HanteraListor
            lista_alla_filer = []
            lista_berpers = []
            lista_trasiga_filer = []
            lista_for_stora_excelfiler = []

            for root, dirs, files in os.walk(self.filvag_ebokslut):
                for fil in files:
                    utvardera_fil = UtvarderaFil(root, fil)
                    filplats = utvardera_fil.returnera_filplats()
                    lista_alla_filer.append(filplats)

                    if utvardera_fil.avgor_om_excel() and not utvardera_fil.storlek():
                        lista_for_stora_excelfiler.append(filplats)

                    sheet = utvardera_fil.avgor_om_berper()

                    if sheet == "fel":
                        lista_trasiga_filer.append(filplats)

                    else:
                        if sheet != None:
                            lista_berpers.append(filplats)
                            filnamn_clean = utvardera_fil.filnamn_clean()
                            berper = KontrolleraBerper(sheet, filnamn_clean, filplats)
                            berper = berper.validera()
                            hanteraoutput = HanteraOutput(berper, wb_berperdata, berperdata)
                            hanteraoutput.skriv_output()

            listor = HanteraListor(lista_alla_filer, lista_berpers, lista_trasiga_filer, lista_for_stora_excelfiler, wb_berperdata)
            listor.kontrollera_listor()


            try:
                wb_berperdata.save(berperdata)
                wb_berperdata.close()


                #if rakna_berper > 0:
                    #self.label_fil["text"] = "Klar"
                hanteraoutput.starta()
            except PermissionError:
                messagebox.showerror("OBS!", "Du har filen för Berperdata öppen, stäng ned den och börja om.")




















