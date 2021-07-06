from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
import pandas as pd
import openpyxl
import os
import threading
from PIL import ImageTk, Image

from itereramappfil import *




# Bild i knapp: https://www.geeksforgeeks.org/python-add-image-on-a-tkinter-button/

class Gui:

    def __init__(self, master):

        self.master = master
        self.master.title("KBC 1.1")
        self.master.geometry("700x350")

        # Skapar tabs (använder ej flera tabs, men kan vara bra att ha framöver).
        self.tabcontrol = Notebook(master)
        self.tab1 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab1, text="fredrik.bergstrom@ltu.se")
        self.tabcontrol.pack(expand=1, fill="both")

        # Ikon
        self.master.iconbitmap('Docs\\favicon.ico')

        # LTU-logga
        self.bild_ltu_logga = ImageTk.PhotoImage(Image.open("Docs\\ltu3.jpg"))
        self.label_ltu_logga = Label(self.tab1, image=self.bild_ltu_logga)
        self.label_ltu_logga.place(y=40, x=40, anchor="center")

        # Knapp för välj mapp
        self.filvag_ebokslut = "tom"
        self.label_vald_filvag = Label(self.tab1, text="")
        self.label_vald_filvag.place(y=270, x=350, anchor="center")
        self.knapp_valj = Button(self.tab1, text="1. Välj mapp för e-bokslut", width=25, command=self.valj_filvag)
        self.knapp_valj.place(y=100, x=350, anchor="center")

        # Knapp Starta
        self.knapp_starta = Button(self.tab1, text="2. Starta", width=25, command=self.thread)
        self.knapp_starta.place(y=140, x=350, anchor="center")

        # Progressbar
        self.prog_bar = Progressbar(self.tab1, orient=HORIZONTAL, length=600, maximum=100, mode='determinate')
        self.prog_bar.place(y=250, x=350, anchor="center")

        # Label kontrollerad fil
        self.label_kontrollerad_fil = Label(self.tab1, text="", font= (20))
        self.label_kontrollerad_fil.place(y=230, x=350, anchor="center")

        # Label antal kontrollerade filare
        self.label_antal_fil = Label(self.tab1, text="Kontrollerade berpers: 0")
        self.label_antal_fil.place(y=10, x=615, anchor="center")

        # Rubrik
        self.label_rubrik = Label(self.tab1, text="Kontroll av berper", font=('Helvetica', 20, 'bold'))
        self.label_rubrik.place(y=50, x=350, anchor="center")






    def valj_filvag(self):
        # Tilldelar variabel vald filväg och uppdaterar label.
        # Satte Stringvar() som self. pga att metoden är knuten till en knapp och filvägen behövs i andra metoder.
        self.filvag_ebokslut = StringVar()
        self.vald_filvag = filedialog.askdirectory()
        self.filvag_ebokslut.set(self.vald_filvag)
        #label_vald_filvag = Label(self.tab1)
        self.label_vald_filvag["text"] = self.vald_filvag


    def thread(self):
        # Använder en annan thread så att gränssnittet inte fryser medans huvudprogrammet körs.
        # Startar lokalt i en egen metod eftersom detta måste instansieras på nytt, annars: RuntimeError: threads can only be started once.
        t = threading.Thread(target=self.starta, daemon=True)
        t.start()

    def starta(self):

        if self.filvag_ebokslut == "tom":
            messagebox.showerror("OBS!", "Du behöver välja en filväg innan du börjar.")
        else:
            self.prog_bar.start(1)
            iterera = ItereraMappFil(self.filvag_ebokslut, self.label_kontrollerad_fil, self.label_antal_fil)
            self.prog_bar.stop()

def main():
    master = ThemedTk(theme="arc")
    gui = Gui(master)
    master.mainloop()


if __name__ == '__main__':
    main()


