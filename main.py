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
        self.master.geometry("900x500")

        # Skapar tabs (använder ej flera tabs, men kan vara bra att ha framöver).
        self.tabcontrol = Notebook(master)
        self.tab1 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab1, text="Kontroll")
        self.tabcontrol.pack(expand=1, fill="both")

        # Ikon
        self.master.iconbitmap('Docs\\favicon.ico')

        # LTU-logga
        self.bild_ltu_logga = ImageTk.PhotoImage(Image.open("Docs\\ltu3.jpg"))
        self.label_ltu_logga = Label(self.tab1, image=self.bild_ltu_logga)
        self.label_ltu_logga.place(y=40, x=40, anchor="center")

        # Knapp för välj mapp
        self.knapp_valj = Button(self.tab1, text="Välj mapp för e-bokslut", width=25, command=self.valj_filvag)
        self.knapp_valj.place(y=100, x=200, anchor="center")

        # Knapp Starta
        self.knapp_starta = Button(self.tab1, text="Starta", width=25, command=self.thread)
        self.knapp_starta.place(y=140, x=200, anchor="center")

        # Progressbar
        self.prog_bar = Progressbar(self.tab1, style="blue.Horizontal.TProgressbar", orient=HORIZONTAL, length=600, maximum=100, mode='determinate')
        self.prog_bar.place(y=400, x=200)





    def valj_filvag(self):
        # Tilldelar variabel vald filväg och uppdaterar label.
        # Satte Stringvar() som self. pga att metoden är knuten till en knapp och filvägen behövs i andra metoder.
        self.filvag_ebokslut = StringVar()
        vald_filvag = filedialog.askdirectory()
        self.filvag_ebokslut.set(vald_filvag)
        label_vald_filvag = Label(self.tab1)
        label_vald_filvag["text"] = vald_filvag
        label_vald_filvag.place(y=200, x=300)

    def thread(self):
        # Använder en annan thread så att gränssnittet inte fryser medans huvudprogrammet körs.
        # Startar lokalt i en egen metod eftersom detta måste instansieras på nytt, annars: RuntimeError: threads can only be started once.
        t = threading.Thread(target=self.starta, daemon=True)
        t.start()

    def starta(self):
        self.prog_bar.start(1)
        iterera = ItereraMappFil(self.filvag_ebokslut)
        self.prog_bar.stop()

def main():
    master = ThemedTk(theme="blue")
    gui = Gui(master)
    master.mainloop()


if __name__ == '__main__':
    main()


