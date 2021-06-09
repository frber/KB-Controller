from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
import pandas as pd
import openpyxl
import os
import threading

from tab1 import *


# Bild i knapp: https://www.geeksforgeeks.org/python-add-image-on-a-tkinter-button/

class Gui:

    def __init__(self, master):

        self.master = master
        self.master.title("KB Controller 1.1")
        self.master.geometry("500x500")

        # Skapar tabs (använder ej flera tabs, men kan vara bra att ha framöver).
        self.tabcontrol = Notebook(master)
        self.tab1 = Frame(self.tabcontrol)
        self.tabcontrol.add(self.tab1, text="  ")
        self.tabcontrol.pack(expand=1, fill="both")

        # Ikon
        self.master.iconbitmap('Docs\\favicon.ico')


        # LTU-logga
        self.bild_ltu_logga = ImageTk.PhotoImage(Image.open("Docs\\ltu3.jpg"))
        self.label_ltu_logga = Label(self.tab1, image=self.bild_ltu_logga)
        self.label_ltu_logga.place(y=100, x=100, anchor="center")

        # Knapp för välj mapp
        self.knapp_valj = Button(self.tab1, text="Välj mapp för e-bokslut", width=25, command="self.get_folder_path")
        self.knapp_valj.place(y=100, x=200, anchor="center")

        # Knapp Starta
        self.knapp_starta = Button(self.tab1, text="Starta", width=25, command="self.thread")
        self.knapp_starta.place(y=140, x=200, anchor="center")




def main():
    master = ThemedTk(theme="black")
    gui = Gui(master)
    master.mainloop()


if __name__ == '__main__':
    main()


