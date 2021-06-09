from tkinter import *
from tkinter.ttk import *
from ttkthemes import ThemedTk
from tkinter import filedialog
from PIL import ImageTk, Image
import os



# Bild i knapp: https://www.geeksforgeeks.org/python-add-image-on-a-tkinter-button/


class Tab1:

    def __init__(self, tab1):

        self.tab1 = tab1

        # LTU-logga
        self.bild_ltu_logga = ImageTk.PhotoImage(Image.open("Docs\\ltu3.jpg"))
        self.label_ltu_logga = Label(self.tab1, image=self.bild_ltu_logga)
        self.label_ltu_logga.place(y=100, x=100, anchor="center")


        # Knapp för välj mapp
        #self.knapp_valj = Button(self.tab1, text="Välj mapp för e-bokslut", width=25, command="self.get_folder_path", image=self.bild_ltu_logga)
        #self.knapp_valj.place(y=100, x=200, anchor="center")


        # Knapp Starta
        #self.knapp_starta = Button(tab1, text="Starta", width=25, command="self.thread")
        #self.knapp_starta.place(y=140, x=200, anchor="center")

