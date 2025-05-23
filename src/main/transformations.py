from tkinter.ttk import Style
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from tokenize import Double
from numpy import double, true_divide
import pandas as pd
#import rpy2.robjects as robjects
#from rpy2.robjects import NULL, pandas2ri
#from rpy2.robjects import r
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import pandas as pd
import csv
import subprocess
import multiprocessing
import threading
import re
import json
import os
from PIL import Image as PILImage, ImageTk
import os

import heatmappage

#import Negating_row

LARGE_FONT = ("Bell Gothic Std Black", 40, 'bold')
MEDIUM_FONT = ("Bell Gothic Std Black", 25, 'bold')
BUTTON_FONT = ('Calibiri', 14, 'bold')
BACKGROUND_COLOR = '#407297'
LIGHT_BLUE = '#d4e1fa'

class Transformations_Page(tk.Frame):
    def __init__(self, parent, controller):
        """
        This function creates the landing page when users decide to run Data Transformations.
        We will be able to select the file we want to run, run it and return to the home page
        """
        tk.Frame.__init__(self, parent, bg="white")

        # Load moon icon for visual consistency
        icon_path = "src\\main\\resources\\icons\\refresh.png"
        self.icon = ImageTk.PhotoImage(PILImage.open(icon_path).resize((50, 50), PILImage.LANCZOS))

        # Title
        title = tk.Label(self, text="Data Transformations", font=("Segoe UI", 36, "bold"),
                         bg="white", fg="#333333", image=self.icon, compound="left", padx=10)
        title.pack(pady=(40, 20))

        # Set variables
        self.filename = "None"
        self.outputname = "None"
        self.tmp = tk.StringVar()
        self.tmp.set("hello")

        # Button container
        button_frame = tk.Frame(self, bg="white")
        button_frame.pack(pady=10)

        button_font = ("Segoe UI", 16, "bold")

        # Select File button
        select_button = tk.Button(button_frame, text="Select File",
                                  command=self.select_file,
                                  font=button_font, bg="white", fg="black", relief="solid", bd=2,
                                  padx=20, pady=10)
        select_button.pack(pady=10, ipadx=10, fill="x")

        # Run Transformations button
        run_button = tk.Button(button_frame, text="Run Transformations",
                               command=self.get_parameters,
                               font=button_font, bg="white", fg="black", relief="solid", bd=2,
                               padx=20, pady=10)
        run_button.pack(pady=10, ipadx=10, fill="x")

        # Back to Home button
        back_button = tk.Button(button_frame, text="Back to Home",
                                command=lambda: controller.show_frame(heatmappage.StartPage),
                                font=button_font, bg="white", fg="black", relief="solid", bd=2,
                                padx=20, pady=10)
        back_button.pack(pady=10, ipadx=10, fill="x")

        # Hover effects
        for b in [select_button, run_button, back_button]:
            b.bind("<Enter>", lambda e, btn=b: btn.config(bg="#e6f2ff", highlightbackground="#3399FF"))
            b.bind("<Leave>", lambda e, btn=b: btn.config(bg="white", highlightbackground="black"))

    def select_file(self):
        """
        This function is handling the selection of a file. We assume that the file is located in
        locations specified by kde_args.json, called from select file button.
        """
        self.filename = filedialog.askopenfilename(initialdir="", title="Select a File",
                                                   filetypes=(("Excel Files", "*.xlsx*"), ("CSV Files", "*.csv*"), ("All Files", "*.*")))
        if self.filename:
            ext = os.path.splitext(self.filename)[1]
            if ext in [".xlsx", ".csv"]:
                self.file_extension = ext

    def get_parameters(self):
        """
        Grabbing all of the information and parameters from the file we have selected, called from Run Transformations button.
        """
        options_box = Params_Page(self.filename)
        options_box.wait_window(options_box)



class Params_Page(tk.Toplevel):

    """_summary_
    This page allows users to select parameters for KDE calculations and
    run the KDE script. Called from the "Run Transformations" button.

    Args:
    tk.Toplevel (self.filename) the name of the file given from Run Transformations

    """


    def __init__(self, filename):
        tk.Toplevel.__init__(self)  # constucting a main window of an application and making sure it is in the front of the screen, this is a popup window
        self.attributes('-topmost', 'true')

        # We know what data we need for the calcularions so we are specifying the types they should all be
        self.filename = tk.StringVar()      # filename
        self.outputname = tk.StringVar()    # unsure
        self.invert_col = tk.StringVar()    # the column you want to invert

        self.filename.set(filename)         # setting filename to the filename the user input

        self.headers = self.get_headers(self.filename.get())            #grabbing the names of the headers from the file we input
    
        invert_col_label = tk.Label(self, text='Invert Column', bg='white')       # Name of the column you want to invert
        invert_col_label.pack()                                                   # called with keyword-option/value pairs that control where the widget is to appear within its container
        invert_col_dropdown = tk.OptionMenu(self, self.invert_col, *self.headers)   # populating column with the data stored in the  column
        invert_col_dropdown.pack()
       
        tmp_button = tk.Button(self, text="Run Transformations",
                                command=lambda: self.run_transformations())
        tmp_button.pack()

    def get_headers(self, file):
        """_summary_
        grabbing the names of the headers from the file we input

        Returns:
            header: returns the name of the header from the fike given
        """
        headers = list(pd.read_excel(file).columns)
        headers.append("N/A")
        return headers

    '''
    Selecting an output directory for the KDE calculations through python before the R script
    '''

    def select_output(self):#??
        validFile = False
        self.outputname = filedialog.askdirectory(title = "Select a Directory for Output")

    def run_transformations(self):
        self.select_output()

        df = pd.read_excel(self.filename.get(), sheet_name=0)


        df[self.invert_col.get()] = -1 * df[self.invert_col.get()]

        file_name = os.path.splitext(os.path.basename(self.filename.get()))[0]
        outdir = self.outputname + "/" + file_name + "_Inverted.xlsx"
        df.to_excel(outdir)
        messagebox.showinfo("Complete", "Data Inversion is Complete")