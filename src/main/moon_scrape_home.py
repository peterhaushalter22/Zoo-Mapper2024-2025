from tkinter.ttk import Style
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from tokenize import Double
from numpy import double, true_divide
import pandas as pd
import numpy as np
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
from moon_scrape_doc_to_excel import Doc_To_Excel_Moon_Scrape_Page
from moon_scrape_excel_to_excel import Excel_To_Excel_Moon_Scrape_Page
from moon_scrape_excel_to_sheet import Excel_To_Sheet_Moon_Scrape_Page
import heatmappage

#import Negating_row

LARGE_FONT = ("Bell Gothic Std Black", 40, 'bold')
MEDIUM_FONT = ("Bell Gothic Std Black", 25, 'bold')
BUTTON_FONT = ('Calibiri', 14, 'bold')
BACKGROUND_COLOR = '#407297'
LIGHT_BLUE = '#d4e1fa'

class Moon_Scrape_Home_Page(tk.Frame):
	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent, bg="white")

		# Load Moon Icon
		icon_path = heatmappage.resource_path("src\\main\\resources\\icons\\moon.png")
		self.moon_icon = ImageTk.PhotoImage(PILImage.open(icon_path).resize((50, 50), PILImage.LANCZOS))

		# Title with Icon
		title = tk.Label(self, text="Moon Scrape", font=("Segoe UI", 48, "bold"),
						 bg="white", fg="#333333", image=self.moon_icon, compound="left", padx=10)
		title.pack(pady=(40, 20))

		# Button container
		button_frame = tk.Frame(self, bg="white")
		button_frame.pack(pady=10)

		# Button styles
		button_font = ("Segoe UI", 16, "bold")
		button_padx = 20
		button_pady = 12

		# Button definitions
		buttons = [
			("Scrape Info from Google Doc to Excel Sheet", lambda: controller.show_frame(Doc_To_Excel_Moon_Scrape_Page)),
			("Scrape Info from Excel and Create New Excel", lambda: controller.show_frame(Excel_To_Excel_Moon_Scrape_Page)),
			("Scrape Info from Excel and Add New Sheet to It", lambda: controller.show_frame(Excel_To_Sheet_Moon_Scrape_Page)),
			("Back to Home", lambda: controller.show_frame(heatmappage.StartPage))
		]

		for text, cmd in buttons:
			b = tk.Button(button_frame, text=text, command=cmd,
						  font=button_font, bg="white", fg="black", relief="solid", bd=2,
						  padx=button_padx, pady=button_pady)
			b.pack(pady=10, ipadx=10, fill="x")

			# Hover Effects
			b.bind("<Enter>", lambda e, btn=b: btn.config(bg="#e6f2ff", highlightbackground="#3399FF"))
			b.bind("<Leave>", lambda e, btn=b: btn.config(bg="white", highlightbackground="black"))
