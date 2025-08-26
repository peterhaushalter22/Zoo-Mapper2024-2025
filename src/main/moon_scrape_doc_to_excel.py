from tkinter.ttk import Style
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from tokenize import Double
from numpy import double, true_divide
import pandas as pd
import numpy as np
import rpy2.robjects as robjects
from rpy2.robjects import NULL, pandas2ri
from rpy2.robjects import r
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from Moon_Scrape_Raw_Python import *
from PIL import Image as PILImage, ImageTk
import pandas as pd
import csv
import subprocess
import multiprocessing
import threading
import re
import json
import os

import moon_scrape_home
import heatmappage
#import Negating_row

LARGE_FONT = ("Bell Gothic Std Black", 40, 'bold')
MEDIUM_FONT = ("Bell Gothic Std Black", 25, 'bold')
BUTTON_FONT = ('Calibiri', 14, 'bold')
BACKGROUND_COLOR = '#407297'
LIGHT_BLUE = '#d4e1fa'


class Doc_To_Excel_Moon_Scrape_Page(tk.Frame):
	def __init__(self, parent, controller):
		def show_back():
			from moon_scrape_home import Moon_Scrape_Home_Page
			controller.show_frame(Moon_Scrape_Home_Page)

		"""
		This function creates the landing page when users run Moon Scrapes.
		We will be able to select the file we want to run, allow us to select the 
		excel sheet we are grabbing data fom, inputting our columns and scraping 
		the moon data 
		Inputs:
			self: Represents the page that we have created
			parent:
			controller: 
		Results:
			The page will be up and ready for the user to interact with
		"""

		# Creating the frame
		tk.Frame.__init__(self, parent, bg="white")

		# Load moon icon
		icon_path = heatmappage.resource_path("src\\main\\resources\\icons\\moon.png")
		self.moon_icon = ImageTk.PhotoImage(PILImage.open(icon_path).resize((50, 50), PILImage.LANCZOS))

		# Creating the title of the web page
		title = tk.Label(self, text="Scraping Moon Data From Doc to New Excel", font=("Segoe UI", 36, "bold"),
						 bg="white", fg="#333333", image=self.moon_icon, compound="left", padx=10)
		title.pack(pady=(40, 20))

		# Setting our variables
		self.filename = "None"           		# Variable will store the name of the file we want to moon scrape
		self.file_extension = "None"	 		# Variable that will store the extension of the file (must be .docx)
		self.info_page = "None"
		self.tmp = tk.StringVar()       		# setting self
		self.tmp.set("hello")

		# Button container
		button_frame = tk.Frame(self, bg="white")
		button_frame.pack(pady=10)

		button_font = ("Segoe UI", 16, "bold")

		# Creating Buttons for web page
		select_button = tk.Button(button_frame, text="Select File",
								  command=lambda: self.select_file(),
								  font=button_font, bg="white", fg="black", relief="solid", bd=2,
								  padx=20, pady=10)         # Select File button, look to function select_file #76 to see what it does   
		select_button.pack(pady=10, ipadx=10, fill="x")      # called with keyword-option/value pairs that control where the widget is to appear within its container

		# Will begin the process of running the webpage
		options_button = tk.Button(button_frame, text="Run Moon Scrape",
								   command=lambda: self.get_parameters(),
								   font=button_font, bg="white", fg="black", relief="solid", bd=2,
								   padx=20, pady=10)          # Taken from kde, repurposed
		options_button.pack(pady=10, ipadx=10, fill="x")

		# Button that allows you to return the homepage
		back_button = tk.Button(button_frame, text="Back to Moon Scrape Home",
								command=lambda: show_back(),
								font=button_font, bg="white", fg="black", relief="solid", bd=2,
								padx=20, pady=10)    # setting up the back to home button. goes back to start page for heat map
		back_button.pack(pady=10, ipadx=10, fill="x")

		# Hover effects
		for b in [select_button, options_button, back_button]:
			b.bind("<Enter>", lambda e, btn=b: btn.config(bg="#e6f2ff", highlightbackground="#3399FF"))
			b.bind("<Leave>", lambda e, btn=b: btn.config(bg="white", highlightbackground="black"))

	def select_file(self):
		"""
		This function is handling the selection of a file. We assume that the file is located in locations specified by kde_args.json
		input: 
			self: The page itself
		result: 
			self.filename is set to the file that was selected 
		"""
		validFile = False       # presuming that the file the user input is not valid, needs to be proven wrong

		# grabbing the filename + path of the file that the user want to run the KBE on 
		self.filename = askopenfilename(initialdir="", title="Select a File",
										filetypes=(("Doc Files", "*.docx*"), ("All Files", "*.*")))

		file_type = self.filename[self.filename.index('.'):]  # grabbing the type of the file

		# Checking to make sure the file is a .docx 
		if file_type == ".docx":
			validFile = True
			self.file_extension = file_type

	def get_parameters(self):
		"""
		grabbing all of the information and parameters from the file we have selected
		"""
		info_page = Params_Page(self.filename)
		info_page.wait_window(info_page)
			
"""
This page allows users to select parameters for KDE calculations and
run the KDE script
"""
class Params_Page(tk.Toplevel):
	def __init__(self, filename):
		"""
		Initializing everything we will use for the KDE calculatiosn
		input:
			self
			filename: Filename of file where we will be grabbing our data from
		"""
		tk.Toplevel.__init__(self)  # constucting a main window of an application and making sure it is in the front of the screen
		self.attributes('-topmost', 'true')

		# We know what data we need for the calcularions so we are specifying the types they should all be
		self.filename = tk.StringVar()      # filename
		self.latitude= tk.StringVar()
		self.longitude= tk.StringVar()
		self.new_excel_name = tk.StringVar()


		self.filename.set(filename)         		# setting filename to the filename the user input

		# Grab Latitude
		latitude_label = tk.Label(self, text="Input Latitude", bg='white')
		latitude_label.pack()
		
		latitude_entry = ttk.Entry(self, textvariable=self.latitude)
		latitude_entry.pack(fill='x', expand=True)
		latitude_entry.focus()
  
		# Grab Longitude
		longitude_label = tk.Label(self, text="Input Longitude", bg='white')
		longitude_label.pack()
		
		longitude_entry = ttk.Entry(self, textvariable=self.longitude)
		longitude_entry.pack(fill='x', expand=True)
		longitude_entry.focus()

		# Grab new Excel File Name
		new_excel_label = tk.Label(self, text="Input Name of New Excel Sheet", bg='white')
		new_excel_label.pack()
		
		new_excel_entry = ttk.Entry(self, textvariable=self.new_excel_name)
		new_excel_entry.pack(fill='x', expand=True)
		longitude_entry.focus()

		# Press to run the Scrape
		tmp_button = tk.Button(self, text="Run Moon Scrape",
								command=lambda: self.run_scrape())
		tmp_button.pack()
	
	def run_scrape(self):
		print("Filename", self.filename.get())
		print("Latitude", self.latitude.get())
		print("Longitude", self.longitude.get())
		print("New Excel Name", self.new_excel_name.get())

		doc_to_excel_Moon_Data(self.filename.get(), self.latitude.get(), self.longitude.get(), self.new_excel_name.get())