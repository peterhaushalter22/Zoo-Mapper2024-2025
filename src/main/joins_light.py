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

import heatmappage
#from joins_home import Joins_Home_Page
#import Negating_row

LARGE_FONT = ("Bell Gothic Std Black", 40, 'bold')
MEDIUM_FONT = ("Bell Gothic Std Black", 25, 'bold')
BUTTON_FONT = ('Calibiri', 14, 'bold')
BACKGROUND_COLOR = '#407297'
LIGHT_BLUE = '#d4e1fa'

class Joins_Page_Light(tk.Frame):
	def __init__(self, parent, controller):
		def show_back():
			from joins_home import Joins_Home_Page
			controller.show_frame(Joins_Home_Page)

		"""
		This function creates the landing page when users decide to run Data Transformations.
		We will be able to select the file we want to
		run, run it and return to the home page
		Inputs:
			self: Represents the page that we have created
			parent:
			controller: 
		Results:
			The page will be up and ready for the user to interact with
		"""
		# Setting our variables
		self.filename = "None"           # setting the file selection to NULL
		self.filename2 = "None"
		self.outputname = "None"          # Setting the outpot name of the file to NULL
		self.tmp = tk.StringVar()
		self.tmp.set("hello")

		# Creating the frame and title
		tk.Frame.__init__(self, parent, bg="white")
		icon_path = heatmappage.resource_path("src\\main\\resources\\icons\\link.png")
		self.icon = ImageTk.PhotoImage(PILImage.open(icon_path).resize((50, 50), PILImage.LANCZOS))
		label = tk.Label(self, text="Light/Temp Data Join", font=("Segoe UI", 36, "bold"), bg="white", fg="#333333", image=self.icon, compound="left", padx=10)    # Creates the title of the web page
		label.pack(pady=(40, 20))                                                # Padding the name

		# Button container
		button_frame = tk.Frame(self, bg="white")
		button_frame.pack(pady=10)
		button_font = ("Segoe UI", 16, "bold")

		# Creating Buttons for web page
		select_button = tk.Button(button_frame, text="Select Light/Temp File",
								  command=self.select_file,
								  font=button_font, bg="white", fg="black", relief="solid", bd=2,
								  padx=20, pady=10)         # Select File button, look to function select_file to see what it does
		select_button.pack(pady=10, ipadx=10, fill="x")

		select_button2 = tk.Button(button_frame, text="Select Data File",
								   command=self.select_file2,
								   font=button_font, bg="white", fg="black", relief="solid", bd=2,
								   padx=20, pady=10)
		select_button2.pack(pady=10, ipadx=10, fill="x")

		options_button = tk.Button(button_frame, text="Run Transformations",
								  command=self.get_parameters,
								  font=button_font, bg="white", fg="black", relief="solid", bd=2,
								  padx=20, pady=10)          # Taken from kde, repurposed
		options_button.pack(pady=10, ipadx=10, fill="x")

		back_button = tk.Button(button_frame, text="Back to Joins",
							  command=show_back,
							  font=button_font, bg="white", fg="black", relief="solid", bd=2,
							  padx=20, pady=10)    # setting up the back to joins page
		back_button.pack(pady=10, ipadx=10, fill="x")

		# Hover effects
		for b in [select_button, select_button2, options_button, back_button]:
			b.bind("<Enter>", lambda e, btn=b: btn.config(bg="#e6f2ff", highlightbackground="#3399FF"))
			b.bind("<Leave>", lambda e, btn=b: btn.config(bg="white", highlightbackground="black"))

	def select_file(self):
		"""
		This function is handling the selection of a file. We assume that the file is located inlocations specified by kde_args.json
		input: 
			self: The page itself

		result: 
			self.filename is set to the file that was selected 
		"""
	   # Tk.withdraw(self)
		validFile = False       # presuming that the file the user input is not valid, needs to be proven wrong

		# grabbing the filename + path of the file that the user want to une the KBE on 
		self.filename = askopenfilename(initialdir="", title="Select a File", filetypes=(("Excel Files", "*.xlsx*"), ("CSV Files", "*.csv*"), ("All Files", "*.*")))

		file_type = self.filename[self.filename.index('.'):] # grabbing the typr of the filw

		# Checking to make sure the file is an .xslx or a .csv
		if file_type == ".xlsx":
			validFile = True
		if file_type == ".csv":
			validFile = True

		# else: # presumabley to make sure that we are not allowing a file that is not valid to be saved
		#     errorMessage(Error.FILETYPE)
		#     self.filename = ""


	def select_file2(self):
		"""
		This function is handling the selection of a file. We assume that the file is located inlocations specified by kde_args.json
		input: 
			self: The page itself

		result: 
			self.filename is set to the file that was selected 
		"""
	   # Tk.withdraw(self)
		validFile = False       # presuming that the file the user input is not valid, needs to be proven wrong

		# grabbing the filename + path of the file that the user want to une the KBE on 
		self.filename2 = askopenfilename(initialdir="", title="Select a File", filetypes=(("Excel Files", "*.xlsx*"), ("CSV Files", "*.csv*"), ("All Files", "*.*")))

		file_type = self.filename2[self.filename2.index('.'):] # grabbing the typr of the filw

		# Checking to make sure the file is an .xslx or a .csv
		if file_type == ".xlsx":
			validFile = True
		if file_type == ".csv":
			validFile = True

		# else: # presumabley to make sure that we are not allowing a file that is not valid to be saved
		#     errorMessage(Error.FILETYPE)
		#     self.filename2 = ""


	def get_parameters(self):
		"""
		grabbing all of the information and parameters from the file we have selected
		"""
		options_box = Params_Page(self.filename, self.filename2)
		options_box.wait_window(options_box)
		
		

"""
This page allows users to select parameters for KDE calculations and
run the KDE script
"""
class Params_Page(tk.Toplevel):
	def __init__(self, filename, filename2):
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
		self.filename2 = tk.StringVar()  
		self.outputname = tk.StringVar()    # unsure
		self.lightDateTime = tk.StringVar()   
		self.rawSessionStartTime = tk.StringVar()   
		self.dateTime = tk.StringVar() 
		

		self.filename.set(filename)         # setting filename to the filename the user input
		self.filename2.set(filename2)
		self.headers = self.get_headers(self.filename.get())            #grabbing the names of the headers from the file we input
		self.headers2 = self.get_headers(self.filename2.get())

		lightDateTime_label = tk.Label(self, text='Light Date Time Column', bg='white')       # Name of the column you want to invert
		lightDateTime_label.pack()                                                   # called with keyword-option/value pairs that control where the widget is to appear within its container
		lightDateTime_dropdown = tk.OptionMenu(self, self.lightDateTime, *self.headers)   # populating column with the data stored in the  column
		lightDateTime_dropdown.pack()
	   
		rawSessionStartTime_label = tk.Label(self, text='Data Session Start Time Column', bg='white')       # Name of the column you want to invert
		rawSessionStartTime_label.pack()                                                   # called with keyword-option/value pairs that control where the widget is to appear within its container
		rawSessionStartTime_dropdown = tk.OptionMenu(self, self.rawSessionStartTime, *self.headers2)   # populating column with the data stored in the  column
		rawSessionStartTime_dropdown.pack()

		dateTime_label = tk.Label(self, text='Data Date Time Column', bg='white')       # Name of the column you want to invert
		dateTime_label.pack()                                                   # called with keyword-option/value pairs that control where the widget is to appear within its container
		dateTime_dropdown = tk.OptionMenu(self, self.dateTime, *self.headers2)   # populating column with the data stored in the  column
		dateTime_dropdown.pack()

		tmp_button = tk.Button(self, text="Run Join",
								command=lambda: self.run_join())
		tmp_button.pack()

	def get_headers(self, file):
		headers = list(pd.read_excel(file).columns)
		headers.append("N/A")
		return headers

	'''
	Selecting an output directory for the KDE calculations through python before the R script
	'''

	def select_output(self):
		validFile = False
		self.outputname = filedialog.askdirectory(title = "Select a Directory for Output")

	def find_closest_time(self, df, datetime):

		#This is the difference at the start, which is datetime - 1 day
		excelDateTime = self.dateTime.get()
		behavior = self.categ.get()
		time_difference = (datetime - (datetime - pd.DateOffset(1))).total_seconds()
		returnIndex = -1

		for i, row in df.iterrows():
			#Remember to convert time to datetime
			sub_time = abs((datetime - pd.to_datetime(row[excelDateTime])).total_seconds())
			#Only currently works with this specific 'Repetitive rubbing' string, case specific
			if sub_time < time_difference and row[behavior] == 'Repetitive rubbing':
				time_difference = sub_time
				returnIndex = i
	
		return returnIndex
		
	def run_join(self):
		self.select_output()
		rawTime = self.rawSessionStartTime.get()
		lightTime = self.lightDateTime.get()
		df_raw = pd.read_excel(self.filename2.get(), sheet_name=0)
		df_light = pd.read_excel(self.filename.get(), sheet_name=0)

		df_light= df_light.rename(columns={lightTime: 'Session Start Time_dup'})

		df_raw[rawTime] = pd.to_datetime(df_raw[rawTime])
		df_light['Session Start Time_dup'] = pd.to_datetime(df_light['Session Start Time_dup'])

		df_raw['Rounded_Session_Start_time'] = df_raw[rawTime].dt.round('15min')
		df_light['Rounded_Session_Start_time'] = df_light['Session Start Time_dup'].dt.round('15min')

		df_merged = pd.merge(df_raw, df_light, on='Rounded_Session_Start_time', how="left")  
		df_merged = df_merged.drop('Session Start Time_dup', axis=1)
		df_merged= df_merged.rename(columns={'#': 'Matching Row'})

		
		#Initial join above #deals with times
		
		file_name = os.path.splitext(os.path.basename(self.filename2.get()))[0]
		outdir = self.outputname + "/" + file_name + "_Data_Join.xlsx"
		df_merged.to_excel(outdir)
		messagebox.showinfo("Complete", "Data Light/Temp Join Complete")