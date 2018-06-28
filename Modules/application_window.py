import datetime
from tkinter import *
from tkinter.ttk import Frame, Button, Style, Label, Combobox
from .certificate_mailer import CertificateMailer

class ApplicationWindow(Frame):


	def __init__(self, base_dir):
		super().__init__()
		self.base_dir = base_dir
		self.initUI()

	def get_entries(self):

		certification_type = certification.get()
		certification_year = year.get()

		Mailer = CertificateMailer(self.base_dir, certification_type, certification_year)

		values = Mailer.mailer()

		# area.insert(INSERT,values)

	def initUI(self):
		# Window Title
		self.master.title("Certification Mailer")
		self.pack(expand=True)

		# Configuring Columns
		#self.columnconfigure(1, weight=1)
		#self.columnconfigure(3, pad=7)
		global certification
		global year
		certification = StringVar()
		year = StringVar()

		# Drop down for certification type
		certification_label = Label(self, text='Certification Type').grid(column=0, row=0, pady=2, sticky=NW)
		certification_type_entry = Combobox(self, textvariable=certification)
		certification_type_entry.grid(column=0, row=1, sticky=NW)

		# Drop down for year
		year_label = Label(self, text='Certification Year').grid(column=0, row=2, pady=2, sticky=NW)
		year_entry = Combobox(self, textvariable=year)
		year_entry.grid(column=0, row=3, pady=2, sticky=NW)

		# Creates dropdown values
		certification_type_entry['values'] = ('CCS', 'CES')
		certification_type_entry.set('CCS')
		year_entry['values'] = ('2018', '2019', '2020', '2021', '2022')
		year_entry.set(str(datetime.datetime.now().year)) # Sets default year selection to current year		

		# Buttons
		send_button = Button(self, text="Send Certificates", width=16, command=self.get_entries)
		send_button.grid(row=4, column=0, pady=4, padx=1,sticky=NW)
		close_button = Button(self, text="Close", command=self.quit, width=16)
		close_button.grid(row=0, column=4, pady=4, padx=1)

		# Header Label
		lbl = Label(self, text="Output")
		lbl.grid(sticky=NW, pady=4, padx=5, row=5)

		# Text Area
		global area
		area = Text(self)
		area.grid(row=6, column=0, columnspan=1,
		padx=1, sticky=E+W+S+N)
