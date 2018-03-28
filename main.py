import os
from tkinter import *
from Modules.application_window import ApplicationWindow
from Modules.certificate_mailer import CertificateMailer

fileDir = os.path.dirname(os.path.realpath(__file__))

def main():
	root = Tk()
	application_window = ApplicationWindow()
	root.geometry("825x600")
	root.mainloop()

if __name__ == '__main__':
	main()