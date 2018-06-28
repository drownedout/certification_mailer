import os
from tkinter import *
from Modules.application_window import ApplicationWindow

fileDir = os.path.dirname(os.path.realpath(__file__))

def main():
	root = Tk()
	application_window = ApplicationWindow(fileDir)
	root.geometry("825x600")
	root.mainloop()

if __name__ == '__main__':
	main()