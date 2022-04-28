#System variable and io handling
import sys
import os
from pathlib import Path

#Regular expression handling
import multiprocessing
from multiprocessing import Process , Queue, Manager
import queue 

from openpyxl import load_workbook

#GUI
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
# Import sleep function
from time import sleep

import webbrowser

# It's strongly recommended name global variables with capital letters
TOOL = "Template"
VERNUM = '0.0.1a'
VER = TOOL  + " " +  VERNUM
DELAY1 = 20

#**********************************************************************************
# UI handle ***********************************************************************
#**********************************************************************************

class Application(Frame):
	def __init__(self, Root, Queue = None, Manager = None,):
		
		Frame.__init__(self, Root) 
		#super().__init__()
		self.parent = Root 
		
		self.parent.resizable(False, False)
		self.parent.title(VER)

		# When the function is execute, it will run on another process
		# The Queue is used to send data from that process to the main process
		self.Process_Queue = Queue['Process_Queue']
		self.Result_Queue = Queue['Result_Queue']
		self.Status_Queue = Queue['Status_Queue']
		self.Debug_Queue = Queue['Debug_Queue']
		# Manager is used to share data between processes
		self.Manager = Manager['Default_Manager']

		# UI Variable
		self.Button_Width_Full = 20
		self.Button_Width_Half = 15
		
		self.PadX_Half = 5
		self.PadX_Full = 10
		self.PadY_Half = 10
		self.PadY_Full = 20
		self.StatusLength = 120
		from libs.languagepack import LanguagePackEN as LanguagePack
		self.LanguagePack = LanguagePack
		self.File_Path = ''

	
		# Init function
		
		self.Notice = StringVar()
		self.Debug = StringVar()
		self.Progress = StringVar()
	
		# Generate Menu tab (Help, About, etc)
		self.Generate_Menu()
		# Add Tab to the windows
		self.Generate_Tab()
		# Generate the UI for Sample Tab
		self.init_UI()

	# UI init
	def init_UI(self):
		self.Generate_Sample_UI(self.Optimizer)

	def Generate_Menu(self):
		menubar = Menu(self.parent) 
		help_ = Menu(menubar, tearoff = 0) 
		menubar.add_cascade(label =  self.LanguagePack.Menu['Help'], menu = help_) 
		help_.add_command(label =  self.LanguagePack.Menu['GuideLine'], command = self.Menu_Function_Open_Main_Guideline) 
		help_.add_separator()
		help_.add_command(label =  self.LanguagePack.Menu['About'], command = self.Menu_Function_About) 
		self.parent.config(menu = menubar)

	def Generate_Tab(self):
		self.TAB_CONTROL = ttk.Notebook(self.parent)
		#Tab
		self.Optimizer = ttk.Frame(self.TAB_CONTROL)
		self.TAB_CONTROL.add(self.Optimizer, text= 'Sample Tab')
		self.TAB_CONTROL.pack(expand=1, fill="both")
		return
		
	def Generate_Sample_UI(self, Tab):
		
		Row = 1
		Label(self.Optimizer, textvariable=self.Notice).grid(row=Row, column=1, columnspan = 10, padx=5, pady=5, sticky= W)
		Row += 1
		self.RawSource = StringVar()
		Label(Tab, text=  'Path',  width = self.Button_Width_Half).grid(row=Row, column=1, columnspan=2, padx=5, pady=5, sticky= W)
	
		
		self.Str_File_Path = StringVar()
		self.TextRawSourcePath = Entry(Tab,width = 80, state="readonly", textvariable=self.Str_File_Path)
		self.TextRawSourcePath.grid(row=Row, column=3, columnspan=6, padx=5, pady=5, sticky=W+E)
		Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Browse'], command= self.BtnLoadRawSource).grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)
		Row+=1

		Label(Tab, text=  'Input text',  width = self.Button_Width_Half).grid(row=Row, column=1, columnspan=2, padx=5, pady=5, sticky= W)
		self.Sample_String_Data = Text(Tab, width = 80, height=1) #
		self.Sample_String_Data.grid(row=Row, column=3, columnspan=6, padx=5, pady=5, sticky=W+E)
		self.Sample_String_Data.insert("end", 'Data')
		Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Execute'], command= self.demo_function_multiProcessing).grid(row=Row, column=9, columnspan=2,padx=5, pady=5, sticky=W)
		
		Row+=1
		Label(Tab, text=  'Progress bar',  width = self.Button_Width_Half).grid(row=Row, column=1, columnspan=2, padx=5, pady=5, sticky= W)
		self.Optimize_Progressbar = Progressbar(Tab, orient=HORIZONTAL, length=750,  mode='determinate')
		self.Optimize_Progressbar["maximum"] = 1000
		self.Optimize_Progressbar.grid(row=Row, column=3, columnspan=7, padx=5, pady=5, sticky=W+E)

	# Menu Function
	def Menu_Function_About(self):
		messagebox.showinfo("About....", "Creator: Evan")

	def Show_Error_Message(self, ErrorText):
		messagebox.showerror('Error...', ErrorText)	

	def Menu_Function_Open_Main_Guideline(self):
		webbrowser.open_new(r"https://confluence.nexon.com/display/NWVNQA/%5BPYTHON%5D+Some+code+you+might+need")
	
	def BtnLoadRawSource(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("Workbook files", "*.xlsx *.xlsm"), ), multiple = False)	
		if filename != "":
			self.File_Path = Path(filename)
			self.Str_File_Path.set(self.File_Path)
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return

	def onExit(self):
		self.quit()

	def demo_function_multiProcessing(self):

		try:
			String_Input = self.Sample_String_Data.get("1.0", END).replace('\n', '')
		except Exception as e:
			ErrorMsg = ('Error message: ' + str(e))
			print(ErrorMsg)
			String_Input = ''
		
		try:
			while True:
				percent = self.Process_Queue.get_nowait()
				print("Remain percent: ", percent)
		except queue.Empty:
			pass
		self.New_Thread = Process(	target= demo_function, 
									kwargs= {	'input_file_path' : self.File_Path,
												'input_string' : String_Input, 
												'status_queue' : self.Status_Queue, 
												'process_queue' : self.Process_Queue, 
											},)
		self.New_Thread.start()
		self.after(DELAY1, self.wait_for_demo_function_thread)	

	def wait_for_demo_function_thread(self):
		if (self.New_Thread.is_alive()):
			while True:
				to_continue = 0
				try:
					percent = self.Process_Queue.get(0)
					self.Optimize_Progressbar["value"] = percent
					self.Optimize_Progressbar.update()
					to_continue = 0
				except queue.Empty:
					to_continue +=1
				try:
					Status = self.Status_Queue.get(0)
					self.Notice.set(Status)
					to_continue = 0
				except queue.Empty:
					to_continue +=1
				if to_continue == 2:
					break

			self.after(DELAY1, self.wait_for_demo_function_thread)
		else:
			try:
				Status = self.Status_Queue.get(0)
				if Status != None:	
					self.Notice.set(Status)
			except queue.Empty:
				pass
			self.New_Thread.terminate()

###########################################################################################
# WRITE YOUR FUNCTION HERE
###########################################################################################
def demo_function(input_file_path, input_string , status_queue, process_queue):
	status_queue.put('Processing...')
	process_queue.put(0)
	# I put the sleep here to simulate the long process
	sleep(1)
	# Show file name
	base_name = input_file_path.name
	notice_string = 'File name: ' + base_name
	# Cut the notice_string into 40 characters
	notice_string = notice_string[:40]
	status_queue.put(notice_string)
	process_queue.put(500)
	# I put the sleep here to simulate the long process
	sleep(1)
	# Show input string
	status_queue.put('Input string: ' + input_string)
	# I put the sleep here to simulate the long process
	sleep(1)
	if input_file_path != None:
		if os.path.isfile(input_file_path):
			try:
				xlsx = load_workbook(input_file_path)
				status_queue.put('File loaded: ' + str(input_file_path))
				for sheet_name in xlsx:
					status_queue.put('Sheet name: ' + str(sheet_name))
			except Exception as e:
				status_queue.put('Error message: ' + str(e))
		else:
			status_queue.put('Error message: File not found')	
	else:
		status_queue.put('Missing input file path')	
	process_queue.put(1000)
	return


###########################################################################################
# END OF YOUR  FUNCTION
###########################################################################################


def main():
	Process_Queue = Queue()
	Result_Queue = Queue()
	Status_Queue = Queue()
	Debug_Queue = Queue()
	
	MyManager = Manager()
	Default_Manager = MyManager.list()
	
	root = Tk()
	My_Queue = {}
	My_Queue['Process_Queue'] = Process_Queue
	My_Queue['Result_Queue'] = Result_Queue
	My_Queue['Status_Queue'] = Status_Queue
	My_Queue['Debug_Queue'] = Debug_Queue

	My_Manager = {}
	My_Manager['Default_Manager'] = Default_Manager

	Application(root, Queue = My_Queue, Manager = My_Manager,)
	root.mainloop()  


if __name__ == '__main__':
	if sys.platform.startswith('win'):
		multiprocessing.freeze_support()

	main()
