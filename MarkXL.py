from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import tkinter as tk
import shelve, openpyxl, os

class GUI(Tk):
    def __init__(self):
        super(GUI, self).__init__()
        self.title('MarkXL')
        self.geometry('500x450')
        self.resizable(0, 0)
        self.browsebutton = Button(self, text = 'Browse for a file', command = self.fileDialog)
        self.browsebutton.pack()
        self.browsebutton.place(relx = 0.37, y=10)

        self.save_button = Button(self, text = 'Save', default = 'active', borderwidth = 2)
        self.save_button.pack()
        self.save_button.place(relx = 0.9, rely = 0.9)

        self.path_complete = False
        self.newwin_created = False
        self.confirmwin_created = False
        self.markexists_created = False
        self.exmarkexists_created = False
        self.confirm_created = False

    def fileDialog(self):
        if not self.path_complete:
            if not os.path.exists('.\\res'):
                os.makedirs('.\\res')
            self.shelfFile = shelve.open('.\\res\\mydata')
            try:
                self.path = self.shelfFile['path']
                self.path = os.path.dirname(os.path.abspath(self.path))
            except:
                self.path = '.\\'
                self.shelfFile['path'] = self.path
            try:
                self.filename = filedialog.askopenfilename(initialdir = self.path, title = 'Select an Excel File', filetype = [('Excel File', '.xlsx')] )
            except:
                self.filename = filedialog.askopenfilename(initialdir = '.\\', title = 'Select an Excel File', filetype = [('Excel File', '.xlsx')] )

            if(not self.filename == ''):
                tk.Text(self, height=2, width=30)
                self.T = tk.Text(root, height=2, width=30)
                self.T.pack()
                self.T.insert(tk.END, self.filename)
                self.T.place(x = 113, y=40)
                self.shelfFile['path'] = self.filename
                self.shelfFile.close()
                self.confirmButton = Button(self, text = 'Confirm File', command = self.confirmfile)
                self.confirmButton.pack()
                self.confirmButton.place(x = 193, y=80)
        else:
            messagebox.showinfo("Title", """You have already conirmed the Excel file to be used.
            Please restart the program to use a new file""")

    def confirmfile(self):
        self.path_complete = True
        self.open()

    def open(self):
        self.save_button = Button(self, text = 'Save', default = 'active', borderwidth = 2)
        self.save_button.pack()
        self.save_button.place(relx = 0.9, rely = 0.9)

        self.column_label = Label(self, text = 'Column Name of search field (eg A, C, AA, etc):', font=('Arial',11))
        self.column_label.pack()
        self.column_label.place(x=9, y=110)
        self.column_entry = Entry(self, width = 5)
        self.column_entry.pack()
        self.column_entry.place(x=10, y=130)

        self.start_row = Label(self, text = 'Number of the first row of the search field:', font=('Arial',11))
        self.start_row.pack()
        self.start_row.place(x=9, y=150)
        self.start_row_entry = Entry(self, width = 5)
        self.start_row_entry.pack()
        self.start_row_entry.place(x=10, y=170)

        self.end_row = Label(self, text = 'Number of the last row of the search field:',font=('Arial',11))
        self.end_row.pack()
        self.end_row.place(x=9, y=190)
        self.end_row_entry = Entry(self, width = 5)
        self.end_row_entry.pack()
        self.end_row_entry.place(x=10, y=210)

        self.confirm_search = Button(self, text = 'Confirm the search field', command = self.confirmcells)
        self.confirm_search.pack()
        self.confirm_search.place(x=10, y = 235)
        self.bind('<Return>', self.confirmcells)

    def confirmcells(self,event=None):
        self.column_value = self.column_entry.get()
        if(not self.column_value.isalpha()):
            messagebox.showinfo("Error", 'Please ensure that the column name only contains alphabets')
        else:
            self.column_value = self.column_value.upper()
            self.start_row_value = self.start_row_entry.get()
            self.end_row_value = self.end_row_entry.get()
            if self.start_row_value.isnumeric() and self.end_row_value.isnumeric():
                if int(self.start_row_value) <0 :
                    messagebox.showinfo("Error", 'The first row in the search field has to be at least 0')
                elif int(self.start_row_value) > int(self.end_row_value) :
                    messagebox.showinfo("Error", 'The last row in the search field has to be at least equal to the start row')
                else:
                    messagebox.showinfo("NICE!", 'The search field is now '+str(self.column_value)+str(self.start_row_value)+':'+str(self.column_value)+str(self.end_row_value))
                    self.search_canvas = Canvas(width=700, height=30)
                    self.search_canvas.pack()
                    self.search_canvas.place(x=150, y = 230)
                    self.search_label = Label(self, text = 'The current search field is '+str(self.column_value)+str(self.start_row_value)+':'+str(self.column_value)+str(self.end_row_value))
                    self.search_label.pack()
                    self.search_label.place(x=150, y=235)
                    self.entervalues()
            else:
                messagebox.showinfo("Error", 'Please ensure the entered rows are integers')

    def entervalues(self):
        self.entry_column = Label(self, text = 'Column Name in which marks are entered:', font=('Arial',11))
        self.entry_column.pack()
        self.entry_column.place(x=9,y=260)
        self.entry_column_value = Entry(self, width = 5)
        self.entry_column_value.pack()
        self.entry_column_value.place(x=10, y=280)

        self.mark = Label(self, text = 'Mark to be entered:', font=('Arial',11))
        self.mark.pack()
        self.mark.place(x=9,y=300)
        self.mark_value = Entry(self, width = 5)
        self.mark_value.pack()
        self.mark_value.place(x=10, y=320)

        self.search_student = Label(self, text = 'First few characters of student identification', font=('Arial',11))
        self.search_student.pack()
        self.search_student.place(x=9,y=340)
        self.search_student_value = Entry(self, width = 10)
        self.search_student_value.pack()
        self.search_student_value.place(x=10, y=360)

        self.save_button = Button(self, text = 'Save', default = 'active', borderwidth = 2, command = self.savevalues)
        self.save_button.pack()
        self.save_button.place(relx = 0.9, rely = 0.9)
        self.bind('<Return>', self.savevalues)

    def savevalues(self, value = None):
        self.search_student_value.selection_range(0, END)
        self.display = Canvas(width=400, height=200)
        self.display.pack()
        self.display.place(x=5, rely=0.89)
        if(not self.entry_column_value.get().isalpha()):
            messagebox.showinfo("Error", 'Please ensure that the column name only contains alphabets')
        else:
            wb = openpyxl.load_workbook(self.filename)
            sheet = wb[wb.sheetnames[0]]
            col = self.entry_column_value.get().upper()
            searchreturn = {}
            number = self.search_student_value.get()
            for i in range(int(self.start_row_value), int(self.end_row_value) + 1):
                 if str(sheet[self.column_value+str(i)].value).upper().startswith(str(number).upper()):
                     searchreturn[i] = str(sheet[self.column_value+str(i)].value)
                     rowvalue=i
            if (len(searchreturn) > 1):
                if not self.newwin_created:
                    self.newwin = Toplevel(self)
                    self.newwin.geometry('500x450')
                    self.newwin_created = True
                    self.newwin.focus()
                    disptext = 'More than one student was found with that identifying information: \n'
                    OPTIONS = []
                    for i in searchreturn.items():
                        disptext += i[1] + " in cell " + self.column_value + str(i[0]) +"\n"
                        OPTIONS.append(i[1])
                    display = Label(self.newwin, text=disptext)
                    display.pack()
                    wb.close()
                    self.drop = StringVar(self)
                    self.drop.set(OPTIONS[0])
                    menu = OptionMenu(self.newwin, self.drop, *OPTIONS)
                    menu.pack()
                    mulsave_button = Button(self.newwin, text = 'Save', command = self.mulsave)
                    mulsave_button.pack()
                    mulsave_button.place(relx=0.9,rely=0.9)
                    self.newwin.protocol("WM_DELETE_WINDOW", self.killnewwin)
                    self.newwin.bind('<Return>', self.mulsave)

            elif(len(searchreturn) <= 0):
                messagebox.showinfo("Error", 'No students were found with that identifier')
                wb.close()

            elif(len(searchreturn) == 1):
                self.exists = sheet[col+str(rowvalue)].value
                mark = self.mark_value.get()
                wb.close()
                self.confirm()


    def confirm(self):
        self.search_student_value.selection_range(0, END)
        if not self.confirm_created:
            self.confirm_created = True
            self.winconfirm = Toplevel(self)
            self.winconfirm.geometry('300x100')
            self.winconfirm.focus()

            col = self.entry_column_value.get().upper()
            mark = self.mark_value.get()
            number = self.search_student_value.get()
            complnumber = 0
            wb = openpyxl.load_workbook(self.filename)
            sheet = wb[wb.sheetnames[0]]

            for i in range(int(self.start_row_value), int(self.end_row_value) + 1):
                 if str(sheet[self.column_value+str(i)].value).upper().startswith(str(number).upper()):
                     rowvalue=i
                     complnumber = sheet[self.column_value+str(i)].value

            wb.close()
            confirm_label = Label(self.winconfirm, text = "Confrim entry for Student "+str(complnumber)+"?")
            confirm_label.pack()
            confirm_label.place(x=53,y=15)
            confirm_label.config(anchor=CENTER)
            confirm_yes = Button(self.winconfirm, text = 'Yes', command = self.savecon)
            confirm_yes.pack()
            confirm_yes.place(relx=0.35, y=40)
            confirm_no = Button(self.winconfirm, text = 'No', command = self.killwinconfirm)
            confirm_no.pack()
            confirm_no.place(relx=0.55, y=40)
            self.winconfirm.bind('<Return>', self.savecon)
            self.winconfirm.protocol("WM_DELETE_WINDOW", self.killwinconfirm)

    def killwinconfirm(self):
        self.confirm_created = False
        self.winconfirm.destroy()
        self.search_student_value.focus()
        self.search_student_value.selection_range(0, END)
        self.bind('<Return>', self.savevalues)

    def savecon(self, event = None):
        self.confirm_created = False
        self.winconfirm.destroy()
        self.focus()
        self.search_student_value.focus()
        self.search_student_value.selection_range(0, END)
        self.bind('<Return>', self.savevalues)
        if self.exists != None:
            self.markexists()
        else:
            self.save()


    def mulsave(self, event=None):
        self.search_student_value.selection_range(0, END)
        if not self.confirmwin_created:
            self.confirmwin = Toplevel(self.newwin)
            self.confirmwin.geometry('300x100')
            self.confirmwin_created = True
            self.confirmwin.focus()
            confirm_label = Label(self.confirmwin, text = "Confrim entry for Student "+self.drop.get()+"?")
            confirm_label.pack()
            confirm_label.place(x=60,y=15)
            confirm_label.config(anchor=CENTER)
            confirm_yes = Button(self.confirmwin, text = 'Yes', command = self.consave)
            confirm_yes.pack()
            confirm_yes.place(relx=0.35, y=40)
            confirm_no = Button(self.confirmwin, text = 'No', command = self.killsave)
            confirm_no.pack()
            confirm_no.place(relx=0.55, y=40)
            self.confirmwin.bind('<Return>', self.consave)
            self.confirmwin.protocol("WM_DELETE_WINDOW", self.killsave)

    def killnewwin(self):
        self.newwin_created = False
        self.newwin.destroy()
        self.focus()
        self.search_student_value.focus()
        self.search_student_value.selection_range(0, END)
        self.bind('<Return>', self.savevalues)

    def consave(self, event=None):
        self.search_student_value.selection_range(0, END)
        self.confirmwin_created = False
        self.newwin_created = False
        self.confirmwin.destroy()
        self.newwin.destroy()
        self.bind('<Return>', self.savevalues)
        if(not self.entry_column_value.get().isalpha()):
            messagebox.showinfo("Error", 'Please ensure that the column name only contains alphabets')
        else:
            wb = openpyxl.load_workbook(self.filename)
            sheet = wb[wb.sheetnames[0]]
            col = self.entry_column_value.get().upper()
            number = self.drop.get()
            for i in range(int(self.start_row_value), int(self.end_row_value) + 1):
                 if str(sheet[self.column_value+str(i)].value).upper().startswith(str(number).upper()):
                     rowvalue=i
                     break

            self.exists = sheet[col+str(rowvalue)].value
            mark = self.mark_value.get()
            wb.close()
            if self.exists != None:
                self.exmarkexists()
            else:
                self.exsave()

    def killsave(self):
        self.confirmwin_created = False
        self.confirmwin.destroy()
        self.newwin.focus()
        self.newwin.bind('<Return>', self.mulsave)

    def save(self):
        self.search_student_value.selection_range(0, END)
        col = self.entry_column_value.get().upper()
        mark = self.mark_value.get()
        number = self.search_student_value.get()
        complnumber = 0
        wb = openpyxl.load_workbook(self.filename)
        sheet = wb[wb.sheetnames[0]]

        for i in range(int(self.start_row_value), int(self.end_row_value) + 1):
             if str(sheet[self.column_value+str(i)].value).upper().startswith(str(number).upper()):
                 rowvalue=i
                 complnumber = sheet[self.column_value+str(i)].value

        if mark.isdigit():
            mark = int(mark)

        try:
            sheet[col+str(rowvalue)] = mark
            wb.save(self.filename)
            wb.close()
            self.search_student_value.selection_range(0, END)
            self.showmark = Label(self, text = 'Marks updated for Student '+str(complnumber)+' found in cell '+self.column_value+str(rowvalue))
            self.showmark.pack()
            self.showmark.place(x=9,rely=0.9)
            self.showmark2 = Label(self, text = 'Mark "'+str(mark)+'" added in cell '+col+str(rowvalue))
            self.showmark2.pack()
            self.showmark2.place(x=9,rely=0.94)
        except:
            wb.close()
            self.search_student_value.selection_range(0, END)
            messagebox.showinfo("Error", 'Please ensure that the Excel File is closed/not read-only and then try saving again')


    def exsave(self):
        self.search_student_value.selection_range(0, END)
        col = self.entry_column_value.get().upper()
        number = self.drop.get()

        wb = openpyxl.load_workbook(self.filename)
        sheet = wb[wb.sheetnames[0]]

        for i in range(int(self.start_row_value), int(self.end_row_value) + 1):
             if str(sheet[self.column_value+str(i)].value).upper().startswith(str(number).upper()):
                 rowvalue=i
                 break

        mark = self.mark_value.get()
        if mark.isdigit():
            mark = int(mark)
        sheet[col+str(rowvalue)] = mark
        try:
            wb.save(self.filename)
            wb.close()
            self.showmark = Label(self, text = 'Marks updated for Student '+self.drop.get()+' found in cell '+self.column_value+str(rowvalue))
            self.showmark.pack()
            self.showmark.place(x=9,rely=0.9)
            self.showmark2 = Label(self, text = 'Mark "'+str(mark)+'" added in cell '+col+str(rowvalue))
            self.showmark2.pack()
            self.showmark2.place(x=9,rely=0.94)
        except:
            wb.close()
            messagebox.showinfo("Error", 'Please ensure that the Excel File is closed/not read-only and then try saving again')

    def markexists(self):
        self.search_student_value.selection_range(0, END)
        if not self.markexists_created:
            col = self.entry_column_value.get().upper()
            wb = openpyxl.load_workbook(self.filename)
            sheet = wb[wb.sheetnames[0]]
            number = self.search_student_value.get()
            for i in range(int(self.start_row_value), int(self.end_row_value) + 1):
                 if str(sheet[self.column_value+str(i)].value).upper().startswith(str(number).upper()):
                     rowvalue=i
                     complnumber = str(sheet[self.column_value+str(i)].value)
                     break
            wb.close()
            self.markexists_created = True
            self.markexistswin = Toplevel(self)
            self.markexistswin.geometry('430x100')
            self.markexistswin.focus()
            confirm_label = Label(self.markexistswin, text = 'A mark of  "'+str(self.exists)+'" already exists for student '+complnumber+'. Do you want to update it?')
            confirm_label.pack()
            confirm_label.place(x=5,y=15)
            confirm_label.config(anchor=CENTER)
            confirm_yes = Button(self.markexistswin, text = 'Yes', command = self.saveconfirm)
            confirm_yes.pack()
            confirm_yes.place(relx=0.35, y=40)
            confirm_no = Button(self.markexistswin, text = 'No', command = self.killmarkexists)
            confirm_no.pack()
            confirm_no.place(relx=0.55, y=40)
            self.markexistswin.bind('<Return>', self.saveconfirm)
            self.markexistswin.protocol("WM_DELETE_WINDOW", self.killmarkexists)


    def exmarkexists(self):
        self.search_student_value.selection_range(0, END)
        if not self.exmarkexists_created:
            wb = openpyxl.load_workbook(self.filename)
            sheet = wb[wb.sheetnames[0]]
            number = self.drop.get()
            for i in range(int(self.start_row_value), int(self.end_row_value) + 1):
                 if str(sheet[self.column_value+str(i)].value).upper().startswith(str(number).upper()):
                     rowvalue=i
                     complnumber = str(sheet[self.column_value+str(i)].value)
                     break
            wb.close()
            self.exmarkexists_created = True
            self.exmarkexistswin = Toplevel(self)
            self.exmarkexistswin.geometry('430x100')
            self.exmarkexistswin.focus()
            confirm_label = Label(self.exmarkexistswin, text = 'A mark of  "'+str(self.exists)+'" already exists for student '+complnumber+'. Do you want to update it?')
            confirm_label.pack()
            confirm_label.place(x=5,y=15)
            confirm_label.config(anchor=CENTER)
            confirm_yes = Button(self.exmarkexistswin, text = 'Yes', command = self.exsaveconfirm)
            confirm_yes.pack()
            confirm_yes.place(relx=0.35, y=40)
            confirm_no = Button(self.exmarkexistswin, text = 'No', command = self.killexmarkexists)
            confirm_no.pack()
            confirm_no.place(relx=0.55, y=40)
            self.search_student_value.selection_range(0, END)
            self.exmarkexistswin.bind('<Return>', self.exsaveconfirm)
            self.exmarkexistswin.protocol("WM_DELETE_WINDOW", self.killexmarkexists)

    def saveconfirm(self, event = None):
        self.markexists_created = False
        self.markexistswin.destroy()
        self.bind('<Return>', self.savevalues)
        self.search_student_value.focus()
        self.search_student_value.selection_range(0, END)
        self.save()

    def exsaveconfirm(self, event = None):
        self.exmarkexists_created = False
        self.exmarkexistswin.destroy()
        self.bind('<Return>', self.savevalues)
        self.focus()
        self.search_student_value.focus()
        self.search_student_value.selection_range(0, END)
        self.exsave()

    def killmarkexists(self):
        self.markexistswin.destroy()
        self.markexists_created = False
        self.bind('<Return>', self.savevalues)
        self.focus()
        self.search_student_value.focus()
        self.search_student_value.selection_range(0, END)

    def killexmarkexists(self):
        self.exmarkexistswin.destroy()
        self.exmarkexists_created = False
        self.bind('<Return>', self.savevalues)
        self.focus()
        self.search_student_value.focus()
        self.search_student_value.selection_range(0, END)

root = GUI()
root.mainloop()
