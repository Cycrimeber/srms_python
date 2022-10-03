from tkinter import *
from openpyxl import Workbook
from openpyxl import load_workbook

class Students:

    def __init__(self, master):
        self.master = master
        self.fram = Frame(self.master).pack(fill=BOTH, expand=1)

        btn = Button(self.fram, text="Exit", command=self.exitWindow)
        btn.place(x=100, y=100)

        btnAddStudent = Button(self.fram, text="Save Student Info", command=self.saveStudentInf)
        btnAddStudent.place(x=100, y=50)

        # self.fram.pack(fill=BOTH, expand=1)

    def exitWindow(self):
        exit()

    def saveStudentInf(self):
        studentDatabase = Workbook()
        sheet = studentDatabase.active

        sheet['A1'] = "Name"
        sheet['B1'] = "CA"
        sheet['C1'] = "Exam"
        sheet['D1'] = "Total"

        sheet['A2'] = "Isa Muhammed"
        sheet['B2'] = "10"
        sheet['C2'] = "30"
        sheet['D2'] = int(sheet['B2'].value) + int(sheet['C2'].value)

        sheet['A3'] = "James Peter"
        sheet['B3'] = "10"
        sheet['C3'] = "30"
        sheet['D3'] = int(sheet['B3'].value) + int(sheet['C3'].value)

        sheet['A4'] = "Ephraim Chukwuebuka Peter"
        sheet['B4'] = "20"
        sheet['C4'] = "30"
        sheet['D4'] = int(sheet['B4'].value) + int(sheet['C4'].value)

        sheet['A5'] = "Emily-King"
        sheet['B5'] = "50"
        sheet['C5'] = "45"
        sheet['D5'] = int(sheet['B5'].value) + int(sheet['C5'].value)
        studentDatabase.save(filename="studentDatabase.xlsx")

        msg = Label(self.fram, text="Save Successfull")
        msg.place(x=100, y=150)
        # self.fram.pack(fill=BOTH, expand=1)

