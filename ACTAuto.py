#ACT automation script for UAT.
#Made by Andrew Maddox - Aided by Jordan Brown for ACTAuto function.
import os
#import xlrd
import xlwt
#import pathlib
import pdfplumber
import tkinter as tk
from tkinter import *
from tkinter.ttk import *
import sys
from tkinter import messagebox
from pathlib import Path
from PyPDF2 import PdfFileReader, PdfFileWriter
from datetime import datetime

#takes current directory path and makes names for future use as to not confuse the system depending on file path.
path = Path(__file__).parent.absolute()
targetfile1 = f"{path}\\target.txt"
targetfile2 = f"{path}\\target2.pdf"
#Creates a new ACT folder for ACTS if it does not already exist
dirName = f'{path}\\ACTS'
if not os.path.exists(dirName):
    os.mkdir(dirName)

#function for changing names of the files we get from our provider.
def filesnames():
    for s in os.listdir('.'):
        s = str(s)
        if s.endswith("txt"):
            if s.startswith("ACT-UNIVERSITY"):
                os.rename(s, "target.txt")

    for z in os.listdir('.'):
        if z.endswith("pdf"):
            if z.startswith("ACT-UNIVERSITY"):
                os.rename(z, "target2.pdf")

#writes import based on information from text file
def ACTAuto():
    # get todays date
    d = datetime.date(datetime.now())  # input("Today's Date MM/DD/YYYY: ")
    ddaystring = ""
    dmonthstring = ""
    dday = d.day
    if dday < 10:
        ddaystring = "0" + str(dday)
    else:
        ddaystring = str(dday)
    dmonth = d.month
    if dmonth < 10:
        dmonthstring = "0" + str(dmonth)
    else:
        dmonthstring = str(dmonth)
    dyearstring = str(d.year - 2000)

    # keep count of errors
    errors = 0

    # make new workbook
    newwb = xlwt.Workbook()
    newsheet = newwb.add_sheet('Sheet1')

    # make default font
    font = xlwt.Font()
    font.name = "Calibri"
    font.height = 11 * 20
    style = xlwt.XFStyle()
    style.font = font

    # date style format
    font = xlwt.Font()
    font.name = "Calibri"
    font.height = 11 * 20
    style2 = xlwt.XFStyle()
    style2.num_format_str = 'DD-MM-YY'

    # write the headers to the sheet
    newsheet.write(0, 0, "First Name", style)
    newsheet.write(0, 1, "Last Name", style)
    newsheet.write(0, 2, "High School Graduation Year", style)
    newsheet.write(0, 3, "Phone Number", style)
    newsheet.write(0, 4, "Street Address", style)
    newsheet.write(0, 5, "City", style)
    newsheet.write(0, 6, "State", style)
    newsheet.write(0, 7, "Postal Code", style)
    newsheet.write(0, 8, "Email", style)
    newsheet.write(0, 9, "CEEB Code", style)
    newsheet.write(0, 10, "ACT Score", style)
    newsheet.write(0, 11, "Lead Source", style)
    newsheet.write(0, 12, "Date of birth", style)
    newsheet.write(0, 13, "Gender", style)
    # newsheet.write(0, 13, "interest", style)
    newsheet.write(0, 14, "Middle Name", style)

    # make the date format
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'mm/dd/yyyy'
    date_format.font = font

    # def phone_format(n):
    #    return format(int(n[:-1), ",").replace("," "-") + n[-1]

    # file1 = open("nameholder2.txt", "w")

    # i will be iterator
    i = 0
    # open the target txt file
    with open(targetfile1, "r") as ifile:
        # loop through the lines
        for line in ifile:
            i += 1
            # write firstname
            s = line[27:43]
            s = s.rstrip()
            newsheet.write(i, 0, s.title(), style)
            # print(s)
            #file1 = open("nameholder2.txt", "a")
            #print(s, file=file1)
            # write lastname
            s = line[2:27]
            s = s.rstrip()
            newsheet.write(i, 1, s.title(), style)
            # print(s)
            #file1 = open("nameholder2.txt", "a")
            #print(s, file=file1)
            # write gradyear
            s = line[222:226]
            newsheet.write(i, 2, s, style)
            # write phone
            s = line[106:116]
            s = ("(" + s[:3] + ")" + s[3:6] + "-" + s[6:])
            newsheet.write(i, 3, s, style)
            # street
            s = line[44:84]
            s = s.rstrip()
            newsheet.write(i, 4, s.title(), style)
            # city
            s = line[116:141]
            s = s.rstrip()
            newsheet.write(i, 5, s.title(), style)
            # state
            s = line[143:145]
            newsheet.write(i, 6, s, style)
            # zip
            s = line[145:150]
            newsheet.write(i, 7, s, style)
            # email
            s = line[550:604]
            s = s.rstrip()
            s = s.lower()
            newsheet.write(i, 8, s, style)
            # hs
            s = line[204:210]
            s = '-'.join(s[i:i + 3] for i in range(0, len(s), 3))
            newsheet.write(i, 9, s, style)
            # score
            s = line[268:270]
            newsheet.write(i, 10, s, style)
            newsheet.write(i, 11, "ACT", style)
            # DOB
            s = line[100:106]
            s = datetime.strptime(s, '%m%d%y').strftime('%m/%d/%y')
            newsheet.write(i, 12, s, style)
            # Gender
            s = line[87:88]
            newsheet.write(i, 13, s, style)
            # s = line[310:311]
            # newsheet.write(i, 14, s, style)
            # Middle Initial
            s = line[43:44]
            newsheet.write(i, 14, s, style)

    # save the act as an xl file
    newwb.save("ACT Import - " + dmonthstring + "." + ddaystring + "." + dyearstring + ".xls")

    print("Successfully ran with " + str(errors) + " errors!")
    # os.system("pause")

#Grabs names from pdfs and creates smaller pdfs named after the student.
def pdffunnel():
    i = 0
    #pdf_reader = PyPDF2.PdfFileReader('target2.pdf')
    pdf_reader = PdfFileReader(open(targetfile2, 'rb'))
    num_pages = pdf_reader.numPages
    progressvalue2 = 75 % num_pages
    for i in range(0,num_pages,2):
        global progressvalue
        progressvalue = progressvalue + progressvalue2
        pdf = pdfplumber.open(targetfile2)
        page = pdf.pages[i]
        amountpages = pdf.pages
        text = page.extract_text()
        nameline = text.split("\n")[0]
        name = nameline[:nameline.index("DOB:")]
        name = name.title()
        names = name.split()
        firstname = names[0]
        lastname = names[-1]
        print(name)
        print(firstname)
        print(lastname)
        newACTdirectory = f'{path}\ACTS'
        newpdfname = f'ACT {lastname}_{firstname}.pdf'
        newpdfname = f'{newACTdirectory}\{newpdfname}'
        print(newpdfname)
        print(page)
        output = PdfFileWriter()
        #output.addPage(pdf_reader.getPage(i))
        output.addPage(pdf_reader.getPage(i))
        output.addPage(pdf_reader.getPage(i+1))
        bar2()
        with open(newpdfname, "wb") as outputStream:
            output.write(outputStream)
def finishedbox():
    window = tk.Tk()
    window.wm_withdraw()
    window.geometry("1x1+200+200")
    messagebox.showinfo(title="ACT AUTO", message="Finished!")

def mainfunction():
    bar1()
    filesnames()
    ACTAuto()
    global progressvalue
    progressvalue = 25
    bar1()
    pdffunnel()
    finishedbox()
    sys.exit(0)

def write_slogan():
    print("Andrew's ACT Auto!")

def mainGUI():
    root = tk.Tk()

    frame = tk.Frame(root)
    frame.pack()

    button = tk.Button(frame,
                    text="Start ACT Auto Program",
                    width=40,
                    height=15,
                    bg="white",
                    fg="blue",
                    command=mainfunction)
    button.pack(side=tk.LEFT)


    root.mainloop()

root = Tk()
progress = Progressbar(root, orient=HORIZONTAL,
                        length=100, mode='determinate')

def bar1():
    progress['value'] = progressvalue
    root.update_idletasks()
def bar2():
    progress['value'] = progressvalue
    root.update_idletasks()
def bar3():
    progress['value'] = progressvalue
    root.update_idletasks()
def bar4():
    progress['value'] = progressvalue
    root.update_idletasks()

def progressbartest():
    root.geometry("300x100+900+500")
    # creating tkinter window
    # Function responsible for the updation
    # of the progress bar value
    progress.pack(pady=10)
    # This button will initialize
    # the progress bar
    Button(root, text='ACT Auto - START', command=mainfunction).pack(pady=10)
    # infinite loop
    mainloop()

progressvalue = 0
progressbartest()
