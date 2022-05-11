import shutil
from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook
import os
from tika import parser

wb = load_workbook("EXCEL.xlsx")

descs = []

def filesget():
    file = filedialog.askopenfilenames(initialdir="C:\\Downloads", defaultextension=".pdf", filetypes=[("PDF files", ".pdf")])
    global files
    files = file

def submit():

    for file in files:
        raw = parser.from_file(file)
        global filedata
        filedata = (raw['content'].split(" "))

        def indexing():

            FreeCells = []
            for cell in ws["e"]:
                if cell.value == None:
                    FreeCells.append(cell.row)

            startdatecell = "E" + str(FreeCells[0])
            endingdatecell = "F" + str(FreeCells[0])
            readingcell = "G" + str(FreeCells[0])
            dayscount = "I" + str(FreeCells[0])
            addadfee = "L" + str(FreeCells[0])

            ws[startdatecell].value = filedata[filedata.index('Class') + 2][7:17]
            ws[endingdatecell].value = filedata[filedata.index('Class') + 3][5:15]
            ws[readingcell].value = filedata[filedata.index('Capacity:') + 5][:-16]
            ws[dayscount].value = filedata[filedata.index('Class') + 5][6:8]
            ws[addadfee].value = filedata[filedata.index('10.00')]

        if filedata[filedata.index('Summary') + 3] == '10054716660':
            path = "C:\\Users\\LENOVO\\Downloads\\42"
            filename = filedata[filedata.index('Class') + 3][5:12]
            filename = filename.replace(filename[filename.index("/")], "-") + ".pdf"

            try:
                destination = os.path.join(path, filename)
                shutil.move(file, destination)
            except shutil.SameFileError:
                pass

            ws = wb['2923-42']

            indexing()

            date = filedata[filedata.index('Class') + 2][7:17] + " الى " + filedata[filedata.index('Class') + 3][5:15]
            sanadDesc = "تم سداد شقة 42 من تاريخ "
            finalDesc = sanadDesc + date

            descs.append(finalDesc)

        elif filedata[filedata.index('Summary') + 3] == '10062328912':
            path = "C:\\Users\\LENOVO\\Downloads\\5"
            filename = filedata[filedata.index('Class') + 3][5:12]
            filename = filename.replace(filename[filename.index("/")], "-") + ".pdf"

            try:
                destination = os.path.join(path, filename)
                shutil.move(file, destination)
            except shutil.SameFileError:
                pass

            ws = wb['5675- 5B']

            indexing()

            date = filedata[filedata.index('Class') + 2][7:17] + " الى " + filedata[filedata.index('Class') + 3][5:15]
            sanadDesc = "تم سداد شقة 5 من تاريخ "
            finalDesc = sanadDesc + date

            descs.append(finalDesc)

        elif filedata[filedata.index('Summary') + 3] == '30020596395':
            path = "C:\\Users\\LENOVO\\Downloads\\11"
            filename = filedata[filedata.index('Class') + 3][5:12]
            filename = filename.replace(filename[filename.index("/")], "-") + ".pdf"

            try:
                destination = os.path.join(path, filename)
                shutil.move(file, destination)
            except shutil.SameFileError:
                pass

            ws = wb['11B']

            indexing()

            date = filedata[filedata.index('Class') + 2][7:17] + " الى " + filedata[filedata.index('Class') + 3][5:15]
            sanadDesc = "تم سداد شقة 11 من تاريخ "
            finalDesc = sanadDesc + date

            descs.append(finalDesc)

        elif filedata[filedata.index('Summary') + 3] == '30013775583':
            path = "C:\\Users\\LENOVO\\Downloads\\27"
            filename = filedata[filedata.index('Class') + 3][5:12]
            filename = filename.replace(filename[filename.index("/")], "-") + ".pdf"

            try:
                destination = os.path.join(path, filename)
                shutil.move(file, destination)
            except shutil.SameFileError:
                pass

            ws = wb['27']

            indexing()

            date = filedata[filedata.index('Class') + 2][7:17] + " الى " + filedata[filedata.index('Class') + 3][5:15]
            sanadDesc = "تم سداد شقة 27 من تاريخ "
            finalDesc = sanadDesc + date

            descs.append(finalDesc)

        elif filedata[filedata.index('Summary') + 3] == '10054624906':
            path = "C:\\Users\\LENOVO\\Downloads\\34"
            filename = filedata[filedata.index('Class') + 3][5:12]
            filename = filename.replace(filename[filename.index("/")], "-") + ".pdf"

            try:
                destination = os.path.join(path, filename)
                shutil.move(file, destination)
            except shutil.SameFileError:
                pass

            ws = wb['2932-34']

            indexing()

            date = filedata[filedata.index('Class') + 2][7:17] + " الى " + filedata[filedata.index('Class') + 3][5:15]
            sanadDesc = "تم سداد شقة 34 من تاريخ "
            finalDesc = sanadDesc + date

            descs.append(finalDesc)

        elif filedata[filedata.index('Summary') + 3] == '10054717729':
            path = "C:\\Users\\LENOVO\\Downloads\\44"
            filename = filedata[filedata.index('Class') + 3][5:12]
            filename = filename.replace(filename[filename.index("/")], "-") + ".pdf"

            try:
                destination = os.path.join(path, filename)
                shutil.move(file, destination)
            except shutil.SameFileError:
                pass

            ws = wb['44']

            indexing()

            date = filedata[filedata.index('Class') + 2][7:17] + " الى " + filedata[filedata.index('Class') + 3][5:15]
            sanadDesc = "تم سداد شقة 44 من تاريخ "
            finalDesc = sanadDesc + date

            descs.append(finalDesc)

        else:
            pass

    wb.save("EXCEL.xlsx")
    frame2 = Frame(window)
    data_string3 = StringVar()
    data_string3.set(descs)
    label = Entry(frame2, textvariable=data_string3).pack()


window = Tk()
window.title("كهرباء شركة المنازل")
image = PhotoImage(file = "Logo.png")
window.iconphoto(True, image)
window.config(background= "#6A172B")
window.geometry("700x700")

frameMain = Frame(window).pack()

data_string = StringVar()
data_string2 = StringVar()
data_string.set("30011715675 - 30020596395 - 30013775583")
entry = Entry(frameMain, state= "readonly", textvariable= data_string, width= 37, font = ("Arial", 20), bd= 0, fg= "black").pack()
data_string2.set("30010522932 - 30010522923 - 30029501703")
entry2 = Entry(frameMain, state= "readonly", textvariable= data_string2, width= 37, font = ("Arial", 20), bd= 0, fg= "black").pack()

button1 = Button(frameMain, text= "Attach files", background= "white", font= ("Arial", 20), command= filesget).pack(pady= 20)
button2 = Button(frameMain, text= "Submit", background= "white", font= ("Arial", 20), command= submit).pack()

window.mainloop()

