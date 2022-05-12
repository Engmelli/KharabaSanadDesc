import shutil
from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook
import os
from tika import parser
import pyperclip

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
            global FreeCells
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


        date = filedata[filedata.index('Class') + 2][7:17] + " الى " + filedata[filedata.index('Class') + 3][5:15]
        accountdesc = filedata[filedata.index('Summary') + 3]
        readingdesc = " بقراءة " + filedata[filedata.index('Capacity:') + 5][:-16] + " "
        averagecons = int(filedata[filedata.index('Quantity:') + 1][:-13]) / int(filedata[filedata.index('Class') + 5][6:8])
        average = "ومعدل استهلاك " + str(int(averagecons)) + " "
        extratext = "في اليوم من تاريخ "

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

            sanadDesc = "سداد شقة 42 حساب رقم "
            finalDesc = sanadDesc + accountdesc + readingdesc + average + extratext + date

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

            sanadDesc = "سداد شقة 5 حساب رقم "
            finalDesc = sanadDesc + accountdesc + readingdesc + average + extratext + date

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

            sanadDesc = "سداد شقة 11 حساب رقم "
            finalDesc = sanadDesc + accountdesc + readingdesc + average + extratext + date

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

            sanadDesc = "سداد شقة 27 حساب رقم "
            finalDesc = sanadDesc + accountdesc + readingdesc + average + extratext + date

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

            sanadDesc = "سداد شقة 34 حساب رقم "
            finalDesc = sanadDesc + accountdesc + readingdesc + average + extratext + date

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

            sanadDesc = "سداد شقة 44 حساب رقم "
            finalDesc = sanadDesc + accountdesc + readingdesc + average + extratext + date

            descs.append(finalDesc)

        else:
            pass

    wb.save("EXCEL.xlsx")

    frame2 = Frame(window).pack()
    data_string3 = StringVar()
    data_string3.set(descs)
    listbox = Listbox(frame2, listvariable=data_string3, font = ("Arial", 10), bd= 0, width= 100, fg= "black", borderwidth=6, relief="solid", justify= CENTER)
    listbox.pack(pady=10)

    def printt():
        currentselect = listbox.get(ANCHOR)
        pyperclip.copy(currentselect)

    button3 = Button(frame2, command=printt, text="Copy", font = ("Arial", 20), bd= 0, fg= "black").pack()

    button1.config(state=DISABLED)
    button2.config(state=DISABLED)








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

button1 = Button(frameMain, text= "Attach files", background= "white", font= ("Arial", 20), command= filesget)
button1.pack(pady= 20)
button2 = Button(frameMain, text= "Submit", background= "white", font= ("Arial", 20), command= submit)
button2.pack()





window.mainloop()

