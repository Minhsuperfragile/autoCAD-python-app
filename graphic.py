from tkinter import *
from tkinter.filedialog import askopenfilename
from process import Utils

inputfilepath1:str
inputfilepath2:str

root = Tk()
root.resizable(False,False)
root.geometry('240x144')
root.title("xldl")

#create a label (line of text)
mainLabel = Label(root, text="Xu ly du lieu")
#display label on the screen
mainLabel.pack()

def triggerT2Xscreen():
    top = Toplevel()
    top.resizable(False,False)
    top.geometry("300x130")
    top.title("convert txt to xlsx")
    
    separatorLabel = Label(top,text="Separator: ",pady=5,padx=5).grid(row=0,column=0)
    separatorEntry = Entry(top,width=20)
    separatorEntry.grid(row=0, column=1)

    inputFileLabel = Label(top,text="Input File: ",pady=5,padx=5).grid(row=1,column=0)
    selectedFileLabel = Label(top,text="Not Selected", pady=5,padx=5)
    selectedFileLabel.grid(row=1,column=1)

    outputFileLabel = Label(top,text="Output to: ",padx=5,pady=5).grid(row=2,column=0)
    outputFileNameLabel = Label(top,text="Same name as input file",padx=5,pady=5)
    outputFileNameLabel.grid(row=2,column=1)

    def selectInputFile():
        global inputfilepath1 
        inputfilepath1 = askopenfilename(filetypes=[("Text files","*.txt")])
        inputfilename = inputfilepath1.split('/')[-1]
        selectedFileLabel.config(text=inputfilename)
        outputFileNameLabel.config(text=inputfilename.replace("txt","xlsx"))

    selectFileButton = Button(top, text="Browse",command=selectInputFile,padx=8,pady=5,width=8).grid(row=1,column=2)

    def startFunction():
        Utils.txt2xlsx(inputfilepath1,inputfilepath1.split('/')[-1].replace("txt","xlsx"),separatorEntry.get())
        return
    startConversionButton = Button(top,text="Start",command=startFunction,padx=8,pady=5,width=8).grid(row=2,column=2)

def triggerX2Tscreen():
    top = Toplevel()
    top.resizable(False,False)
    top.geometry("300x130")
    top.title("convert xlsx to txt")

    inputFileLabel = Label(top,text="Input File: ",pady=5,padx=5).grid(row=0,column=0)
    selectedFileLabel = Label(top,text="Not Selected", pady=5,padx=5, width=18)
    selectedFileLabel.grid(row=0,column=1)

    outputFileLabel = Label(top,text="Output to: ",padx=5,pady=5).grid(row=1,column=0)
    outputFileNameLabel = Label(top,text="Same name as input file",padx=5,pady=5,width=18)
    outputFileNameLabel.grid(row=1,column=1)

    def selectInputFile():
        global inputfilepath2
        inputfilepath2 = askopenfilename(filetypes=[("Excel files","*.xlsx")])
        inputfilename = inputfilepath2.split('/')[-1]
        selectedFileLabel.config(text=inputfilename)
        outputFileNameLabel.config(text=inputfilename.replace("xlsx","txt"))

    selectFileButton = Button(top, text="Browse",command=selectInputFile,padx=8,pady=5,width=8).grid(row=0,column=2)

    rowLabel = Label(top,text="Rows: ",pady=5,padx=5).grid(row=2,column=0)
    rowEntry = Entry(top,width=20)
    rowEntry.grid(row=2, column=1)

    columnLabel = Label(top,text="Columns: ",pady=5,padx=5).grid(row=3,column=0)
    colEntry = Entry(top,width=20)
    colEntry.grid(row=3, column=1)

    def startFunction():
        Utils.xlsx2txt(inputfilepath2,inputfilepath2.split('/')[-1].replace("xlsx","txt"),int(colEntry.get()),int(rowEntry.get()))
        return
    
    startConversionButton = Button(top,text="Start",command=startFunction,padx=8,pady=5,width=8).grid(row=2,column=2)

#create txt to xlsx button
T2XButton = Button(root, text="TXT to XLSX",command=triggerT2Xscreen)
T2XButton.pack(pady=5)

#create xlsx to txt button
X2TButton = Button(root,text="XLSX to TXT", command=triggerX2Tscreen)
X2TButton.pack(pady=5)

#start screen
root.mainloop()