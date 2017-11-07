from tkinter import Tk
from tkinter import *
from tkinter import filedialog

root = Tk()

def get_filename():
    filename =  filedialog.askopenfilename(initialdir = "E:/Images",title = "choose your file",filetypes = (("xml files","*.xml"),("all files","*.*")))
    print(filename)

b = Button(root, text='Wybierz plik .xml',command=get_filename,padx=30,pady=15, compound=CENTER)
b.pack()

def main():
    root.mainloop()
    #root.withdraw()

main()