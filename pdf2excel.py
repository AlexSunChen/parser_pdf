# -*- coding: UTF-8 -*-

import Tkinter, Tkconstants, tkFileDialog
import xlwt
import os

class ParserPdfProcess(Tkinter.Frame):

  def __init__(self, root):

    Tkinter.Frame.__init__(self, root)

    # options for buttons
    button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}
    self.FolderName = Tkinter.StringVar()
    # define buttons
    Tkinter.Button(self, text='Select the PDF Folder', command=self.askdirectory).pack(**button_opt)
    Tkinter.Label(self, textvariable=self.FolderName).pack()
    self.FolderName.set("Please Choose a directory")
    Tkinter.Button(self,text="Export Excel", command=self.ExportExcel).pack(**button_opt)

    # defining options for opening a directory
    self.dir_opt = options = {}
    options['initialdir'] = 'C:\\'
    options['mustexist'] = False
    options['parent'] = root
    options['title'] = 'Please choose a PDF directory'

  def askdirectory(self):

    """Returns a selected directoryname."""
    self.FolderName.set(tkFileDialog.askdirectory(**self.dir_opt))
    return

  def ExportExcel(self):
    ParserPdfProcess.PdfToTxt(self)
    lists = ParserPdfProcess.ParserListWord(self)
    dicts = {'橋本': '1213', '長じ': '121', 'カツミ': '3258'}
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('MN', cell_overwrite_ok=True)

    i = 0
    j = 0
    for key, value in dicts.items():
      sheet.write(i, j, key.decode('utf-8'))
      sheet.write(i, j+1, value)
      i += 1
    book.save(self.FolderName.get() + '/export.xls')

    self.FolderName.set("Completely Export the data !")

  def ParserListWord(self):
    fileList = os.listdir(self.FolderName.get())
    os.chdir(self.FolderName.get())
    for file in fileList:
      if ".txt" in file:
        curFile = open(file, mode='r')
        while 1:
          line = curFile.readline()
          line = line.replace(' ','')
          keyword = 'ニフティ株式会社'
          if keyword in line:
            print line
          if not line:
            break
    return

  def PdfToTxt(self):
    fileList = os.listdir(self.FolderName.get())
    os.chdir(self.FolderName.get())
    for file in fileList:
      if ".pdf" in file:
        os.system('pdf2txt.py  -o ' + file[0:-4] + '.txt ' + file)
    return


if __name__=='__main__':
  root = Tkinter.Tk()
  ParserPdfProcess(root).pack()
  root.mainloop()
