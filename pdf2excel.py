import Tkinter, Tkconstants, tkFileDialog

class ParserPdfDialog(Tkinter.Frame):

  def __init__(self, root):

    Tkinter.Frame.__init__(self, root)

    # options for buttons
    button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}
    self.FolderName = Tkinter.StringVar()
    # define buttons
    Tkinter.Button(self, text='Select the PDF Folder', command=self.askdirectory).pack(**button_opt)
    Tkinter.Label(self, textvariable=self.FolderName).pack()
    self.FolderName.set("Please Choose a directory")
    #Tkinter.Button(self, text='Export Excel', )

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

if __name__=='__main__':
  root = Tkinter.Tk()
  ParserPdfDialog(root).pack()
  root.mainloop()