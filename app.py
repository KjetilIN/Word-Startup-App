from tkinter import *
from tkinter.font import Font
from tkinter.simpledialog import *
from tkinter.filedialog import *
import math
from docx import Document
from os import *





# classes
class mainWindow():
    def setup():
        # basick informasjon
        root = Tk()
        root.title('Word Startup')
        root.geometry('600x200')
        root.resizable(False, False) # making the window not resizable 
        root.configure(bg='deep sky blue')

        # FONTS
        font1 = Font(family='Oswald ExtraLight', size = 24, weight= 'bold')
        font2 = Font(family='Inconsolata UltraCondensed Extr', size = 24)


        # defining labels
        wordLB = Label(root, text="Word Startup", font= font1, bg='deep sky blue', fg='white')
        filepathLB = Label(root, text="Document placement:", font=font2, bg='deep sky blue', fg='white')
        documentnameLb = Label(root, text = "Name of file:", font=font2, bg='deep sky blue', fg='white')
        nameET= Entry(root, width=50)

        filepathSTR = Entry(root, width= 50)

        #functions  
        def getPath():
            folder = tkinter.filedialog.askdirectory() # makes the user open folder placement
            filepathSTR.insert(0, folder) # setting the path in the entry


        def createdoc():
            # document changes
            path = filepathSTR.get()
            nameofDoc = nameET.get() +".docx" # naming the document in the folder
            pathfile = os.path.join(path, nameofDoc) # path and name of the file to save
            document = Document() # makes an document
            document.add_heading(nameET.get(), 0) # adding heading with the name
            document.save(pathfile) # saving the document

            # clearing all entry
            filepathSTR.delete(0, END)
            nameET.delete(0, END)

            # toplevel
            top = Toplevel()
            top.configure(bg='green')
            #success message to let the user know that the file was created correctly
            suLB= Label(top, text="Success!", bg='green', fg='white', font=font1)
            suLB.grid(column=0, row=0)
            

            
           
            


        # important labels that need function
        
        filepathBT = Button(root, text="File placement", command= getPath)
        crtBT = Button(root, text="Create document", command=createdoc)

        # display all

        wordLB.grid(columnspan=True, row=0)
        filepathLB.grid(column=0, row=1)
        filepathSTR.grid(column=1, row=1)
        filepathBT.grid(column=2, row=1, padx=10)
        documentnameLb.grid(column=0, row=2)
        nameET.grid(column=1, row=2)
        crtBT.grid(columnspan=True, row =3)




        # loop
        root.mainloop()



if __name__ == "__main__":
    mainWindow.setup()