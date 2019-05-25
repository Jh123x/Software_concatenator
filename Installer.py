import os
from subprocess import Popen,call
from time import sleep
import tkinter
import ctypes 
from sys import executable


#Creating the GUI
class Window(tkinter.Frame):
    def __init__(self, master=None):
        tkinter.Frame.__init__(self, master)               
        self.master = master

        #Fixing the frame size
        self.master.geometry('500x300')

        #Title of the widget
        self.master.title("Installer")

        #Creating a textbox:
        self.main_tvar = tkinter.StringVar(self.master)
        self.main_tvar.set("Welcome to installer")
        self.maintext = tkinter.Label(self.master,textvariable = self.main_tvar)

        #Creating the install all button
        self.all_btn = tkinter.Button(self.master,text = "Install all in Common Folder", command = lambda: c())
        def c():
            if install_all_c():
                self.main_tvar.set('Successfully installed items in common')

        #Creating the instructions for the tool
        t = 'This tool automatically installs all exe files directly within the Common folder\nIf the exe file is inside another folder please create a shortcut to the folder\n files in Softwares will automatically appear as independent buttons to install'
        self.instructions = tkinter.Label(self.master,text=t)

        #Placing all non-loop items in a grid
        self.instructions.grid(row=10,column = 1, columnspan = 10)
        self.all_btn.grid(row=2,column = 1,columnspan=2)
        self.maintext.grid(row=1,column = 1,columnspan = 2)
        

        #Creating the buttons for the files within the Softwares file
        self.btns = []
        files = get_exe("Softwares",False)
        path = os.getcwd()+'\\' + "Softwares" + '\\'
        if len(files) > 0:
            for i in range(len(files)):
                temp = path + files[i]
                self.btns.append(tkinter.Button(self.master,text = files[i] ,command = lambda: open_file(temp)))
                self.btns[i].grid(row=3+i,column = 1,columnspan=2)


        #Creating the Labels for the files within the Common file
        self.c_lbls = []
        files2 = get_exe("Common",False)
        names_tvar_t = tkinter.StringVar(self.master)
        names_tvar="List of items in common folder\n"
        
        for i in range(len(files2)):
            names_tvar += files2[i]+'\n'
            
        if len(files2) > 0:
            names_tvar_t.set(names_tvar)
        else:
            names_tvar_t.set("There is nothing in the common folder")

            
        self.c_lbl = tkinter.Label(self.master, height = len(files2)+2, width = 30, textvariable = names_tvar_t)
        self.c_lbl.grid(row=2,column = 8, rowspan = len(files2)+1)
        

def install_all_c():
    '''Install all the exe files and shortcut to exe files in the common folder'''
    files = get_exe('Common')
    for i in files:
        open_file(i)
    return True
        


#Retrieving the files
def get_exe(folder,full=True):
    '''To obtain all the exe files within the folder'''
    path = os.getcwd()+'\\' + folder + '\\'
    filenames = list(os.listdir(path))
    files = []

    for i in filenames:
        if i[-4:] == '.exe' or i[-8:] == '.exe.lnk':
            if full:
                files.append(path + i)
            else:
                files.append(i)
    return files



def open_file(file_path):
    '''Opens a file and function will keep running until the installation is over
        returns True once the file is stopped
    '''
    if file_path[-4:] == '.exe':
        proc = Popen(file_path)
    elif file_path[-4:] == '.lnk':
        proc = Popen(file_path,shell=True)
    while proc.poll() is None:
        sleep(2)
    return True




def main():
    #Function to check if there is admin permission
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    if is_admin():

        #Checking if the relavant folders exits:
        path = os.getcwd()
        if not os.path.isdir(path + "\\Softwares\\"):
            os.mkdir(path + "\\Softwares\\")
        if not os.path.isdir(path + "\\Common\\"):
            os.mkdir(path + "\\Common\\")

        
        #Loading the tkinter window
        root = tkinter.Tk()
        root.resizable(True, True) 
        app = Window(root)
        root.mainloop()
        
    else:
        # Re-run the program with admin rights
        ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, __file__, None, 1)
    


if __name__ == '__main__':
    main()
