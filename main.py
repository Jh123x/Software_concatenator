import os
from subprocess import Popen,call
import tkinter
from tkinter import scrolledtext as tst
import ctypes 
from sys import executable


#Creating the GUI
class Window(tkinter.Frame):
    def __init__(self, master=None):
        '''Initialise the variables'''
        tkinter.Frame.__init__(self, master=None)               
        self.master = master
        self.items_to_run = []

        #Creating a main textbox
        self.maintext = tkinter.Label(self.master, text = "Welcome to AppR")

        #Fixing the frame size
        self.master.geometry('700x450')

        #Title of the widget
        self.master.title("AppR")

        #Creating a log box
        self.log = tst.ScrolledText(self.master, wrap = tkinter.WORD, width = 70, height = 7, borderwidth = 2)

        #Getting all the items in the folders
        self.s_files = get_exe('Softwares')
        self.s_files_n =  get_exe('Softwares',False)
        self.c_files = get_exe('Common')
        self.c_files_n = get_exe('Common',False)
        self.all_files = tuple(self.s_files + self.c_files)

        #Creating an install button
        self.install_btn = tkinter.Button(self.master,text = 'Install' , command = lambda : install_selected(self))

        #Creating the instructions for the tool
        t = 'This tool automatically installs all exe files directly within the Common folder\nIf the exe file is inside another folder please create a shortcut to the folder\n files in Softwares will automatically appear as independent Checkbox'
        self.instructions = tkinter.Label(self.master,text=t)

        #Placing all items in a grid
        self.maintext.grid(row = 0, column = 1)
        self.install_btn.grid(row = 9,column = 10, columnspan = 1)
        self.log.grid(row = 13, column = 1, columnspan = 7, rowspan = 5)

        #Creating a menu of checkboxes for user to indicate what to install
        self.checkboxes = []
        self.c_len = len(self.c_files)
        self.s_len = len(self.s_files)


        #Creating a menu of checkboxes for items in common folder
        for f in range(self.c_len):
            fname = self.c_files_n[f]
            file_c = self.c_files[f]
            self.checkboxes.append(tkinter.Checkbutton(self.master,text = fname,command = create_software_func(add_to_selected,file_c,self),anchor = 'w'))
            add_to_selected(file_c,self)
            self.checkboxes[f].toggle()
            self.checkboxes[f].grid(row = 3+f,column = 1,columnspan = 1,sticky="W")

        #Creating a menu of checkboxes for items in Software folder
        for i in range(self.s_len):
            fname = self.s_files_n[i]
            file_s = self.s_files[i]
            self.checkboxes.append(tkinter.Checkbutton(self.master,text = fname, command = create_software_func(add_to_selected,file_s,self),anchor = 'w'))
            self.checkboxes[self.c_len + i].grid(row = (3 + self.c_len + i),column = 1,columnspan = 1,sticky="W")

        #Placing the instructions item at the bottom of the page
        self.instructions.grid(row=(len(self.all_files) + 40),column = 1, columnspan = 3)

            

def create_software_func(function,*arg):
    '''A Function to create a copy of a function'''

    def helper():
        '''A Helper function for tkinter commands'''
        function(*arg)
        return

    return helper

def log(phrase,app):
    '''Function to Log activities'''

    #Insert the words into the scrolled text
    app.log.insert(tkinter.END,phrase+'\n')

    #Makes the scrolled text move to the end
    app.log.see(tkinter.END)
    return 
        
def install_selected(app):
    '''Function to install the files which are selected'''
    files = app.all_files
    for i in app.items_to_run:
        log('Running ' + files[i],app)
        open_file(files[i],app)
    log('All Done!',app)
    return
    
def add_to_selected(item,app):
    '''Function to add/remove the apps selected'''
    files = app.all_files
    it = files.index(item)
    if it not in app.items_to_run:
        app.items_to_run.append(it)
        log('Added ' + files[it],app)
        return True
    else:
        app.items_to_run.remove(it)
        log('Removed ' + files[it] ,app)
        return False

#Retrieving the files
def get_exe(folder,full=True):
    '''To obtain all the exe files within the folder'''

    #Obtaining Absolute path to the folder
    path = os.getcwd()+'\\' + folder + '\\'
    filenames = list(os.listdir(path))
    files = []

    #Getting all .exe(Applications) .lnk (Shortcuts) within a specified folder
    for i in filenames:
        if i[-4:] == '.exe' or i[-8:] == '.exe.lnk':
            if full:
                files.append(path + i)
            else:
                files.append(i)
    return files

def open_file(file_path,app):
    '''Opens a file and function will keep running until the window is closed
        returns True once the file is stopped
    '''
    #Log the activity
    log(file_path,app)

    #Opens the path based on whether the file is an application file or a shortcut file
    if file_path[-4:] == '.exe':
        proc = Popen(file_path)
    elif file_path[-4:] == '.lnk':
        proc = Popen(file_path,shell=True)

    #Wait till the shortcut/application
    while proc.poll() is None:
        pass
    return True

def is_admin():
    '''A Function to check if the app is running as administrator'''
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def main():
    '''Main function'''
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
        app.mainloop()

    else:
        #Re-run the program with admin rights
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(None, 'runas',
            '"' + executable + '"',
            '"' + __file__ + '"',
            None, 1)


if __name__ == "__main__":
    main()
