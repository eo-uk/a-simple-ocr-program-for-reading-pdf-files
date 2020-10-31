#An OCR tool to convert PDF Files into plain text#

import os
import threading
import traceback
import configparser
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import *

from ocr import PDF2Text


class Program():
    '''TkInter GUI for PDF OCR'''
    
    def __init__(self, settings_file='settings.ini'):
        '''Sets up the root window and starts the mainloop'''
        #Read settings.ini
        self.settings_file = settings_file
        self.read_settings(self.settings_file)
        #Create root window
        self.root = Tk()
        self.root.title("PDF 2 TEXT OCR")
        self.root.minsize(600, 250)
        self.root.resizable(True, False)
        #Define callback function for error messages
        self.root.report_callback_exception = self.display_error
        #Create string and int variables
        self.setup_variables()
        #Create section for picking tesseract.exe
        self.create_section(
            section = 'exe',
            text = "Location of 'tesseract.exe':",
        )
        #Create section for picking the source file
        self.create_section(
            section = 'src',
            text = "PDF file to convert:",
        )
        #Create section for picking the output file
        self.create_section(
            section = 'out',
            text = "Output File:",
        )
        #Create label for status display
        Label(
            self.root,
            textvariable = self.status
        ).pack()
        #Create checkbox for pre-processing of images for OCR
        Checkbutton(
            self.root,
            text = 'Pre-Process (Slower)',
            variable = self.prep
        ).pack(
            pady = 10
        )
        #Create button to convert source file
        self.btn_convert = Button(
            self.root,
            text = "Convert",
            command = self.start_conversion
        )
        self.btn_convert.pack(
            pady = (10, 30)
        )
        #Start the main loop
        mainloop()

    def read_settings(self, path):
        '''
        Loads the settings file.

        path (str) : Path to the settings file
        '''
        #Read settings file
        config = configparser.ConfigParser()
        config.read('./' + path)
        #Load values into self.settings dict
        settings = {}
        for section in config.sections():
            settings[section] = {}
            for key, value in config[section].items():
                settings[section][key] = value
        setattr(self, 'settings', settings)

    def update_settings(self, section, key, value):
        '''
        Updates runtime settings and the settings file

        section (str) : settings section
        key (str) : key under the section
        value (str) : new value
        '''
        #Set the value if section and key exist
        try:
            self.settings[section][key] = value
        #Create section if it doesn't exist
        except KeyError:
            self.settings[section] = {key : value}
        #Update the settings.ini
        self.dump_settings(self.settings_file)

    def dump_settings(self, path):
        '''
        Dumps settings to the settings file

        path (str) : Path to the settings file
        '''
        #Create and populate the configparser
        config = configparser.ConfigParser()
        for section, content in self.settings.items():
            config[section] = content
        #Save to file
        with open(path, 'w') as file:
            config.write(file)

    def create_section(self, section, text):
        '''
        Creates a section for picking a file.

        file (str) : Which file is to be picked (exe, src or out)
        text (str) : Text to display at the top of the section
        '''
        #Frame to contain file pickers
        frame = Frame(
            self.root
        )
        frame.pack(
            pady = 10,
            padx = 30,
            fill = BOTH
        )
        #Label to display at the top
        Label(
            frame,
            text = text
        ).pack(
            side = TOP
        )
        #Entry box to pick file
        Entry(
            frame,
            textvariable = getattr(self, section),
            width = 50,
            state = (lambda section:
                     'readonly' if (section == 'exe') else NORMAL)(section)
        ).pack(
            side = LEFT,
            fill = X,
            expand = 1
        )
        #Button to pick file
        Button(
            frame,
            text = "Select File",
            command = lambda: self.select_file(section)
        ).pack(
            padx = 10
        )

    def setup_variables(self):
        '''Sets up str and int variables needed'''
        #tesseract.exe path from the settings file
        try:
            self.exe = StringVar(value=self.settings['PATHS']['exe_path'])
        except KeyError:
            self.exe = StringVar()
        #Set a callback to update settings file if var changes
        self.exe.trace(
            mode = 'w',
            callback = lambda *args: self.update_settings(
                section = 'PATHS',
                key = 'exe_path',
                value = self.exe.get()
            )
        )
        self.src = StringVar() #source file path
        self.out = StringVar() #output file path
        self.prep = IntVar() #for the preprocessing checkbox
        self.status = StringVar() #for status display

    def select_file(self, file):
        '''
        Prompts user to pick a file.

        file (str) : exe, src or out, depending on which file is to be picked
        '''
        #Filetypes allowed based on file
        FILETYPES = {
            'exe': (
                ('EXE file', '*.exe'),
            ),
            'src': (
                ('PDF file', '*.pdf'),
            ),
            'out': (
                ('Plain Text', '*.txt'),
            )
        }
        #Set file StringVar to the path of file picked
        if file == 'out': #Use saveasfilename for output file
            path = filedialog.asksaveasfilename(
                initialdir = os.getcwd(),
                defaultextension = ".txt",
                filetypes = FILETYPES[file]
            )
        else: #Use askopenfilename for the others
            path = filedialog.askopenfilename(
                initialdir = os.getcwd(),
                filetypes = FILETYPES[file]
            )
        if path: #If a path is chosen
            getattr(self, file).set(path)
            
    def start_conversion(self):
        '''Starts the conversion of source file into text.'''
        #Validate exe path
        if os.path.basename(self.exe.get()) != 'tesseract.exe':
            messagebox.showwarning(
                title = "Invalid EXE Selected",
                message = "Please select a valid tesseract.exe file."
            )
        #Validate other paths
        elif not self.src.get() or not self.exe.get() or not self.out.get():
            messagebox.showwarning(
                title = "File Not Selected",
                message = "Please select all files to proceed."
            )
        #If source file, output file and tesseract.exe are all selected
        else:
            self.thread = threading.Thread(
                target = self.convert_file,
                daemon = True
            )
            self.thread.start()

    def convert_file(self):
        '''Converts the source file into text and saves it.'''
        try:
            #Disable the convert button
            self.btn_convert.config(state = DISABLED)
            #Notify user that conversion has started
            self.status.set("Converting...")
            #Initialize PDF2Text object with the exe path
            p2t = PDF2Text(self.exe.get())
            #Read source file and save it
            p2t.open(self.src.get(), self.prep.get())
            p2t.save(self.out.get())
            #Notify the user that the process is complete
            self.status.set("Done!")
        except Exception as e:
            self.display_error(type(e), e, e.__traceback__)
        finally:
            #Enable the convert button
            self.btn_convert.config(state = NORMAL)

    def display_error(self, *args):
        '''Displays errors in a messagebox.'''
        #Format exceptions, display in messagebox
        error = traceback.format_exception(*args)
        messagebox.showerror(
            title = 'Error',
            message = 'An Error Has Occured: \n\n' + '\n'.join(error)
        )

def main():
    Program()


if __name__ == "__main__":
    main()
