from tkinter import *
from tkinter import ttk, font
from tkinter.filedialog import askopenfilename, askopenfilenames
from tkinter.messagebox import showerror
from tkinter.ttk import Style


class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("Навантаження кафедри")
        self.master.rowconfigure(5, weight=1)
        self.master.columnconfigure(5, weight=1)
        self.grid(sticky=W + E + N + S)
        self.master.geometry("800x600")
        self.style = Style()
        self.style.configure('TButton', font=('calibri', 15, 'normal'), borderwidth='4', padding=4, background="#ccc")
        self.fName = StringVar()
        # button 1

        self.button = ttk.Button(self, text="Відкрити", command=self.load_file, width=10,  style='TButton')
        self.button.grid(row=1, column=0, columnspan=3, padx=50, pady=50)
        self.label_Files = ttk.Label(self, textvariable=self.fName, width=50, style='TButton')
        self.label_Files.grid(row=2, column=0, columnspan=3, padx=50, pady=50)

    def load_file(self):
        files_name = askopenfilenames(filetypes=(("2003 Excel files", "*.xls"),
                                               ("2007+ Excel files", "*.xlsx")))

        if files_name:
            try:
                self.fName.set(files_name)
            except FileNotFoundError:
                showerror("Open Source File", "Failed to read file\n'%s'" % files_name)
            return


if __name__ == "__main__":
    MyFrame().mainloop()
