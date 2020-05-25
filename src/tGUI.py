from tkinter import *
from tkinter import ttk, font, StringVar
from tkinter.filedialog import askopenfilename, askopenfilenames
from tkinter.messagebox import showerror
from tkinter.ttk import Style

from src.convert import ConverterXLS


class StudyLoadFrame(Frame):

    def __init__(self):
        Frame.__init__(self)
        self.master.title("Навантаження кафедри")
        self.master.rowconfigure(7, weight=1)
        self.master.columnconfigure(6, weight=1)
        self.grid(sticky=W + E + N + S)
        self.master.geometry("800x600")
        self.style = Style()
        self.style.configure('TButton', font=('calibri', 13, 'normal'), borderwidth='4', padding=4, background="#ccc")
        self.combostyle = ttk.Style()

        # ATTENTION: this applies the new style 'combostyle' to all ttk.Combobox
        # self.combostyle.theme_use('combostyle')

        self.fName = StringVar()
        self.fName.set(" ")
        self.default_file_names = []
        self.only_file_name = []
        self.file_names_split = StringVar()
        self.button_open_kafedra = ttk.Button(self, text="Відкрити", command=self.load_file, width=10, style='TButton')
        self.button_open_rnp = ttk.Button(self, text="Відкрити", command=self.load_file2, width=10, )
        self.label_Files = ttk.Label(self, textvariable="Відкрийте декілька файлів для ", width=50)
        self.label_Files = ttk.Label(self, textvariable=self.fName, width=50, style='TButton')

        self.combobox_Files = ttk.Combobox(self, state="readonly", width=30)

        self.button_converter = ttk.Button(self, text="Конвертувати",
                                           command=self.convert_file,
                                           width=15,
                                           style='TButton')
        self.button_converter.grid(row=2, column=0, columnspan=3, padx=50, pady=30)
        self.label_Files.grid(row=3, column=0, columnspan=3, padx=30, pady=30)
        self.button_open_kafedra.grid(row=1, column=1, columnspan=3, padx=30, pady=30)
        self.combobox_Files.grid(row=1, column=0, padx=30)

    def load_file(self):
        files_name = askopenfilenames(filetypes=(("2003 Excel files", "*.xls"),
                                                 ("2007+ Excel files", "*.xlsx")))

        if files_name:
            try:
                self.only_file_name = []
                temporary_array = []
                for item in range(0, len(files_name)):
                    temporary_array.append(files_name[item].replace('/', '\\', 10))

                files_name = temporary_array
                self.default_file_names = files_name

                for item in range(0, len(files_name)):
                    temporary_array2 = files_name[item].split('\\')
                    self.only_file_name.append(temporary_array2[-1])
                self.combobox_Files['values'] = self.only_file_name
                self.fName.set(self.only_file_name)
                self.combobox_Files.current(0)

            except FileNotFoundError:
                showerror("Open Source File", "Failed to read file\n'%s'" % files_name)
            return

    def convert_file(self):
        try:
            ConverterXLS(self.default_file_names[self.combobox_Files.current()])
            self.default_file_names[self.combobox_Files.current()] += 'x'

            temp_array = list(self.combobox_Files['values'])
            print(temp_array)
            temp_array[temp_array.index(self.combobox_Files.get())] += 'x'
            self.combobox_Files.set(self.combobox_Files.get() + 'x')
            self.combobox_Files['values'] = temp_array
            print(self.combobox_Files['values'])
        except IndexError:
            showerror("Open Source File", "Failed to read file")
        return


if __name__ == "__main__":
    StudyLoadFrame().mainloop()
