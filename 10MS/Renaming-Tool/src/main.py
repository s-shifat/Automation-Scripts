import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askdirectory
import os
from shutil import rmtree

# ১ ২ ৩ ৪ ৫ ৬ ৭ ৮ ৯ ০
BAN_DIGIT = {ban: eng for ban, eng in zip([*'১২৩৪৫৬৭৮৯০'], [*'1234567890'])}


def get_engDigit(bangla_digit, to_number=False):
    translation = ''.join(list(map(lambda x: BAN_DIGIT.get(x, x), bangla_digit)))
    if to_number:
        try:
            return int(translation)
        except ValueError:
            return float(translation)
    return translation


class Application(tk.Frame):
    def __init__(self, master):
        super(Application, self).__init__()
        self.master = master

        # self.table.field_names = ['Previous Name', 'New Name']
        # input folder name area
        self.folderInputPath = tk.StringVar()
        self.label = tk.Label(self.master, text='Input Folder Location:')
        self.label.grid(row=0, column=0, sticky='nsew', padx=10, pady=20)
        self.inputFolderLabel = tk.Text(width=30, height=1, padx=5)
        self.inputFolderLabel.grid(row=0, column=1, sticky='nsew', padx=10, pady=20, columnspan=2)
        self.inputButton = ttk.Button(text='Select',
                                      command=lambda: self.selectFolder(self.inputFolderLabel, self.inputFileLister,
                                                                        self.inputFolderLabel))
        self.inputButton.grid(row=0, column=3, sticky='nsew', padx=10, pady=20)

        # output folder name area
        self.label1 = tk.Label(self.master, text='Output Folder Location:')
        self.label1.grid(row=1, column=0, sticky='nsew', padx=10, pady=20)
        self.outputFolderLabel = tk.Text(width=30, height=1, padx=5)
        self.outputFolderLabel.grid(row=1, column=1, sticky='nsew', padx=10, pady=20, columnspan=2)
        self.outputButton = ttk.Button(text='Select',
                                       command=lambda: self.selectFolder(self.outputFolderLabel, self.outputFileLister,
                                                                         self.outputFolderLabel))
        self.outputButton.grid(row=1, column=3, sticky='nsew', padx=10, pady=20)

        self.inputFolderPath = self.inputFolderLabel.get('1.0', 'end-1c')
        self.outputFolderPath = self.outputFolderLabel.get('1.0', 'end-1c')

        self.previewLabel = tk.Label(text='Preview')
        self.previewLabel.grid(row=2, column=0, columnspan=4)

        # input files listing
        self.inputFileListingFrame = tk.Frame()
        self.inputFileListHeader = tk.Label(self.inputFileListingFrame, text='Input Folder')
        self.inputFileListHeader.pack()
        self.inputFileLister = tk.Listbox(self.inputFileListingFrame, width=40, height=20)

        self.inputFileLister.pack()
        self.inputFileListingFrame.grid(row=3, column=0, columnspan=2, sticky='nsew', padx=10, pady=20)

        # output files listing
        self.outputFileListingFrame = tk.Frame()
        self.outputFileListHeader = tk.Label(self.outputFileListingFrame, text='Output Folder')
        self.outputFileListHeader.pack()
        self.outputFileLister = tk.Listbox(self.outputFileListingFrame, width=40, height=20)

        self.outputFileLister.pack()
        self.outputFileListingFrame.grid(row=3, column=2, columnspan=2, sticky='nsew', padx=5, pady=20)

        # execution frame @ row3, col4
        self.executionFrame = tk.Frame()
        self.delete_input_folder = tk.BooleanVar()
        self.delete_input_folder.set(False)
        self.generate_report_file = tk.BooleanVar()
        self.generate_report_file.set(True)
        self.checkButton = ttk.Checkbutton(self.executionFrame, text='Delete Input Folder',
                                           variable=self.delete_input_folder)
        self.checkButton.pack()
        self.reportFileButton = ttk.Checkbutton(self.executionFrame, text='Generate a report file',
                                                variable=self.generate_report_file)
        self.reportFileButton.pack()
        self.renameButton = ttk.Button(self.executionFrame, text='Rename All', command=self.rename_all)
        self.renameButton.pack(pady=20)

        self.executionFrame.grid(row=3, column=4)

    def listAdder(self, listObject, path_):
        if os.path.isdir(path_):
            data = [x for x in os.listdir(path_) if os.path.isfile(os.path.join(path_, x))]
            data = self.sortdir(data)
            for i in range(len(data)):
                listObject.insert(i + 1, data[i])
        else:
            self.error_message("Folder doesn't exist")

    def selectFolder(self, variable, listObject, path_):
        variable.delete('1.0', tk.END)
        selection = askdirectory()
        variable.insert(tk.END, selection)
        variable.config(wrap=None)
        # time.sleep(1)
        listObject.delete(0, tk.END)
        self.listAdder(listObject, path_.get('1.0', 'end-1c'))

    def rename_all(self):
        inputPath = self.inputFolderLabel.get('1.0', 'end-1c')
        outputPath = self.outputFolderLabel.get('1.0', 'end-1c')

        self.inputFileList = self.sortdir(os.listdir(inputPath))
        self.outputFileList = self.sortdir(os.listdir(outputPath))
        if len(self.inputFileList) != len(self.outputFileList):
            self.error_message("Number of files doesn't match")
        elif inputPath == 0 or outputPath == 0:
            self.error_message("Folder field empty!")
        elif len(self.inputFileList) == len(self.outputFileList):
            for inputFileName, outputFileName in zip(self.inputFileList, self.outputFileList):
                newName = os.path.join(outputPath, inputFileName)
                oldName = os.path.join(outputPath, outputFileName)
                os.rename(oldName, newName)

            self.finalWindow = self.error_message("Rename operation completed!!")
            button = tk.Button(self.finalWindow, text="Open Output Folder", command=lambda: os.startfile(outputPath),
                               pady=10, padx=10)
            button.pack(pady=20, expand=True)

            if self.delete_input_folder.get():
                rmtree(inputPath, ignore_errors=True)
            if self.generate_report_file.get():
                self.make_report(outputPath, inputPath)

    def error_message(self, message):
        win = tk.Toplevel(width=len(message))
        win.geometry('200x100')
        win.title('Report')
        l = tk.Label(win, text=message)
        # win.config(width=len(message))
        l.pack(expand=True, pady=10)

        return win

    def make_report(self, outputPath, inputPath):
        with open(os.path.join(outputPath, 'report.csv'), 'w', encoding='utf-16') as f:
            f.write(
                f'File names taken from this Folder (input): {inputPath} \n Renamed files are in this Folder: {outputPath}\n')
            f.write("Previous Name, New Name\n")
            for o, n in zip(self.outputFileList, self.inputFileList):
                f.write(f"{o}, {n}\n")

    def sortdir(self, oslistdir):
        try:
            return sorted(oslistdir, key=lambda x: get_engDigit(x.split(')')[0].strip(), to_number=True))
        except ValueError:
            return sorted(oslistdir, key=len)


def main():
    root = tk.Tk()
    root.geometry('680x530')
    root.title('Renaming Tool')
    root.iconbitmap('./icons/robot-160.ico')
    Application(root)
    root.mainloop()


if __name__ == '__main__':
    main()
