#---------------------------------------------------------------------#
# Read basic Excel table of part numbers and match to filenames in
# genius attachments folder, copy the files, then save in a new
# directory on the desktop.

# Nicholas Gregg
# 04/28/2020
#---------------------------------------------------------------------#

import os
import xlrd
from shutil import copyfile
import tkinter as tk
from tkinter import Frame, Label, Button, Text, filedialog, Menu


class MainApp(tk.Tk):
    def __init__(self, main_win):
        # Main window
        self.main_win = main_win
        self.main_win.geometry("530x170+500+350")
        self.main_win.resizable(0, 0)
        self.main_win.configure(bg="#303030")
        self.main_win.title("File Transfer App")

        # Menu Bar
        self.menubar = Menu(self.main_win)
        self.main_win.config(menu=self.menubar)

        # File menu
        fileMenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="File", menu=fileMenu)

        fileMenu.add_command(label="New", command=None)
        fileMenu.add_separator()
        fileMenu.add_command(label="Exit", command=None)

        # Settings menu
        settingsMenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Settings", menu=settingsMenu)

        settingsMenu.add_command(label="Output Location", command=None)

        # Help menu
        self.menubar.add_command(label="Help", command=self.help_menu)

        # Labels
        self.fin_label = Label(self.main_win, text="Input File", bg="#303030",
                               fg="white")
        self.fin_label.place(x=15, y=15)
        self.fout_label = Label(self.main_win, text="Output Folder", bg="#303030",
                                fg="white")
        self.fout_label.place(x=15, y=70)

        # Text fields
        self.fin_text = Text(self.main_win, width=52, height=1,
                             bg="#353535", fg="#909090")
        self.fin_text.insert(1.0, "Select a file")
        self.fin_text.place(x=15, y=40)
        self.fout_text = Text(self.main_win, width=52, height=1,
                              bg="#353535", fg="#909090")
        self.fout_text.insert(1.0, "Select a folder")
        self.fout_text.place(x=15, y=90)

        # Buttons
        self.fin_button = Button(self.main_win, command=self.fin_click, width=8, text="Browse", bg="#353535",
                                 fg="white", activebackground="#404040", activeforeground="#808080")
        self.fin_button.place(x=450, y=38)
        self.fout_button = Button(self.main_win, command=self.fout_click, width=8, text="Browse", bg="#353535",
                                  fg="white", activebackground="#404040", activeforeground="#808080")
        self.fout_button.place(x=450, y=88)
        self.run_button = Button(self.main_win, command=self.run_click, width=8, text="Run", bg="#353535",
                                 fg="white", activebackground="#404040", activeforeground="#808080")
        # self.run_button.place(x=200, y=130)
        self.run_button.pack(side="bottom", pady=10)

    # File search when input file browse button is clicked

    def fin_click(self):
        self.fin_clicked = filedialog.askopenfilename()
        self.fin_text.delete(1.0, 2.0)
        self.fin_text.insert(1.0, self.fin_clicked)

    def fout_click(self):
        self.fout_clicked = filedialog.askdirectory()
        self.fout_text.delete(1.0, 2.0)
        self.fout_text.insert(1.0, self.fout_clicked)

    def run_click(self):
        self.fin_path = self.fin_text.get(1.0, 2.0)
        self.fout_path = self.fout_text.get(1.0, 2.0)

        # Create list from part number column
        workbook = xlrd.open_workbook(self.fin_path.rstrip())
        worksheet = workbook.sheet_by_index(0)
        part_numbers = worksheet.col_values(0)

        # Add desired file extension.
        # TODO: Add selectable filetypes.
        for i in part_numbers:
            part_numbers[part_numbers.index(i)] += ".PDF"

        # Search folder for matching names.
        # If match found, copy file to new folder on desktop.
        # TODO: Add settings menu to allow change for source directory.
        src = "\\\\andros-dc\\groups\\ENG DATA\\ENG_Directory\\Attachments_Genius\\"
        dst = self.fout_path.rstrip() + "\\"

        for fname in os.listdir(src):
            for partno in part_numbers:
                if fname.upper() == partno.upper():
                    copyfile(src + fname, dst + fname)

    # Application instructions pop-up window.
    def help_menu(self):
        pass


# Run program

def main():
    root = tk.Tk()
    MainApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
