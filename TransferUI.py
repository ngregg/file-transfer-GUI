#---------------------------------------------------------------------#
# Read basic Excel table of part numbers and match to filenames in
# genius attachments folder, copy the files, then save in a new
# directory on the desktop.

# Nicholas Gregg
# 04/28/2020
#---------------------------------------------------------------------#

import os
from os.path import basename
import xlrd
from shutil import copyfile
import zipfile
import tkinter as tk
from tkinter import Frame, Label, Button, Text, filedialog, Menu
from tkinter.messagebox import showinfo


class MainApp(tk.Tk):
    def __init__(self, main_win):
        # Main window
        self.main_win = main_win
        self.main_win.geometry("530x190+500+350")
        self.main_win.resizable(0, 0)
        self.main_win.configure(bg="#303030")
        self.main_win.title("PDF Transfer App")

        # Menu Bar
        self.menubar = Menu(self.main_win)
        self.main_win.config(menu=self.menubar)

        # File menu
        file_menu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="File", menu=file_menu)

        file_menu.add_command(label="New", command=self.new)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=main_win.quit)

        # Settings menu
        settings_menu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Settings", menu=settings_menu)

        settings_menu.add_command(
            label="Search Location", command=self.search_dir)

        # Default search directory
        self.src = "\\\\andros-dc\\groups\\ENG DATA\\ENG_Directory\\Attachments_Genius\\"

        # Help menu
        self.menubar.add_command(label="Help", command=self.help_menu)

        # Labels
        self.fin_label = Label(self.main_win, text="Input File", bg="#303030",
                               fg="white")
        self.fin_label.place(x=15, y=15)
        self.fout_label = Label(self.main_win, text="Output Folder", bg="#303030",
                                fg="white")
        self.fout_label.place(x=15, y=70)

        self.search_label = Label(
            self.main_win, text="Searching: " + self.src, bg="#303030", fg="white")
        self.search_label.place(x=15, y=125)

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

        if self.fin_clicked == '':
            pass
        else:
            self.fin_text.delete(1.0, 2.0)
            self.fin_text.insert(1.0, self.fin_clicked)

    def fout_click(self):
        self.fout_clicked = filedialog.askdirectory()

        if self.fout_clicked == '':
            pass
        else:
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
        # If match found, copy file to new zip folder.
        dst = self.fout_path.rstrip() + "/"
        is_match = []

        for fname in os.listdir(self.src):
            for partno in part_numbers:
                if fname.upper() == partno.upper():
                    # copyfile(self.src + fname, dst + fname)
                    is_match.append(self.src + fname)

        with zipfile.ZipFile(dst + "transfer.zip", "w", compression=zipfile.ZIP_DEFLATED) as zipf:
            for match in is_match:
                zipf.write(match, basename(match))

    # Reset file selections.

    def new(self):
        self.fin_text.delete(1.0, 2.0)
        self.fin_text.insert(1.0, "Select a file")
        self.fout_text.delete(1.0, 2.0)
        self.fout_text.insert(1.0, "Select a Folder")
        self.src = "\\\\andros-dc\\groups\\ENG DATA\\ENG_Directory\\Attachments_Genius\\"
        self.search_label.config(text="Searching: " + self.src)

    def search_dir(self):
        self.search_dir_clicked = filedialog.askdirectory()

        if self.search_dir_clicked == '':
            pass
        else:
            self.src = self.search_dir_clicked.rstrip() + "/"
            self.search_label.config(text="Searching: " + self.src)

    # Application instructions pop-up window.

    def help_menu(self):
        Instructions = "Insert instructions."
        showinfo("Help", Instructions)

# Run program


def main():
    root = tk.Tk()
    MainApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
