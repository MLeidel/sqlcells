'''
code file: sqlcells.py
date: Dec 2024
comments:
    Use SQL on Spreadsheets (.xlsx, csv)
        input spreadsheet(s)
        output spreadsheet or csv
'''
import os
import sys
from tkinter.font import Font
from tkinter import Listbox
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
from datetime import datetime
import subprocess
import platform
import pandas as pd
import pandasql as psql
import sqlite3
import threading
from ttkbootstrap import *
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
from ttkbootstrap.toast import ToastNotification

cdf = ""
cfile = ""
ctype = ""
d1, d2, d3, d4, d5, d6, d7 = "", "", "", "", "", "", ""
# ds = []
toast = ToastNotification(
    title="SQLcells",
    message="Query Setup Saved!",
    duration=2500,
)


class Application(Frame):
    ''' main class docstring '''
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.pack(fill=BOTH, expand=True, padx=4, pady=4)
        self.create_widgets()
        self.savefile = ""

    def create_widgets(self):
        ''' creates GUI for app '''

        self.lstn = Listbox(self, exportselection=False, width=60, height=4)
        self.lstn.grid(row=1, column=1, sticky="wnse")

        self.lstn.bind("<<ListboxSelect>>", self.prompt_info)

        self.scroll_list = Scrollbar(self, orient=VERTICAL, command=self.lstn.yview)
        self.scroll_list.grid(row=1, column=2, sticky="wns")  # use nse
        self.lstn['yscrollcommand'] = self.scroll_list.set


        frm1 = Frame(self)
        frm1.grid(row=2, column=1, columnspan=2)

        btn_input = Button(frm1, text='Load Input Files', bootstyle="outline", command=self.on_input)
        btn_input.grid(row=1, column=1, padx=4, pady=4)
        ToolTip(btn_input, text="locate an .xlsx or .csv file for input")

        btn_clear = Button(frm1, text='Clear', bootstyle="outline", command=self.on_clear)
        btn_clear.grid(row=1, column=3, padx=4, pady=4)
        ToolTip(btn_clear, text="remove all input files")

        self.sqltext = Text(self)
        self.sqltext.grid(row=3, column=1, sticky="nsew")

        efont = Font(family="Fira Code", size=10)
        self.sqltext.configure(font=efont)
        self.sqltext.config(wrap="word", # wrap=NONE
                            undo=True, # Tk 8.4
                            #width=20,
                            fg='lightgreen',
                            height=8,
                            padx=5, # inner margin
                            insertbackground='#fff',   # cursor color
                            tabs=(efont.measure(' ' * 4),))
        self.scroll_text = Scrollbar(self, orient=VERTICAL, command=self.sqltext.yview)
        self.scroll_text.grid(row=3, column=2, sticky="wns")  # use nse
        self.sqltext['yscrollcommand'] = self.scroll_text.set

        #
        # The following two lines permit
        # the Sql text area to expand
        # everything else stays put
        #
        self.rowconfigure(3, weight=1)
        self.columnconfigure(1, weight=1)


        frm2 = Frame(self)
        frm2.grid(row=4, column=1)

        self.ventr = StringVar()
        entr = Entry(frm2, textvariable=self.ventr, width=65)
        entr.grid(row=1, column=1, padx=0, pady=4)
        ToolTip(entr, text="designate fullpath to query result file")

        btn_submit = Button(frm2, text='Output', bootstyle="outline", command=self.on_output)
        btn_submit.grid(row=1, column=2, padx=8, pady=4)
        ToolTip(btn_submit, text="designate fullpath to query result file")

        frm3 = Frame(self)
        frm3.grid(row=5, column=1)

        btn_submit = Button(frm3, text='Submit', bootstyle="outline", command=self.on_submit)
        btn_submit.grid(row=1, column=1, padx=8, pady=4)
        ToolTip(btn_submit, text="run the query")

        self.vckbox = IntVar()
        ckbox = Checkbutton(frm3, variable=self.vckbox, text='Launch')
        ckbox.grid(row=1, column=3, padx=8, pady=4)
        ToolTip(ckbox, text="open resulting query in LibreOffice Calc")

        self.vSckbox = IntVar()
        Sckbox = Checkbutton(frm3, variable=self.vSckbox, text=' Log ')
        Sckbox.grid(row=1, column=4, padx=8, pady=4)
        ToolTip(Sckbox, text="record the query setup in the log file")

        btn_save = Button(frm3, text='Save', bootstyle="outline", command=self.on_save)
        btn_save.grid(row=1, column=5, padx=8, pady=4)
        ToolTip(btn_save, text="Save this query setup to a file")

        btn_open = Button(frm3, text='Open', bootstyle="outline", command=self.on_open)
        btn_open.grid(row=1, column=6, padx=8, pady=4)
        ToolTip(btn_open, text="Open a saved query setup file")

        btn_close = Button(frm3, text='Close', bootstyle="outline", command=self.on_exit)
        btn_close.grid(row=1, column=7, padx=8, pady=4)
        ToolTip(btn_close, text="Ctrl-Q")

        root.bind("<Control-q>", self.on_exit)
        root.bind("<Control-s>", self.quicksave)

        # to use sqlcells in a "batch" command line operation
        # simply use a saved SQL setup file as argument 1
        if len(sys.argv) > 1:
            self.read_saved_query(sys.argv[1])
            self.on_submit()
            self.on_exit()
        else:
            if os.path.isfile("lastquery"):
                self.read_saved_query('lastquery')  # open last query setup

        ###################################################################
        # txt bg = #333
        # txt fg = #DEE
        self.sqltext.tag_configure("literals",foreground="darkorange")
        self.sqltext.tag_configure("remarks", foreground="gray")

        ###################################################################

        self.highlite()  # starts off syntax highliting thread

    # ----------------------------------------------------------------------------

    def on_save(self):
        ''' Save SQL query setup to a txt file '''
        path = self.lstn.get(0)
        self.parse_input(path)
        fname = filedialog.asksaveasfilename(confirmoverwrite=True,
                                            initialdir=os.path.dirname(cfile),
                                            title = "Save Query",
                                            filetypes = (("all files","*.*"),("text files","*.txt")) )
        if fname:
            self.save_query(fname)


    def on_open(self):
        ''' Reads the SQL setup to a file.
            Uses path from d1 input file '''
        path = self.lstn.get(0)
        if path == "":
            dirname = p
        else:
            self.parse_input(path)
            dirname = cfile  # global cfile from parse_input
        fname =  filedialog.askopenfilename(initialdir=os.path.dirname(dirname),
                                            title = "Open Query",
                                            filetypes = (("all files","*.*"),("text files","*.txt")))
        if fname:
            self.read_saved_query(fname)


    def on_input(self):
        ''' set an input table (csv or excel file) '''
        fname =  filedialog.askopenfilename(initialdir = p,
                                            title = "Open file",
                                            filetypes = (("xlsx files","*.xls*"),
                                            ("csv files","*.csv"),("all files","*.*")))
        if fname:
            try:
                s = self.lstn.size()
                if s >= 7:
                    messagebox.showwarning("Input Limit", "7 inputs max")
                    return
                d = "d" + str(s + 1)
                self.lstn.insert(tk.END, d + ": " + fname)
            except:
                messagebox.showerror("Open File", "Failed to open file\n'%s'" % fname)

    def on_output(self):
        ''' select an output file (xlsx or csv) for the query '''
        path = self.lstn.get(0)
        self.parse_input(path)
        fname = filedialog.asksaveasfilename(confirmoverwrite=True,
                                            initialdir=os.path.dirname(cfile),
                                            title = "Save Results",
                                            filetypes = (("xlsx files","*.xls*"),
                                            ("csv files","*.csv"),("all files","*.*")) )
        if fname:
            self.ventr.set(fname)

    def on_clear(self):
        ''' Remove all input files from the list '''
        self.lstn.delete(0, tk.END)

    def on_submit(self):
        ''' load the data frames and execute the SQL
        optionally launch the result and optionally
        log the query information '''
        outfile = self.ventr.get()
        if outfile == "":
            messagebox.showerror("Output", "Output file missing")
            return
        if self.lstn.size() == 0:
            messagebox.showerror("Input", "Input files missing")
            return

        query = self.sqltext.get("1.0", END)
        if len(query) < 5:
            messagebox.showerror("Query", "Query Code missing")
            return

        # Filter out query lines that start with #
        lines = query.splitlines()
        sql_lines = [line for line in lines if not line.lstrip().startswith("#")]
        query = "\n".join(sql_lines)

        self.load_data_frames()
        # now get the SQL code and execute
        try:
            result_df = psql.sqldf(query, globals())
        except Exception as e:
            messagebox.showerror("An error occurred", e)
            return
        # now create output file and optionally launch it
        if outfile.endswith((".xlsx", ".xls")):
            result_df.to_excel(outfile, index=False)  # save to new spreadsheet
        elif outfile.endswith(".csv"):
            result_df.to_csv(outfile)
        elif outfile.endswith((".sqlite", ".db")):
            # Export to SQLite Database
            # Connect to (or create) the SQLite database
            conn = sqlite3.connect(outfile)
            # Choose a table name, e.g., 'result_table'. You can make this dynamic if needed.
            table_name = 'result_table'
            # Write the DataFrame to the SQLite table
            result_df.to_sql(table_name, conn, if_exists='replace', index=False)
            # Commit changes and close the connection
            conn.commit()
            conn.close()
        else:
            messagebox.showerror("Unsupported file format", "The specified file format is not supported.")
            return
        # check to see if launch spreadsheet requested
        if outfile.endswith((".sqlite", ".db")):
            messagebox.showinfo("Sqlite", "Database with result_table was created")
        elif self.vckbox.get() == 1:
            if platform.system() == 'Windows':
                subprocess.Popen(["C:\\Program Files\\LibreOffice\\program\\scalc.exe",  outfile])
            else:
                subprocess.Popen(['libreoffice', '--calc', outfile])
        # check to see if logging requested
        if self.vSckbox.get() == 1:
            with open("sqllog.txt", "a", encoding='utf-8') as fout:
                fout.write(str(datetime.today()) + "\n\n")
                items = list(self.lstn.get(0, tk.END))
                for f in items:
                    fout.write(f + "\n")
                fout.write("\n" + query.strip() + "\n\n" + outfile + "\n-------------------------------------\n")

    def parse_input(self, strg):
        ''' split out the data frame name file path,
        and file type into cdf and cfile global vars '''
        global cdf, cfile, ctype
        items = strg.split(": ")
        if len(items) == 0:
            return
        cdf = items[0]
        cfile = items[1]
        if cfile.endswith("xlsx") or cfile.endswith("xls"):
            ctype = "xls"
        elif cfile.endswith("csv"):
            ctype = "csv"
        else:
            messagebox.showerror("Error", "invalid file type")

    def load_data_frames(self):
        ''' load data frames from the list of input files (tables)
        using global vars d1 ... d6 - MAX 6 files '''
        global d1, d2, d3, d4, d5, d6, d7
        ds = []
        # ds.clear()
        items = list(self.lstn.get(0, tk.END))
        for f in items:
            self.parse_input(f)
            if ctype == "csv":
                df = pd.read_csv(cfile)
            else:
                df = pd.read_excel(cfile)
            ds.append(df)
        # d1, d2, d3, d4, d5 ... = ds[0], ds[1], ds[2], ds[3], ds[4] ...
        nds = len(items)
        if nds == 1:
            d1 = ds[0]
        elif nds == 2:
            d1, d2 = ds[0], ds[1]
        elif nds == 3:
            d1, d2, d3 = ds[0], ds[1], ds[2]
        elif nds == 4:
            d1, d2, d3, d4 = ds[0], ds[1], ds[2], ds[3]
        elif nds == 5:
            d1, d2, d3, d4, d5 = ds[0], ds[1], ds[2], ds[3], ds[4]
        elif nds == 6:
            d1, d2, d3, d4, d5, d6 = ds[0], ds[1], ds[2], ds[3], ds[4], ds[5]
        elif nds == 7:
            d1, d2, d3, d4, d5, d6, d7 = ds[0], ds[1], ds[2], ds[3], ds[4], ds[5], ds[6]

    def read_saved_query(self, filepath):
        ''' reads file of saved query code and displays in user's GUI '''
        code = ""
        self.savefile = filepath # for quicksave
        try:
            with open(filepath, "r", encoding='utf-8') as fin:
                line = fin.readline().strip()
                self.on_clear()
                while line != "SQL":
                    self.lstn.insert(tk.END, line)
                    line = fin.readline().strip()
                while True:
                    line = fin.readline()  # now reading the SQL lines (with EOLs)
                    if line == '' or line.startswith("OUTPUT"):
                        break
                    code += line  # concatenate all the SQL lines
                path = fin.readline().strip() # read the output path
                # now read until end of file checking for LAUNCH and LOG
                while line != '':
                    line = fin.readline().strip()
                    if line == "LAUNCH":
                        self.vckbox.set(1)
                    if line == "LOG":
                        self.vSckbox.set(1)
        except:
            messagebox.showerror("Reading File Error", "Re-check the FILE TYPE")
            return
        self.sqltext.delete("1.0", END)  # clear the Text widget
        self.sqltext.insert(1.0, code.strip())  # insert the SQL code
        self.ventr.set(path)  # output path

    def save_query(self, filepath):
        ''' writes input file paths and SQL code to filepath '''
        if filepath.endswith(".csv") or filepath.endswith(".xlsx") or filepath.endswith(".xls"):
            messagebox.showwarning("Saving Query Code", "Incorrect File Type!")
            return
        self.savefile = filepath  # for quicksave
        with open(filepath, "w", encoding='utf-8') as fout:
            items = list(self.lstn.get(0, tk.END))
            for line in items:
                fout.write(line + "\n")
            fout.write("SQL" + "\n")
            sql = self.sqltext.get("1.0", END)
            fout.write(sql)
            fout.write("OUTPUT" + "\n")
            outfile = self.ventr.get()
            fout.write(outfile + "\n")
            if self.vckbox.get() == 1:
                fout.write("LAUNCH" + "\n")
            if self.vSckbox.get() == 1:
                fout.write("LOG" + "\n")
        toast.show_toast()

    def on_exit(self, e=None):
        ''' Control-Q saves the current query details '''
        self.save_query("lastquery")
        save_location()  # exit program

    def quicksave(self, e=None):
        ''' Save with Ctrl-S when self.savefile != ""
            otherwise, open for save with filedialog.
            savefile set with on_open and on_save. '''
        if self.savefile != "":
            self.save_query(self.savefile)
        else:
            self.on_save()

    def prompt_info(self, e=None):
        ''' Does user want to see spreadsheet or only column names/types '''
        list_item = self.lstn.get(ANCHOR)
        if list_item == "":
            return  # nothing selected
        request = simpledialog.askinteger("Viewing",
                            "Enter 1 to open spreadsheet\nEnter 2 to view file Info\n",
                            parent=self,
                            initialvalue=1)
        self.parse_input(list_item)
        if request == 1:
            if platform.system() == 'Windows':
                subprocess.Popen(["C:\\Program Files\\LibreOffice\\program\\scalc.exe",  cfile])
            else:
                subprocess.Popen(['libreoffice', '--calc', cfile])
        elif request == 2:
            if ctype == "csv":
                df = pd.read_csv(cfile)
            else:
                df = pd.read_excel(cfile)
            self.info_window(df.dtypes)

    def info_window(self, dframe):
        ''' user want to see column names/types '''
        t = Toplevel(self)
        t.wm_title("Info")
        l = Label(t, text=dframe)
        l.grid(row=0, column=0, padx=10, pady=10)
        btn = Button(t, text="Close", command = t.destroy)
        btn.grid(row=1, column=0, sticky='sew', padx=5, pady=5)


    # SQL code hilighting for literals and remarks follows

    def highlite(self):
        global t

        self.highlight_pattern(r"(#.*|//.*)\n", "remarks", regexp=True)

        self.highlight_pattern(r"[\"\'`]((?:.|\n)*?)[\'\"`]",
                               "literals", regexp=True)

        t = threading.Timer(1.25, self.highlite)  # every 1.5 seconds
        t.daemon = True  # for threading runtime error
        t.start()

    def highlight_pattern(self, pattern, tag, start="1.0", end="end", regexp=False):
        start = self.sqltext.index(start)
        end = self.sqltext.index(end)
        self.sqltext.tag_remove(tag, start, end)
        self.sqltext.mark_set("matchStart", start)
        self.sqltext.mark_set("matchEnd", start)
        self.sqltext.mark_set("searchLimit", end)

        count = IntVar()
        while True:
            index = self.sqltext.search(pattern, "matchEnd","searchLimit",
                                count=count, regexp=True)
            if index == "": break
            if count.get() == 0: break # degenerate pattern zero-length strings
            self.sqltext.mark_set("matchStart", index)
            self.sqltext.mark_set("matchEnd", "%s+%sc" % (index, count.get()))
            self.sqltext.tag_add(tag, "matchStart", "matchEnd")


# change working directory to path for this file
p = os.path.realpath(__file__)
os.chdir(os.path.dirname(p))

# THEMES
# 'cosmo', 'flatly', 'litera', 'minty', 'lumen',
# 'sandstone', 'yeti', 'pulse', 'united', 'morph',
# 'journal', 'darkly', 'superhero', 'solar', 'cyborg',
# 'vapor', 'simplex', 'cerculean'
root = Window("SQLcells", "darkly", size=(673, 372))

# UNCOMMENT THE FOLLOWING TO SAVE GEOMETRY INFO
def save_location(e=None):
    ''' executes at WM_DELETE_WINDOW event - see below '''
    with open("winfo", "w", encoding='utf-8') as fout:
        fout.write(root.geometry())
    root.destroy()

# UNCOMMENT THE FOLLOWING TO SAVE GEOMETRY INFO
if os.path.isfile("winfo"):
    with open("winfo") as z:
        lcoor = z.read()
    root.geometry(lcoor.strip())
else:
    root.geometry("673x372") # WxH+left+top


root.protocol("WM_DELETE_WINDOW", save_location)  # UNCOMMENT TO SAVE GEOMETRY INFO
Sizegrip(root).place(rely=1.0, relx=1.0, x=0, y=0, anchor='se')
# root.resizable(0, 0) # no resize & removes maximize button
root.minsize(650, 375)  # width, height
# root.maxsize(680, 379)
# root.overrideredirect(True) # removed window decorations
# root.attributes('-type', 'splash')  # don't show in taskbar
# root.iconphoto(False, PhotoImage(file='icon.png'))
# root.attributes("-topmost", True)  # Keep on top of other windows

Application(root)

root.mainloop()