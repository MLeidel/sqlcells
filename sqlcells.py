'''
code file: sqlcells.py
date: Dec 2024
comments:
    Use SQL on Spreadsheets (.xlsx, csv)
'''
import os, sys
from tkinter.font import Font
from tkinter import Listbox
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
from datetime import datetime
import subprocess
import pandas as pd
import pandasql as psql
from ttkbootstrap import *
from ttkbootstrap.constants import *

cdf = ""
cfile = ""
ctype = ""
tb1, tb2, tb3, tb4, tb5, tb6 = "", "", "", "", "", ""
tbs = []


class Application(Frame):
    ''' main class docstring '''
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.pack(fill=BOTH, expand=True, padx=4, pady=4)
        self.create_widgets()

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

        btn_clear = Button(frm1, text='Clear', bootstyle="outline", command=self.on_clear)
        btn_clear.grid(row=1, column=3, padx=4, pady=4)

        self.sqltext = Text(self)
        self.sqltext.grid(row=3, column=1, sticky="nsew")

        efont = Font(family="Fira Code", size=10)
        self.sqltext.configure(font=efont)
        self.sqltext.config(wrap="word", # wrap=NONE
                            undo=True, # Tk 8.4
                            #width=20,
                            height=8,
                            padx=5, # inner margin
                            insertbackground='#fff',   # cursor color
                            tabs=(efont.measure(' ' * 4),))
        self.scroll_text = Scrollbar(self, orient=VERTICAL, command=self.sqltext.yview)
        self.scroll_text.grid(row=3, column=2, sticky="wns")  # use nse
        self.sqltext['yscrollcommand'] = self.scroll_text.set


        frm2 = Frame(self)
        frm2.grid(row=4, column=1)


        self.ventr = StringVar()
        # self.ventr.trace("w", self.eventHandler)
        entr = Entry(frm2, textvariable=self.ventr, width=65)
        entr.grid(row=1, column=1, padx=0, pady=4)
        self.ventr.set("out.xlsx")

        btn_submit = Button(frm2, text='Output', bootstyle="outline", command=self.on_output)
        btn_submit.grid(row=1, column=2, padx=8, pady=4)


        frm3 = Frame(self)
        frm3.grid(row=5, column=1)


        btn_submit = Button(frm3, text='Submit', bootstyle="outline", command=self.on_submit)
        btn_submit.grid(row=1, column=1, padx=8, pady=4)

        self.vckbox = IntVar()
        ckbox = Checkbutton(frm3, variable=self.vckbox, text='Launch')
        ckbox.grid(row=1, column=3, padx=8, pady=4)
        self.vckbox.set(1)

        self.vSckbox = IntVar()
        Sckbox = Checkbutton(frm3, variable=self.vSckbox, text=' Log ')
        Sckbox.grid(row=1, column=4, padx=8, pady=4)
        self.vSckbox.set(1)

        root.bind("<Control-q>", self.on_exit)
        self.read_saved_query('lastquery')

    # ----------------------------------------------------------------------------

    def on_input(self):
        ''' set an input table (csv or excel file) '''
        fname =  filedialog.askopenfilename(initialdir = p,
                                                         title = "Open file",
                                                         filetypes = (("xlsx files","*.xlsx"),
                                                         ("csv files","*.csv"),("all files","*.*")))
        if fname:
            try:
                tb = "tb" + str(self.lstn.size() + 1)
                self.lstn.insert(tk.END, tb + ": " + fname)
            except:
                messagebox.showerror("Open File", "Failed to open file\n'%s'" % fname)

    def on_output(self):
        ''' select an output file (xlsx or csv) for the query '''
        fname = filedialog.asksaveasfilename(confirmoverwrite=True,
                                            initialdir=os.path.dirname(os.path.abspath(__file__)),
                                            title = "Save Results",
                                            filetypes = (("xlsx files","*.xlsx"),
                                            ("csv files","*.csv"),("all files","*.*")) )
        if fname:
            self.ventr.set(fname)

    def on_clear(self):
        ''' Remove all input files from the list '''
        items = list(self.lstn.get(0, tk.END))
        for f in items:
            self.lstn.delete(0)

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
        self.load_data_frames()
        # now get the SQL code and execute
        try:
            result_df = psql.sqldf(query, globals())
        except Exception as e:
            messagebox.showerror("An error occurred", e)
            return
        # now create output file and optionally launch it
        if outfile.endswith(".xlsx"):
            result_df.to_excel(outfile, index=False)  # save to new spreadsheet
        else:
            result_df.to_csv(outfile)
        # check to see if launch spreadsheet requested
        if self.vckbox.get() == 1:
            subprocess.Popen(['libreoffice', '--calc', outfile])
        # check to see if logging requested
        if self.vSckbox.get() == 1:
            with open("sqllog.txt", "a", encoding='utf-8') as fout:
                fout.write(str(datetime.today()) + "\n\n")
                items = list(self.lstn.get(0, tk.END))
                for f in items:
                    fout.write(f + "\n")
                fout.write("\n" + query + "-------------------\n")

    def parse_input(self, strg):
        ''' split out the data frame name file path,
        and file type into cdf and cfile global vars '''
        global cdf, cfile, ctype
        items = strg.split(": ")
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
        using global vars tb1 ... tb6 - MAX 6 files '''
        global tb1, tb2, tb3, tb4, tb5, tb6, tbs
        tbs.clear()
        items = list(self.lstn.get(0, tk.END))
        for f in items:
            self.parse_input(f)
            if ctype == "csv":
                df = pd.read_csv(cfile)
            else:
                df = pd.read_excel(cfile)
            tbs.append(df)
        # tb1, tb2, tb3, tb4, tb5 = tbs[0], tbs[1], tbs[2], tbs[3], tbs[4]
        ntbs = len(items)
        if ntbs == 1:
            tb1 = tbs[0]
        elif ntbs == 2:
            tb1, tb2 = tbs[0], tbs[1]
        elif ntbs == 3:
            tb1, tb2, tb3 = tbs[0], tbs[1], tbs[2]
        elif ntbs == 4:
            tb1, tb2, tb3, tb4 = tbs[0], tbs[1], tbs[2], tbs[3]
        elif ntbs == 5:
            tb1, tb2, tb3, tb4, tb5 = tbs[0], tbs[1], tbs[2], tbs[3], tbs[4]
        elif ntbs == 6:
            tb1, tb2, tb3, tb4, tb5, tb6 = tbs[0], tbs[1], tbs[2], tbs[3], tbs[4], tbs[5]

    def read_saved_query(self, filepath):
        ''' reads file of saved query code and displays in user's GUI '''
        code = ""
        with open(filepath, "r", encoding='utf-8') as fin:
            line = fin.readline().strip()
            self.on_clear()
            while line != "SQL":
                self.lstn.insert(tk.END, line)
                line = fin.readline().strip()
            while True:
                line = fin.readline()
                if line == '':
                    break
                code += line
        self.sqltext.insert(1.0, code)

    def save_query(self, filepath):
        ''' writes input file paths and SQL code to filepath '''
        with open(filepath, "w", encoding='utf-8') as fout:
            items = list(self.lstn.get(0, tk.END))
            for line in items:
                fout.write(line + "\n")
            fout.write("SQL" + "\n")
            sql = self.sqltext.get("1.0", END)
            fout.write(sql)

    def on_exit(self, e=None):
        ''' Control-Q saves the current query details '''
        self.save_query("lastquery")
        save_location()  # exit program

    def prompt_info(self, e=None):
        ''' Does user want to see spreadsheet or only column names/types '''
        request = simpledialog.askinteger("Spreadsheet Action",
                                          "Enter 1 to open spreadsheet\nEnter 2 to view file Info",
                                          parent=self,
                                          initialvalue=1)
        list_item = self.lstn.get(ANCHOR)
        self.parse_input(list_item)
        if request == 1:
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

# change working directory to path for this file
p = os.path.realpath(__file__)
os.chdir(os.path.dirname(p))

# THEMES
# 'cosmo', 'flatly', 'litera', 'minty', 'lumen',
# 'sandstone', 'yeti', 'pulse', 'united', 'morph',
# 'journal', 'darkly', 'superhero', 'solar', 'cyborg',
# 'vapor', 'simplex', 'cerculean'
root = Window("SQLcells", "darkly", size=(400, 400))

# UNCOMMENT THE FOLLOWING TO SAVE GEOMETRY INFO
def save_location(e=None):
    ''' executes at WM_DELETE_WINDOW event - see below '''
    with open("winfo", "w", encoding='utf-8') as fout:
        fout.write(root.geometry())
    root.destroy()

# UNCOMMENT THE FOLLOWING TO SAVE GEOMETRY INFO
if os.path.isfile("winfo"):
    with open("winfo") as f:
        lcoor = f.read()
    root.geometry(lcoor.strip())
else:
    root.geometry("400x300") # WxH+left+top


root.protocol("WM_DELETE_WINDOW", save_location)  # UNCOMMENT TO SAVE GEOMETRY INFO
Sizegrip(root).place(rely=1.0, relx=1.0, x=0, y=0, anchor='se')
#root.resizable(0, 0) # no resize & removes maximize button
# root.minsize(w, h)  # width, height
# root.maxsize(w, h)
# root.overrideredirect(True) # removed window decorations
# root.attributes('-type', 'splash')  # don't show in taskbar
# root.iconphoto(False, PhotoImage(file='icon.png'))
# root.attributes("-topmost", True)  # Keep on top of other windows

Application(root)

root.mainloop()