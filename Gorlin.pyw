#!/usr/bin/env python3
#

from zipfile import ZipFile
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showerror
from tkinter.filedialog import askopenfilename
from tkinter.scrolledtext import ScrolledText
from io import BytesIO
import os
from base64 import b64decode, encodebytes
import json
import time

EXTERNAL_TEMPLATE = None # if set, will use that file instead of internal template
USE_BRACKETSAVER = True # if set, will attempt to correct for MS spellchecker.
SPLITTER = "###"+"~~~"+"###"
DEBUG = False

def update_template(newfn='', newjson=''):
    if newfn:
        with open(newfn, 'rb') as f:
            file_data = f.read()
            file_data = encodebytes(file_data).decode()
        file_meta = str({
            'name':os.path.split(newfn)[1],
            'date':time.strftime("%Y-%m-%d, %H:%M:%S")})
    else:
        file_data = template_docx.strip()
        file_meta = str(template_docx_meta)
    new_data = '\ntemplate_docx = """\n'+file_data+'"""\n'
    new_data += 'template_docx_meta = '+file_meta+'\n'

    if newjson:
        json_data = newjson
    else:
        json_data = json.dumps(template_options, indent=2)
    new_data += 'template_options = json.loads("""' + json_data + '\n""")\n'

    with open(__file__, 'r') as f:
        prog_data = f.read()
        prog_data = prog_data.split(SPLITTER)

    if len(prog_data) != 3:
        print('big error')
        return

    prog_data[1] = new_data
    with open(__file__, 'w') as f:
        f.write(SPLITTER.join(prog_data))

def bracket_saver(data):
    """Function to de-styleize python escape sequences from MS formatting"""
    # ~ data = ''.join(data.split())
    output = ""
    chunks = data.split("}")
    for chunk in chunks[:-1]:
        before, chunk = chunk.split('{', 1)
        new_data = ""
        chunk = iter(chunk)
        for char in chunk:
            if char == "<":
                while next(chunk) != ">":
                    pass
            else:
                new_data += char
        output += f"{before}{{{new_data}}}"
    return output+chunks[-1]

def intract(zipread, fn, newdata, new_zname):
    with ZipFile(new_zname, 'w') as zipwrite:
        for item in zipread.infolist():
            if item.filename != fn:
                data = zipread.read(item.filename)
                zipwrite.writestr(item, data)
        zipwrite.writestr(fn, newdata)
        #~ zipwrite.write(fn, 'content.xml')
    print('file made')

def create_wordfile(fn, datain):
    if EXTERNAL_TEMPLATE:
        zf = ZipFile(EXTERNAL_TEMPLATE)
    else:
        zf = ZipFile(BytesIO(b64decode(template_docx)))
    data = zf.open("word/document.xml")
    content = data.read().decode()
    if DEBUG:
        with open('before.xml', 'w') as f:
            f.write(content.replace('><', '>\n<'))
    if USE_BRACKETSAVER:
        content = bracket_saver(content)
    if DEBUG:
        with open('after.xml', 'w') as f:
            f.write(content.replace('><', '>\n<'))
    newdata = content.format(**datain)
    intract(zf, "word/document.xml", newdata, fn)

def make_sentence(data, prefix="Patient "):
    if not data:
        return 'No notes'
    output = ", ".join(data[:-1])
    if len(data) > 1: output += " and "
    output += data[-1]
    return prefix + output

class Checkbutton(ttk.Checkbutton):
    def __init__(self, master=None, value=0, **kwargs):
        self.var = tk.IntVar(value=value)
        super().__init__(master, variable=self.var, **kwargs)
        self.set, self.get = self.var.set, self.var.get

class Settings(tk.Toplevel):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)

        lf = ttk.LabelFrame(self, text="MS Word template:")
        lf.pack(anchor=tk.W)
        self.template_path = None
        btn = ttk.Button(lf, text="Browse for new template", command=self.template_browse)
        btn.pack()
        self.template_lbl = ttk.Label(lf, text=template_docx_meta['name'])
        self.template_lbl.pack()
        self.template_lbl_date = ttk.Label(lf, text=template_docx_meta['date'])
        self.template_lbl_date.pack()

        lf = ttk.LabelFrame(self, text="options template:")
        # ~ lf.pack(anchor=tk.W)
        self.json_data = ScrolledText(lf, width=40)
        self.json_data.insert(tk.END, json.dumps(template_options, indent=2))
        self.json_data.pack()

        btn = ttk.Button(self, text="apply", command=self.apply)
        btn.pack()

        lbl = tk.Label(self,
            text="Program must be restarted\nfor changes to take effect",
            fg='red')
        lbl.pack()

        self.transient(master)
        self.grab_set()
        master.wait_window(self)

    def template_browse(self):
        path = askopenfilename(
            initialdir = ".",
            title = "Select file",
            filetypes = (("Word files","*.docx"),("all files","*.*")))
        if not path:
            return # user cancel
        self.template_path = path
        _, fn = os.path.split(path)
        self.template_lbl.config(text=fn)
        self.template_lbl_date.config(text="not yet updated")

    def apply(self):
        update_template(self.template_path, self.json_data.get(0.0,tk.END))
        self.destroy()
        self.master.quit()

class Menu(tk.Menu):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)

        self.add_command(label="Settings", command=self.settings_open)

    def settings_open(self):
        Settings(self.master)

class GUI(tk.Frame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)

        m = Menu(tk._default_root)
        tk._default_root.config(menu=m)

        lf = ttk.LabelFrame(self, text="Enter patient number")
        lf.grid(sticky='ew')
        self.pat_num = ttk.Entry(lf)
        self.pat_num.focus()
        self.pat_num.pack()

        lf = ttk.LabelFrame(self, text="Lungs")
        lf.grid(sticky='ew')
        self.lungs = [
            Checkbutton(lf, text="coughs"),
            Checkbutton(lf, text="does not cough"),
            Checkbutton(lf, text="oxygen level is low"),
            ]
        for widget in self.lungs:
            widget.pack(anchor=tk.W)

        lf = ttk.LabelFrame(self, text="Cardio")
        lf.grid(sticky='ew')
        self.cardio = [
            Checkbutton(lf, text="heart beat is fast"),
            Checkbutton(lf, text="heart beat is slow"),
            Checkbutton(lf, text="heart wheezing can be heard"),
            ]
        for widget in self.cardio:
            widget.pack(anchor=tk.W)

        btn = ttk.Button(self, text='Create word file', command=self.create)
        btn.grid()

    def create(self):
        pat_num = self.pat_num.get()
        if not pat_num:
            showerror("Error", "Patient number is required")
            return # stop this function

        data = dict(
            lungs = make_sentence([w['text'] for w in self.lungs if w.get()]),
            cardio = make_sentence([w['text'] for w in self.cardio if w.get()]),
            pat_num = pat_num,
            )

        filename = f'{pat_num}.docx'
        if os.path.exists(filename):
            showerror('Error', "A file for that patient already exists")
            return

        try:
            create_wordfile(filename, data)
            os.startfile(filename)
        except Exception as e:
            showerror("Unknown Error", str(e))


# TEMPLATE DATA
###~~~###
template_docx = """
UEsDBBQABgAIAAAAIQDfpNJsWgEAACAFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAAC
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0
lMtuwjAQRfeV+g+Rt1Vi6KKqKgKLPpYtUukHGHsCVv2Sx7z+vhMCUVUBkQpsIiUz994zVsaD0dqa
bAkRtXcl6xc9loGTXmk3K9nX5C1/ZBkm4ZQw3kHJNoBsNLy9GUw2ATAjtcOSzVMKT5yjnIMVWPgA
jiqVj1Ykeo0zHoT8FjPg973eA5feJXApT7UHGw5eoBILk7LXNX1uSCIYZNlz01hnlUyEYLQUiep8
6dSflHyXUJBy24NzHfCOGhg/mFBXjgfsdB90NFEryMYipndhqYuvfFRcebmwpCxO2xzg9FWlJbT6
2i1ELwGRztyaoq1Yod2e/ygHpo0BvDxF49sdDymR4BoAO+dOhBVMP69G8cu8E6Si3ImYGrg8Rmvd
CZFoA6F59s/m2NqciqTOcfQBaaPjP8ber2ytzmngADHp039dm0jWZ88H9W2gQB3I5tv7bfgDAAD/
/wMAUEsDBBQABgAIAAAAIQAekRq37wAAAE4CAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJLBasMw
DEDvg/2D0b1R2sEYo04vY9DbGNkHCFtJTBPb2GrX/v082NgCXelhR8vS05PQenOcRnXglF3wGpZV
DYq9Cdb5XsNb+7x4AJWFvKUxeNZw4gyb5vZm/cojSSnKg4tZFYrPGgaR+IiYzcAT5SpE9uWnC2ki
Kc/UYySzo55xVdf3mH4zoJkx1dZqSFt7B6o9Rb6GHbrOGX4KZj+xlzMtkI/C3rJdxFTqk7gyjWop
9SwabDAvJZyRYqwKGvC80ep6o7+nxYmFLAmhCYkv+3xmXBJa/ueK5hk/Nu8hWbRf4W8bnF1B8wEA
AP//AwBQSwMEFAAGAAgAAAAhACrrAznmAwAAHhAAABEAAAB3b3JkL2RvY3VtZW50LnhtbNxXW4+i
SBR+32T/g+G9m5sgktbJ2GhvJ3sx6+zzpIRSSFMUqSq1ncn+9z1VgIA4HdRNNtkXoc7lq3P56hQ+
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
"""
template_docx_meta = {'name': 'template.docx', 'date': '2020-03-25, 13:56:49'}
template_options = json.loads("""{
  "pat_num": {
    "display": "Patient Number",
    "required": 2,
    "type": "entry"
  },
  "cardio": {
    "type": "checkbuttons",
    "options": [
      "heart beat is fast",
      "heart beat is slow",
      "heart wheezing can be heard"
    ]
  },
  "lungs": {
    "type": "checkbuttons",
    "options": [
      "coughs",
      "does not cough",
      "oxygen level is low"
    ]
  }
}

""")
###~~~###
# END TEMPLATE DATA

if __name__ == "__main__":
    root = tk.Tk()
    win = GUI(root)
    win.pack()
    root.mainloop()

