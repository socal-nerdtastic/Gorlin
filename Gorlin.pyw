#!/usr/bin/env python3
#

from zipfile import ZipFile
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showerror
from tkinter.filedialog import askopenfilename
from io import BytesIO
import os
from base64 import b64decode, encodebytes

EXTERNAL_TEMPLATE = None # if set, will use that file instead of internal template
USE_BRACKETSAVER = True # if set, will attempt to correct for MS spellchecker.
SPLITTER = b"###"+b"~~~"+b"###"
DEBUG = False

def update_template(newfn):
    with open(newfn, 'rb') as f:
        new_data = f.read()
        new_data = encodebytes(new_data)
        new_data = b'template_docx = """\n'+new_data+b'"""'

    with open(__file__, 'rb') as f:
        prog_data = f.read()
        prog_data = prog_data.split(SPLITTER)

    if len(prog_data) != 3:
        print('big error')
        return

    prog_data[1] = new_data
    with open(__file__, 'wb') as f:
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
        self.template_lbl = ttk.Label(lf)
        self.template_lbl.pack()

        btn = ttk.Button(self, text="apply", command=self.apply)
        btn.pack()

        lbl = tk.Label(self,
            text="Program must be restarted\nfor changes to take effect",
            fg='red')
        lbl.pack()

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


    def apply(self):
        if self.template_path:
            update_template(self.template_path)
        self.quit()

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
AP//AwBQSwMEFAAGAAgAAAAhAIB32AW8AwAARw4AABEAAAB3b3JkL2RvY3VtZW50LnhtbLyXW2+j
OBSA31ea/xDx3oK5haAmoyQ0nUqzUrTdfR654ARUjJHtJM2O5r/vsYGQlG6XUGlfAr6cz+eOc/f1
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
FVLQQoUItFAhAi1UiEALFSLQQoUIlFCBuZdQIQUtVIhACxUi0EI1PxYOECq0xwkV2vsIFVJ8hAop
aKFCBFqoEIEWKkSghQoRaKFCBEqowNxLqJCCFipEoIUKEV3xaX+idF1mv4s/6+m8Yr//T1e2Uz/q
t3LXUfv9UVWv3Kz+9yKcS/kUtN54uG/qjX4QMY2FNKeoHT+r17nmkgjUD5/fL7rv8KnTBz50yd4L
YX4zBfCDvpbgnMpBV8jXLUGRd9AV6XVLsOo86Mq+dUswDR50JV2jy+qiFDUdAeOuNFMz3nWYd2Xr
mjkc4q4cXTOEI9yVmWuGcIC78nHN8DDQyXnT+rDnOB2tri8FhK5wrBGO3YSusIS+qtIxFEZfp7kJ
fb3nJvR1o5uA8qcTg3esG4X2sBvl52ooM6yr/YXqJmBdDQlergYYf1dDlLerIcrP1TAxYl0NCVhX
+ydnN8HL1QDj72qI8nY1RPm5Gk5lWFdDAtbVkIB19cAJ2YnxdzVEebsaovxcDRd3WFdDAtbVkIB1
NSR4uRpg/F0NUd6uhig/V4MqGe1qSMC6GhKwroYEL1cDjL+rIcrb1RDV5WpzFqXhapSHa+a4RVjN
EDch1wxxyblm6FEt1aw9q6UawbNagr6qfI6rlupOcxP6es9N6OtGNwHlTycG71g3Cu1hN8rP1bhq
qc3V/kJ1E7CuxlVLTlfjqqVOV+OqpU5X46olt6tx1VKbq3HVUpur/ZOzm+Dlaly11OlqXLXU6Wpc
teR2Na5aanM1rlpqczWuWmpz9cAJ2YnxdzWuWup0Na5acrsaVy21uRpXLbW5GlcttbkaVy05XY2r
ljpdjauWOl2Nq5bcrsZVS22uxlVLba7GVUttrsZVS05X46qlTlfjqqVOV+OqpVtlIggeATVJWFYE
dM+Lu2L5vGDDH074M814LuNnHgXoQx2/NN5Zpdswb5hT3y/UgerHltfuMYrKx7ZaoPnidbR6t5Q2
1j0K7Fu87GbTcfsba9miMdzS1Apuf+DdA/j166RMC1Omjuq7HhXQeKqfZtiyXXux2l41czFnWbl3
HV/Vd6yCmgO5cXgvJ1mu6nxrsbNzfHF8uGPlb19C9sT54k51yWzTH25EynPzaf1+sql+NpgalANz
C5V9W5nVniyfvXTzHFftVL60LXS+6o392fGqN73zm92m9zfe9tawXL/tTW8+X73tLdRarfq1d3l4
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
AAAA"""
###~~~###
# END TEMPLATE DATA

if __name__ == "__main__":
    root = tk.Tk()
    win = GUI(root)
    win.pack()
    root.mainloop()

