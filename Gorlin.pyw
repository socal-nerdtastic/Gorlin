#!/usr/bin/env python3
#

from zipfile import ZipFile
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showerror
from io import BytesIO
import tempfile
import shutil
import os
from base64 import b64decode

def intract(zipread, fn, newdata, new_zname):
    tempdir = tempfile.mkdtemp()
    try:
        tempname = os.path.join(tempdir, 'new.zip')
        with ZipFile(tempname, 'w') as zipwrite:
            for item in zipread.infolist():
                if item.filename != fn:
                    data = zipread.read(item.filename)
                    zipwrite.writestr(item, data)
            zipwrite.writestr(fn, newdata)
            #~ zipwrite.write(fn, 'content.xml')
        shutil.move(tempname, new_zname)
    finally:
        shutil.rmtree(tempdir)
    print('file made')

def create_wordfile(fn, datain):
    # ~ zf = ZipFile('template.docx')
    zf = ZipFile(BytesIO(b64decode(template_docx)))
    data = zf.open("word/document.xml")
    content = data.read().decode()
    newdata = content.format(**datain)
    intract(zf, "word/document.xml", newdata, fn)

def make_sentence(data, prefix="Patient "):
    if not data:
        return ''
    output = ", ".join(data[:-1])
    if len(data) > 1: output += " and "
    output += data[-1]
    return prefix + output

class Checkbutton(ttk.Checkbutton):
    def __init__(self, master=None, value=0, **kwargs):
        self.var = tk.IntVar(value=value)
        super().__init__(master, variable=self.var, **kwargs)
        self.set, self.get = self.var.set, self.var.get

class GUI(tk.Frame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)

        lf = ttk.LabelFrame(self, text="Enter patient number")
        lf.grid(sticky='ew')
        self.pat_num = ttk.Entry(lf)
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
            )

        filename = f'{pat_num}.docx'
        create_wordfile(filename, data)
        os.startfile(filename)

# TEMPLATE DATA
###
template_docx = """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
AASqokXNAQAA2QMAABAAAABkb2NQcm9wcy9hcHAueG1snVPLbtswELwX6D8IvMe0nSBIDZpB4aDI
oW0MWEnOW2plE6VIgtwIcb++KylW5aan8jT74HD2QXX72riixZRt8GuxmM1Fgd6Eyvr9WjyWXy5u
RJEJfAUueFyLI2Zxqz9+UNsUIiaymAum8HktDkRxJWU2B2wgzzjsOVKH1ACxmfYy1LU1eBfMS4Oe
5HI+v5b4SugrrC7iSCgGxlVL/0taBdPpy0/lMTKfViU20QGh/t7ddLMqUKPk6FVlIHClbVBf3rB/
tNQW9pj1QskBqOeQqqyvlByA2hwggSFuoF5+UnJiqs8xOmuAuLP6mzUp5FBT8dDLLbrrSk5TFJew
Q/OSLB31XMmpqb5aP8gYAMtKsE8QD2/aRkvtDDjccPG6BpdRyT8OdY/QDXYLttPX0qpFQyEV2f7i
0S5F8QMydi1bixaSBU9iSBuMHruYKenSkmPu0e7hNG2K7VUncgDniXLUwPhcXf9Cfqi5NvqH2MVU
bK9BTOS9U3Z64y/WTWgieO6vHBE3+Gd+jGW46xbjrYfnzsnQny0ddhEMz+RyOR3/JKB27MWK5zmO
ZHSoey4guY6e7/o9Vqec94FuoZ6Gn6oX17M5n36DTj7eg/EL6d9QSwMEFAAAAAAAAmV2UGboynqF
DQAAhQ0AABEAAAB3b3JkL2RvY3VtZW50LnhtbDw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9
IlVURi04IiBzdGFuZGFsb25lPSJ5ZXMiPz4KPHc6ZG9jdW1lbnQgeG1sbnM6d3BjPSJodHRwOi8v
c2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTAvd29yZHByb2Nlc3NpbmdDYW52
YXMiIHhtbG5zOmN4PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5n
LzIwMTQvY2hhcnRleCIgeG1sbnM6Y3gxPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29m
ZmljZS9kcmF3aW5nLzIwMTUvOS84L2NoYXJ0ZXgiIHhtbG5zOmN4Mj0iaHR0cDovL3NjaGVtYXMu
bWljcm9zb2Z0LmNvbS9vZmZpY2UvZHJhd2luZy8yMDE1LzEwLzIxL2NoYXJ0ZXgiIHhtbG5zOmN4
Mz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2UvZHJhd2luZy8yMDE2LzUvOS9j
aGFydGV4IiB4bWxuczpjeDQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL2Ry
YXdpbmcvMjAxNi81LzEwL2NoYXJ0ZXgiIHhtbG5zOmN4NT0iaHR0cDovL3NjaGVtYXMubWljcm9z
b2Z0LmNvbS9vZmZpY2UvZHJhd2luZy8yMDE2LzUvMTEvY2hhcnRleCIgeG1sbnM6Y3g2PSJodHRw
Oi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTYvNS8xMi9jaGFydGV4
IiB4bWxuczpjeDc9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL2RyYXdpbmcv
MjAxNi81LzEzL2NoYXJ0ZXgiIHhtbG5zOmN4OD0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNv
bS9vZmZpY2UvZHJhd2luZy8yMDE2LzUvMTQvY2hhcnRleCIgeG1sbnM6bWM9Imh0dHA6Ly9zY2hl
bWFzLm9wZW54bWxmb3JtYXRzLm9yZy9tYXJrdXAtY29tcGF0aWJpbGl0eS8yMDA2IiB4bWxuczph
aW5rPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTYvaW5r
IiB4bWxuczphbTNkPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5n
LzIwMTcvbW9kZWwzZCIgeG1sbnM6bz0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTpvZmZpY2U6
b2ZmaWNlIiB4bWxuczpyPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNl
RG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzIiB4bWxuczptPSJodHRwOi8vc2NoZW1hcy5vcGVu
eG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9tYXRoIiB4bWxuczp2PSJ1cm46c2No
ZW1hcy1taWNyb3NvZnQtY29tOnZtbCIgeG1sbnM6d3AxND0iaHR0cDovL3NjaGVtYXMubWljcm9z
b2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nRHJhd2luZyIgeG1sbnM6d3A9
Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9kcmF3aW5nbWwvMjAwNi93b3JkcHJv
Y2Vzc2luZ0RyYXdpbmciIHhtbG5zOncxMD0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTpvZmZp
Y2U6d29yZCIgeG1sbnM6dz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3dvcmRw
cm9jZXNzaW5nbWwvMjAwNi9tYWluIiB4bWxuczp3MTQ9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29m
dC5jb20vb2ZmaWNlL3dvcmQvMjAxMC93b3JkbWwiIHhtbG5zOncxNT0iaHR0cDovL3NjaGVtYXMu
bWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEyL3dvcmRtbCIgeG1sbnM6dzE2Y2lkPSJodHRw
Oi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS93b3JkLzIwMTYvd29yZG1sL2NpZCIgeG1s
bnM6dzE2c2U9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAxNS93
b3JkbWwvc3ltZXgiIHhtbG5zOndwZz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZp
Y2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nR3JvdXAiIHhtbG5zOndwaT0iaHR0cDovL3NjaGVt
YXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8yMDEwL3dvcmRwcm9jZXNzaW5nSW5rIiB4bWxu
czp3bmU9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3dvcmQvMjAwNi93b3Jk
bWwiIHhtbG5zOndwcz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvd29yZC8y
MDEwL3dvcmRwcm9jZXNzaW5nU2hhcGUiIG1jOklnbm9yYWJsZT0idzE0IHcxNSB3MTZzZSB3MTZj
aWQgd3AxNCI+Cjx3OmJvZHk+Cjx3OnAgdzE0OnBhcmFJZD0iNzRGREUyODMiIHcxNDp0ZXh0SWQ9
Ijc3Nzc3Nzc3IiB3OnJzaWRSPSIwMDdDNzUwNyIgdzpyc2lkUkRlZmF1bHQ9IjAwN0M3NTA3IiB3
OnJzaWRQPSIwMDdDNzUwNyI+Cjx3OnBQcj4KPHc6cFN0eWxlIHc6dmFsPSJIZWFkaW5nMiIvPgo8
L3c6cFByPgo8dzpyPgo8dzp0Pkx1bmdzOjwvdzp0Pgo8L3c6cj4KPC93OnA+Cjx3OnAgdzE0OnBh
cmFJZD0iNzA2N0JCMEYiIHcxNDp0ZXh0SWQ9IjRBRTdCNUYzIiB3OnJzaWRSPSIwMDA4NTJGRiIg
dzpyc2lkUkRlZmF1bHQ9IjAwREQyOTU2IiB3OnJzaWRQPSIwMDdDNzUwNyI+Cjx3OnBQcj4KPHc6
cFN0eWxlIHc6dmFsPSJOb1NwYWNpbmciLz4KPC93OnBQcj4KPHc6cj4KPHc6dD57bHVuZ3N9PC93
OnQ+CjwvdzpyPgo8L3c6cD4KPHc6cCB3MTQ6cGFyYUlkPSIzNjNCRjY5NSIgdzE0OnRleHRJZD0i
MzU2RjdBQTkiIHc6cnNpZFI9IjAwN0M3NTA3IiB3OnJzaWRSRGVmYXVsdD0iMDA3Qzc1MDciIHc6
cnNpZFA9IjAwN0M3NTA3Ij4KPHc6cFByPgo8dzpwU3R5bGUgdzp2YWw9Ik5vU3BhY2luZyIvPgo8
L3c6cFByPgo8dzpib29rbWFya1N0YXJ0IHc6aWQ9IjAiIHc6bmFtZT0iX0dvQmFjayIvPgo8dzpi
b29rbWFya0VuZCB3OmlkPSIwIi8+CjwvdzpwPgo8dzpwIHcxNDpwYXJhSWQ9IjEwQzE3OEE1IiB3
MTQ6dGV4dElkPSI1Q0FGQjAxOCIgdzpyc2lkUj0iMDA3Qzc1MDciIHc6cnNpZFJEZWZhdWx0PSIw
MDdDNzUwNyIgdzpyc2lkUD0iMDA3Qzc1MDciPgo8dzpwUHI+Cjx3OnBTdHlsZSB3OnZhbD0iSGVh
ZGluZzIiLz4KPC93OnBQcj4KPHc6cj4KPHc6dD5DYXJkaW86PC93OnQ+CjwvdzpyPgo8L3c6cD4K
PHc6cCB3MTQ6cGFyYUlkPSIzREYxNDJGMCIgdzE0OnRleHRJZD0iNjQxRDExNTEiIHc6cnNpZFI9
IjAwNzkxODI0IiB3OnJzaWRSRGVmYXVsdD0iMDA3Qzc1MDciIHc6cnNpZFA9IjAwN0M3NTA3Ij4K
PHc6cFByPgo8dzpwU3R5bGUgdzp2YWw9Ik5vU3BhY2luZyIvPgo8L3c6cFByPgo8dzpyPgo8dzp0
PntjYXJkaW99PC93OnQ+CjwvdzpyPgo8L3c6cD4KPHc6c2VjdFByIHc6cnNpZFI9IjAwNzkxODI0
Ij4KPHc6cGdTeiB3Onc9IjEyMjQwIiB3Omg9IjE1ODQwIi8+Cjx3OnBnTWFyIHc6dG9wPSIxNDQw
IiB3OnJpZ2h0PSIxNDQwIiB3OmJvdHRvbT0iMTQ0MCIgdzpsZWZ0PSIxNDQwIiB3OmhlYWRlcj0i
NzIwIiB3OmZvb3Rlcj0iNzIwIiB3Omd1dHRlcj0iMCIvPgo8dzpjb2xzIHc6c3BhY2U9IjcyMCIv
Pgo8dzpkb2NHcmlkIHc6bGluZVBpdGNoPSIzNjAiLz4KPC93OnNlY3RQcj4KPC93OmJvZHk+Cjwv
dzpkb2N1bWVudD4KUEsBAi0AFAAGAAgAAAAhAN+k0mxUAQAAIAUAABMAAAAAAAAAAAAAAAAAAAAA
AFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAHpEat+kAAABOAgAACwAAAAAAAAAA
AAAAAACFAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEA1mSzUe0AAAAxAwAAHAAAAAAAAAAA
AAAAAACXAgAAd29yZC9fcmVscy9kb2N1bWVudC54bWwucmVsc1BLAQItABQABgAIAAAAIQC29GeY
owYAAMkgAAAVAAAAAAAAAAAAAAAAAL4DAAB3b3JkL3RoZW1lL3RoZW1lMS54bWxQSwECLQAUAAYA
CAAAACEAU/WyygYEAABKCwAAEQAAAAAAAAAAAAAAAACUCgAAd29yZC9zZXR0aW5ncy54bWxQSwEC
LQAUAAYACAAAACEA4DaRNSsMAADTdgAADwAAAAAAAAAAAAAAAADJDgAAd29yZC9zdHlsZXMueG1s
UEsBAi0AFAAGAAgAAAAhAL3Ujb8gAQAAjwIAABQAAAAAAAAAAAAAAAAAIRsAAHdvcmQvd2ViU2V0
dGluZ3MueG1sUEsBAi0AFAAGAAgAAAAhAK9WPaS9AQAAiwUAABIAAAAAAAAAAAAAAAAAcxwAAHdv
cmQvZm9udFRhYmxlLnhtbFBLAQItABQABgAIAAAAIQAc1gQkdAEAAAMDAAARAAAAAAAAAAAAAAAA
AGAeAABkb2NQcm9wcy9jb3JlLnhtbFBLAQItABQABgAIAAAAIQAEqqJFzQEAANkDAAAQAAAAAAAA
AAAAAAAAAAMgAABkb2NQcm9wcy9hcHAueG1sUEsBAhQAFAAAAAAAAmV2UGboynqFDQAAhQ0AABEA
AAAAAAAAAAAAAIAB/iEAAHdvcmQvZG9jdW1lbnQueG1sUEsFBgAAAAALAAsAwQIAALIvAAAAAA==
"""
###
# END TEMPLATE DATA

if __name__ == "__main__":
    root = tk.Tk()
    win = GUI(root)
    win.pack()
    root.mainloop()

