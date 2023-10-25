#
# File Dialog 호출하여 Path를 반환한다.
#
# 231019. Focus 받도록 수정

import tkinter as tk
import tkinter.filedialog as fd
import os

def askopenfilename() -> str :
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    path = fd.askopenfilename(
        parent=root,
        initialdir=os.getcwd())
    root.destroy()
    return path

def askdirectory() -> str :
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    path = fd.askdirectory(
        parent=root,
        initialdir=os.getcwd())
    root.destroy()
    return path    

