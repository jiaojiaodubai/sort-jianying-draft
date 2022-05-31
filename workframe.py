import base64
import os
import tkinter as tk
from tkinter import ttk

import extittle
import exvoice
import help
import packdraft
import public as pb


# import winshell
# from PIL import Image
# def cal_corner(x: int, y: int):
#     if x < y:
#         corner = (0, (y - x) // 2, x, ((y - x) // 2) + x)
#     else:
#         corner = ((x - y) // 2, 0, ((x - y) // 2) + y, y)
#     return corner
import unpackdraft


class WorkFrame:
    mainWin: tk.Tk
    message: tk.Label

    # noinspection SpellCheckingInspection
    def __init__(self):
        self.mainWin = tk.Tk()
        self.mainWin.title('剪映草稿助手')
        with open('tmp.ico', 'wb') as tmp:
            tmp.write(base64.b64decode(pb.img))
        self.mainWin.iconbitmap('tmp.ico')
        os.remove('tmp.ico')
        self.mainWin.geometry('580x215+300+250')
        self.mainWin.resizable(False, False)
        self.message = tk.Label(self.mainWin, text='等待中...')
        notebook = ttk.Notebook(self.mainWin)
        notebook.grid(row=0, column=0, padx=5, pady=5)
        # notebook.grid(row=0, column=0, padx=5, pady=5)
        frame1 = packdraft.PackDraft(notebook, self.message)
        frame2 = unpackdraft.UnpackDraft(notebook, self.message)
        frame3 = extittle.ExTittle(notebook, self.message)
        frame4 = exvoice.ExVoice(notebook, self.message)
        frame5 = help.Help(notebook)
        notebook.add(frame1, text='{: ^19}'.format('打包草稿'))
        notebook.add(frame2, text='{: ^19}'.format('导入草稿'))
        notebook.add(frame3, text='{: ^19}'.format('导出字幕'))
        notebook.add(frame4, text='{: ^19}'.format('导出配音'))
        notebook.add(frame5, text='{: ^19}'.format('帮助'))
        self.message.grid(row=1, column=0, sticky='w')
        self.mainWin.mainloop()
