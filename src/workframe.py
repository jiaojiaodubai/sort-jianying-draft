from base64 import b64decode
from os import remove
from tkinter import Toplevel, Label
from tkinter.ttk import Notebook
import extittle
import exvoice
import help
import packdraft
import unpackdraft
from public import img


# import winshell
# from PIL import Image
# def cal_corner(x: int, y: int):
#     if x < y:
#         corner = (0, (y - x) // 2, x, ((y - x) // 2) + x)
#     else:
#         corner = ((x - y) // 2, 0, ((x - y) // 2) + y, y)
#     return corner

# 注意，Toplevel是子窗口，相当于带独立窗口边框的frame，且不需要loop，它会自动随着主窗口loop
class WorkFrame(Toplevel):
    message: Label

    # noinspection SpellCheckingInspection
    def __init__(self, master):
        super().__init__(master)
        self.title('剪映草稿助手')
        with open('tmp.ico', 'wb') as tmp:
            tmp.write(b64decode(img))
        self.iconbitmap('tmp.ico')
        remove('tmp.ico')
        self.geometry('580x220+300+250')
        self.resizable(False, False)
        self.message = Label(self, text='等待中...')
        notebook = Notebook(self, width=560, height=155)
        notebook.grid(row=0, column=0, padx=5, pady=5)
        frame1 = packdraft.PackDraft(notebook, self.message)
        frame2 = unpackdraft.UnpackDraft(notebook, self.message)
        frame3 = extittle.ExTittle(notebook, self.message)
        frame4 = exvoice.ExVoice(notebook, self.message)
        frame5 = help.Help(notebook)
        # 适用format格式化字符串是将标签宽度限定为一个定值的好办法
        notebook.add(frame1, text='{: ^19}'.format('打包草稿'))
        notebook.add(frame2, text='{: ^19}'.format('导入草稿'))
        notebook.add(frame3, text='{: ^19}'.format('导出字幕'))
        notebook.add(frame4, text='{: ^19}'.format('导出配音'))
        notebook.add(frame5, text='{: ^19}'.format('帮助'))
        # sticky属性表示组件的相对对齐方式，w表示西边
        self.message.grid(row=1, column=0, sticky='w')
        # self.mainloop()
