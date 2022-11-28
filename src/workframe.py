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


# import testmould


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
        width, height = 580, 220
        x, y = int((self.winfo_screenwidth() - width) / 2), int((self.winfo_screenheight() - height) / 2)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))  # 大小以及位置
        self.resizable(False, False)
        self.message = Label(self, text='等待中...')
        notebook = Notebook(self, width=567, height=155)
        notebook.grid(row=0, column=0, padx=5, pady=5)
        frame1 = packdraft.PackDraft(notebook, self.message)
        frame2 = unpackdraft.UnpackDraft(notebook, self.message)
        frame3 = extittle.ExTittle(notebook, self.message)
        frame4 = exvoice.ExVoice(notebook, self.message)
        # frame5 = testmould.TestMould(notebook, self.message)
        frame6 = help.Help(notebook)
        # 适用format格式化字符串是将标签宽度限定为一个定值的好办法
        notebook.add(frame1, text='{: ^14}'.format('打包草稿'))
        notebook.add(frame2, text='{: ^14}'.format('导入草稿'))
        notebook.add(frame3, text='{: ^14}'.format('导出字幕'))
        notebook.add(frame4, text='{: ^14}'.format('导出配音'))
        # notebook.add(frame5, text='{: ^14}'.format('测试模块'))
        # TODO：增加“清除草稿”功能，删除草稿附带的所有依赖文件
        # TODO：为导出草稿增加“导出后删除原草稿”选项
        notebook.add(frame6, text='{: ^14}'.format('帮助'))

        # sticky属性表示组件的相对对齐方式，w表示西边
        self.message.grid(row=1, column=0, sticky='w')
