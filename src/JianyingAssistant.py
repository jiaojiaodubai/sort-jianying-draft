# encoding = utf-8
"""
主程序
"""

from base64 import b64decode
from os import remove
from tkinter import Tk, messagebox
from tkinter.ttk import Notebook, Label

from psutil import process_iter

from public import Initializer, img
from src import help, packdraft, testmould, unpackdraft, extittle, exvoice


class MainWin(Tk):
    p = Initializer()
    message: Label

    def __init__(self):
        super().__init__()
        self.title('剪映导出助手')
        with open('tmp.ico', 'wb') as tmp:
            tmp.write(b64decode(img))
        self.iconbitmap('tmp.ico')
        remove('tmp.ico')
        width, height = 580, 220
        x, y = int((self.winfo_screenwidth() - width) / 2), int((self.winfo_screenheight() - height) / 2)
        self.minsize(400, 220)
        # self.maxsize(1200, 220)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))  # 大小以及位置
        self.resizable(True, False)
        self.message = Label(self, text='等待中...')
        notebook = Notebook(self, width=567, height=155)
        notebook.grid(row=0, column=0, padx=5, pady=5, sticky='EW')
        self.grid_columnconfigure(index=0, weight=1, minsize=300)
        # notebook.pack(padx=5, pady=5, )
        frame1 = packdraft.PackDraft(notebook, self.message)
        frame2 = unpackdraft.UnpackDraft(notebook, self.message)
        frame3 = extittle.ExTittle(notebook, self.message)
        frame4 = exvoice.ExVoice(notebook, self.message)
        frame5 = testmould.TestMould(notebook, self.message)
        frame6 = help.Help(notebook)
        # 适用format格式化字符串是将标签宽度限定为一个定值的好办法
        notebook.add(frame1, text='{: ^14}'.format(frame1.module_name_display))
        notebook.add(frame2, text='{: ^14}'.format(frame2.module_name_display))
        notebook.add(frame3, text='{: ^14}'.format(frame3.module_name_display))
        notebook.add(frame4, text='{: ^14}'.format(frame4.module_name_display))
        notebook.add(frame5, text='{: ^14}'.format(frame5.module_name_display))
        notebook.add(frame6, text='{: ^14}'.format('帮助'))

        # sticky属性表示组件的相对对齐方式，w表示西边
        self.message.grid(row=1, column=0, sticky='W')
        # self.w = initializer.Initializer()
        #
        # # 禁止与下层窗口交互
        # # https://python.tutorialink.com/how-to-make-pop-up-window-with-force-attention-in-tkinter/
        # self.w.transient(self)
        # self.w.focus_set()
        # self.w.protocol("WM_DELETE_WINDOW", self.close)
        # self.attributes('-disabled', True)

        for process in process_iter(attrs=['name']):
            if process.name() == 'JianyingPro.exe':
                messagebox.showerror(title='剪映未关闭',
                                     message='请关闭剪映再启动本程序！'
                                     )
                exit(1)
        self.mainloop()

    # def close(self):
    #     self.attributes('-disabled', False)
    #     self.w.destroy()


# 主类仅作为测试入口
if __name__ == '__main__':
    m = MainWin()

# TODO：找到mac创建文件夹的快捷方式（替身）的API，替换Win的Dispatch("WScript.Shell")
# TODO：在mac上测试不同的文件复制方法，看是否有进度提示，如果没有，设置进度等待界面。
