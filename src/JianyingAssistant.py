u"""
主程序窗口，将各功能模块集成到一个窗口显示。

MainWin类是实现该窗口的类，其为Tk的子类。
"""

from base64 import b64decode
from os import remove
from tkinter import Tk, messagebox
from tkinter.ttk import Notebook, Label

from psutil import process_iter

import extittle
import exvoice
import help
import packdraft
import setting
import unpackdraft
from lib import img


class MainWin(Tk):
    """
    主程序窗口。

    窗口默认大小为580x220，相对屏幕水平和垂直居中。
    窗口组件如下：
        message: tkinter.Label，显示程序工作状态。
        notebook: ttk.Notebook，分标签布局程序功能模块。
        frame?: ttk.Frame，notebook的标签。

    Attributes:
        message: tkinter.Label，用于展示程序运行状态。
    """
    message: Label

    def __init__(self):
        super().__init__()
        # 窗口名称和图标
        self.title('剪映导出助手')
        with open('tmp.ico', 'wb') as tmp:
            tmp.write(b64decode(img))
        self.iconbitmap('tmp.ico')
        remove('tmp.ico')
        # 窗口位置和大小
        width, height = 580, 220
        self.geometry('{}x{}+{}+{}'.format(
            width,
            height,
            (self.winfo_screenwidth() - width) // 2,
            (self.winfo_screenheight() - height) // 2,

        ))
        self.minsize(400, 220)
        # 窗口缩放行为
        self.resizable(True, False)
        self.grid_columnconfigure(index=0, weight=1)
        # 窗口组件
        self.message = Label(self, text='等待中...')
        self.message.grid(row=1, column=0, sticky='W')
        notebook = Notebook(self, width=567, height=155)
        notebook.grid(row=0, column=0, padx=5, pady=5, sticky='EW')
        # TODO: 面板构造完成前就县构造Init进行初始化
        frame1 = packdraft.PackDraft(notebook, self.message)
        frame2 = unpackdraft.UnpackDraft(notebook, self.message)
        frame3 = extittle.ExTittle(notebook, self.message)
        frame4 = exvoice.ExVoice(notebook, self.message)
        frame5 = setting.Setting(notebook, self.message)
        # frame5 = testmould.TestMould(notebook, self.message)
        frame6 = help.Help(notebook)
        notebook.add(frame1, text='{: ^14}'.format(frame1.module_name_display))
        notebook.add(frame2, text='{: ^14}'.format(frame2.module_name_display))
        notebook.add(frame3, text='{: ^14}'.format(frame3.module_name_display))
        notebook.add(frame4, text='{: ^14}'.format(frame4.module_name_display))
        notebook.add(frame5, text='{: ^14}'.format(frame5.module_name_display))
        # notebook.add(frame5, text='{: ^14}'.format(frame5.module_name_display))
        notebook.add(frame6, text='{: ^14}'.format('帮助'))
        # 运行环境检测
        for process in process_iter(attrs=['name']):
            if process.name() == 'JianyingPro.exe' or process.name() == 'CapCut.exe':
                messagebox.showerror(title='剪映未关闭',
                                     message='请关闭剪映再启动本程序！'
                                     )
                exit(1)
        self.mainloop()


if __name__ == '__main__':
    m = MainWin()

# TODO：找到mac创建文件夹的快捷方式（替身）的API，替换Win的Dispatch("WScript.Shell")
# TODO：在mac上测试不同的文件复制方法
