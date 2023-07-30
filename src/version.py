from base64 import b64decode
from os import remove
from tkinter import Toplevel, IntVar, Tk
from tkinter.ttk import Radiobutton, LabelFrame, Button

from lib import img


class Top(Toplevel):
    version_bt_cn: Radiobutton
    version_bt_it: Radiobutton
    version: IntVar

    def __init__(self, master, val: int = 0):
        super().__init__(master=master)
        self.version = IntVar()
        self.version.set(value=val)
        self.resizable(False, False)
        # with open会自动帮我们关闭文件，不需要手动关闭
        with open('tmp.ico', 'wb') as tmp:  # wb表示以覆盖写模式、二进制方式打开文件
            tmp.write(b64decode(img))  # 通过base64的decode解码图标文件
        self.iconbitmap('tmp.ico')
        remove('tmp.ico')
        self.title('设置当前版本')
        width, height = 300, 140
        x, y = int((self.winfo_screenwidth() - width) / 2), int((self.winfo_screenheight() - height) / 2)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))  # 大小以及位置
        # 不要把主功能和loop加入到else子句，否则会出现意外（因为找不到loop）
        self.columnconfigure(index=0, weight=1)
        version_fr = LabelFrame(self, text='选择你的剪映发行版本', width=560, padding=5)
        version_fr.grid(row=0, column=0, sticky='EW', padx=10)
        version_fr.columnconfigure(index=0, weight=1)
        version_fr.columnconfigure(index=1, weight=1)
        self.version_bt_cn = Radiobutton(master=version_fr, text='中文版', variable=self.version, value=0)
        self.version_bt_cn.grid(row=0, column=0)
        self.version_bt_it = Radiobutton(master=version_fr, text='国际版', variable=self.version, value=1)
        self.version_bt_it.grid(row=0, column=1)
        version_fr.grid(row=0, column=0, sticky='EW', padx=5, pady=20)
        version_set = Button(master=self, text='确定', command=self.submit)
        version_set.grid(row=1, column=0)
        self.protocol("WM_DELETE_WINDOW", self.close)

    def submit(self):
        self.destroy()
        self.master.destroy()

    def close(self):
        self.destroy()
        self.master.destroy()

    def get_var(self):
        # https://stackoverflow.com/questions/28443749/how-do-i-return-a-result-from-a-dialog
        self.wait_window()
        s = ('JianyingPro', 'CapCut')[self.version.get()]
        return s


def SetVersion():
    m = Tk()
    m.withdraw()
    t = Top(m)
    t.focus_set()
    t.grab_set()
    data = t.get_var()
    m.mainloop()
    return data


if __name__ == '__main__':
    print(SetVersion())
