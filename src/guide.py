import _thread
import base64
import os
import pathlib as pl
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import threading
import public
import psutil
import workframe

'''
这是本程序打开后的引导窗口，用于获取草稿路径、缓存路径、内置素材路径
https://stackoverflow.com/questions/24885827/python-tkinter-how-can-i-ensure-only-one-child-window-is-created-onclick-and-no
'''


class Guide(tk.Tk):
    message: tk.Label
    draft_comb: ttk.Combobox
    cache_comb: ttk.Combobox
    cloud_comb: ttk.Combobox
    combs = [tk.ttk.Combobox]  # 把复选框绑定起来
    draft_button: tk.Button
    cache_button: tk.Button
    cloud_button: tk.Button
    buttons = [tk.Button]  # 把按钮绑定起来
    progress_bar: ttk.Progressbar
    submit_button: tk.Button
    p = public.PathManager()

    def __init__(self):
        super().__init__()
        self.resizable(False, False)
        self.w = None
        # with open会自动帮我们关闭文件，不需要手动关闭
        with open('tmp.ico', 'wb') as tmp:  # wb表示以覆盖写模式、二进制方式打开文件
            tmp.write(base64.b64decode(public.img))  # 通过base64的decode解码图标文件
        self.iconbitmap('tmp.ico')
        os.remove('tmp.ico')
        self.title('导入路径')
        self.geometry('560x255+300+250')
        self.config(padx=5)  # 5px的边缘不至于太拥挤
        # 第一列标签的初始化及捆绑
        draft_path_label = tk.Label(self, text='Draft路径：')
        cache_path_label = tk.Label(self, text='Cache路径：')
        cloud_path_label = tk.Label(self, text='Resources路径：')
        labels = [draft_path_label, cache_path_label, cloud_path_label]
        # 第二列输入框的初始化及捆绑
        self.draft_comb = ttk.Combobox(self, width=52)
        self.cache_comb = ttk.Combobox(self, width=52)
        self.cloud_comb = ttk.Combobox(self, width=52)
        self.combs = [self.draft_comb, self.cache_comb, self.cloud_comb]
        # 第三列按钮的初始化及捆绑，使用lambda表达式实现传参
        self.draft_button = tk.Button(text='选择路径',
                                      command=lambda: self.manual_search('Draft', 0),
                                      state=tk.DISABLED)
        self.cache_button = tk.Button(text='选择路径',
                                      command=lambda: self.manual_search('Cache', 1),
                                      state=tk.DISABLED)
        self.cloud_button = tk.Button(text='选择路径',
                                      command=lambda: self.manual_search('Resources', 2),
                                      state=tk.DISABLED)
        self.buttons = [self.draft_button, self.cache_button, self.cloud_button]
        # 一些单独的组件
        self.message = tk.Label(self)
        self.message.grid(row=0, column=0, columnspan=3, padx=5, pady=10)
        self.progress_bar = ttk.Progressbar(self, length=540, mode='indeterminate', orient=tk.HORIZONTAL)
        self.progress_bar.grid(row=4, column=0, columnspan=3, padx=5, pady=10)
        # 通过采用frame容器重新打包，完成按钮之间的对齐
        bt_frame = tk.Frame(self, width=255, height=20)
        self.submit_button = tk.Button(bt_frame, text='下一步>>>', state=tk.DISABLED, width=30, command=self.submit)
        self.submit_button.grid(row=0, column=1, padx=10)
        self.auto_button = tk.Button(bt_frame, text='自动搜索', width=30,
                                     # 注意此处lambda表达式的写法，因为开启线程要求必须传入参数，因此选定了一个不痛不痒的速度参数
                                     command=lambda: _thread.start_new_thread(self.auto_search, (15, )))
        self.auto_button.grid(row=0, column=0, padx=10)
        bt_frame.grid(row=5, column=0, columnspan=3)
        # 批量布局标签、组合框、按钮
        for i in range(3):
            labels[i].grid(row=i + 1, column=0, pady=5, sticky='e')
            self.combs[i].grid(row=i + 1, column=1, pady=5)
            self.buttons[i].grid(row=i + 1, column=2, pady=5)
        pids = psutil.pids()
        for pid in pids:
            # 由于pid是动态分配的，因此通过程序名称来定位进程更合理
            if psutil.Process(pid).name() == 'JianyingPro.exe':
                messagebox.showerror(title='遇到异常', message='检测到剪映正在后台运行，\n请关闭剪映后重新启动本程序！')
                sys.exit()
        # 不要把主功能和loop加入到else子句，否则会出现意外（因为找不到loop）
        # 执行窗体功能
        thread = threading.Thread(target=self.is_have, args=(30,))
        thread.start()
        # 窗口基本参数
        self.mainloop()

    def is_have(self, speed: int = 5):
        self.update_comb(speed)  # 第一次update时为了显示历史记录或者提示语句
        self.message.config(text='正在读取历史记录...')
        self.progress_bar.start()
        # 检查历史记录
        if not self.p.read_path():
            if messagebox.askokcancel(title='遇到异常', message='缺少历史记录，是否使用自动搜索？'):
                self.message.config(text='正在自动搜索...')
                thread = threading.Thread(target=self.auto_search, args=(True,))
                thread.start()
                thread.join()  # 主线程等待子线程结束后才结束
                self.message.config(text='已完成路径检索！')
            else:
                self.message.config(text='请点击相应按钮来选取路径...')
                self.activate_button()
        # 此时只检查是否有历史记录以及时给出反馈，并解锁按钮，
        # 手动输入的值在稍后提交时再验证
        if self.update_comb(speed):
            self.message.config(text='已找到历史记录！')
            self.activate_button()

    def activate_button(self):
        for i in range(3):
            self.combs[i].config(state=tk.NORMAL)
            self.buttons[i].config(state=tk.NORMAL)
        self.submit_button.config(state=tk.NORMAL)

    def auto_search(self, is_show: bool = False):
        disk_infor = psutil.disk_partitions(all=False)
        i = 0  # 遍历计数单位
        for item in disk_infor:
            # disk_info原来都是一个partitions对象，这里为了类型兼容，强制转换为list
            for super_path, sub_path, sub_files in os.walk(list(item)[1]):
                i += 1
                if is_show:
                    # 限制为显示50个字符，太多了显示不完
                    self.message.config(text='已查找{}个文件，正在检查{}...'.format(i, super_path)[:50] + '...')
                for j in range(3):
                    if public.is_match(super_path, j):
                        self.p.paths[j].insert(0, super_path)
                        # 实现在不改变列表顺序的情况下，以原来的次序作为键值进行去重
                        self.p.paths[j] = list(dict.fromkeys(self.p.paths[j]))
        self.message.config(text='自动查找完毕')
        # 自动搜索就是依次遍历，不会出现重复，不需要去重
        self.p.write_path()

    def manual_search(self, path_name: str, position: int):
        path_call = ['Draft', 'Cache', 'Resources']
        item = filedialog.askdirectory(parent=self, title='选取{}路径'.format(path_name))
        if pl.Path(item).is_dir():
            # 采用自定义方法判断是否复合路径的特定规则
            if public.is_match(item, position):
                self.p.paths[position].insert(0, item)
                self.p.paths[position] = list(dict.fromkeys(self.p.paths[position]))
                self.update_comb()
            else:
                messagebox.showwarning(title='路径错误',
                                       message='您输入的{}路径有误，\n请重新选择！'.format(path_call[position]))
        else:
            messagebox.askretrycancel('未选中', '未选中任何目录')

    def update_comb(self, speed: int = 5):
        is_path = True
        # 更新组合框的参数及其显示
        for i in range(3):
            if len(self.p.paths[i]) > 0:
                self.combs[i].config(values=self.p.paths[i])  # 依据类更新选项
            else:
                self.combs[i].config(values=('等待输入路径...',))
            self.combs[i].current(0)  # 显示最新的选项
            # Python似乎会把空字符串视为路径，因此要加以排除
            is_path = is_path and pl.Path(self.combs[i].get()).is_dir() and self.combs[i].get() != ''
        # 路经检查全部通过则给个进度提示
        if is_path:
            self.progress_bar.stop()
            self.ok_progress(speed)
            return True
        else:
            return False

    # 秀一下进度条给个安慰
    def ok_progress(self, speed: int = 5):
        # 变为单向进度条
        self.progress_bar.config(mode="determinate")
        for i in range(100 // speed + 1):
            self.progress_bar['value'] += speed
            self.update()
            time.sleep(0.1)

    def submit(self):
        path_call = ['Draft', 'Cache', 'Resources']
        is_correct = True
        for i in range(3):
            now_path = self.combs[i].get()
            if pl.Path(now_path).is_dir():  # 检查输入，如果是路径
                is_correct = is_correct and public.is_match(now_path, i)
                if public.is_match(now_path, i):
                    # 如果是新路径，还要加入到类变量
                    if self.combs[i].get() in self.p.paths[i]:
                        self.p.paths[i].remove(self.combs[i].get())
                    self.p.paths[i].insert(0, self.combs[i].get())
                else:
                    # 只要有一个错就提示错误
                    messagebox.showwarning(title='路径有误',
                                           message='您输入的{}路径有误，\n请重新选择！'.format(path_call[i]))
                    break
            else:
                # 只要有一个错就提示错误
                messagebox.showwarning(title='路径有误',
                                       message='请确认你输入的{}路径是否存在！'.format(path_call[i]))
                break
        if is_correct:
            self.update_comb(15)
            self.p.write_path()
            self.w = workframe.WorkFrame(self)
            # 以下代码的作用是使workframe窗口仅有一个，且打开workframe时不允许过多操作主窗口
            self.w.transient(self)
            self.w.focus_set()
            self.w.grab_set()
            self.wait_window(self.w)
            self.message.config(text='路径提交成功！')

# g = Guide()
