import base64
import os
import re
import time
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
from pathlib import Path
import threading
import public
import psutil

import workframe


class Guide:
    guide: tk.Tk
    message: tk.Label
    draft_comb: ttk.Combobox
    cache_comb: ttk.Combobox
    cloud_comb: ttk.Combobox
    combs = [tk.ttk.Combobox]
    draft_button: tk.Button
    cache_button: tk.Button
    cloud_button: tk.Button
    buttons = [tk.Button]
    progress_bar: ttk.Progressbar
    submit_button: tk.Button
    p = public.PathManager()

    def __init__(self):
        self.guide = tk.Tk()
        # with open会自动帮我们关闭文件，不需要手动关闭
        with open('tmp.ico', 'wb') as tmp:  # wb表示以覆盖写模式、二进制方式打开文件
            tmp.write(base64.b64decode(public.img))
        self.guide.iconbitmap('tmp.ico')
        os.remove('tmp.ico')
        self.guide.title('导入路径')
        self.guide.geometry('550x255+300+250')
        # 第一列标签的初始化及捆绑
        draft_path_label = tk.Label(self.guide, text='Draft路径：')
        cache_path_label = tk.Label(self.guide, text='Cache路径：')
        cloud_path_label = tk.Label(self.guide, text='Resources路径：')
        labels = [draft_path_label, cache_path_label, cloud_path_label]
        # 第二列输入框的初始化及捆绑
        self.draft_comb = ttk.Combobox(self.guide, width=52)
        self.cache_comb = ttk.Combobox(self.guide, width=52)
        self.cloud_comb = ttk.Combobox(self.guide, width=52)
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
        self.message = tk.Label(self.guide)
        self.message.grid(row=0, column=0, columnspan=3, padx=5, pady=10)
        self.progress_bar = ttk.Progressbar(self.guide, length=540, mode='indeterminate', orient=tk.HORIZONTAL)
        self.progress_bar.grid(row=4, column=0, columnspan=3, padx=5, pady=10)
        self.submit_button = tk.Button(self.guide, text='下一步>>>', state=tk.DISABLED, padx=40, command=self.submit)
        self.submit_button.grid(row=5, column=1)
        # 批量布局标签、组合框、按钮
        for i in range(3):
            labels[i].grid(row=i + 1, column=0, pady=5, sticky='e')
            self.combs[i].grid(row=i + 1, column=1, pady=5)
            self.buttons[i].grid(row=i + 1, column=2, pady=5)
        # 执行窗体功能
        thread = threading.Thread(target=self.is_have, args=(30,))
        thread.start()
        # 窗口基本参数
        self.guide.resizable(False, False)
        self.guide.mainloop()

    def is_have(self, speed: int = 5):
        self.update_comb(speed)
        self.message.config(text='正在读取历史记录...')
        self.progress_bar.start()
        # 检查历史记录
        if not self.p.read_path():
            if messagebox.askokcancel(title='遇到异常', message='缺少历史记录，是否使用自动搜索？'):
                self.message.config(text='正在自动搜索...')
                thread = threading.Thread(target=self.auto_search, args=(True,))
                thread.start()
                thread.join()
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
        rule = [
            [r'.*\\JianyingPro Drafts$', r'.*\\com[.]lveditor[.]draft$', ],
            [r'.*\\JianyingPro\\User Data\\Cache$', ],
            [r'.*\\JianyingPro\\(\d|[.])*\\Resources$', r'.*\\JianyingPro\\Apps\\(\d|[.])*\\Resources$']
        ]
        i = 0  # 遍历计数单位
        for item in disk_infor:
            # disk_info原来都是一个partitions对象，这里为了类型兼容，强制转换为list
            for super_path, sub_path, sub_files in os.walk(list(item)[1]):
                i += 1
                if is_show:
                    self.message.config(text='已查找{}个文件，正在检查{}...'.format(i, super_path)[:50] + '...')
                for j in range(3):
                    for key in rule[j]:
                        if re.match(key, super_path):
                            self.p.paths[j].insert(0, super_path)
                            self.p.paths[j] = list(dict.fromkeys(self.p.paths[j]))
        # 自动搜索就是依次遍历，不会出现重复，不需要去重
        self.p.write_path()

    def manual_search(self, path_name: str, position: int):
        select = False
        while not select:
            item = filedialog.askdirectory(parent=self.guide, title='选取{}路径'.format(path_name))
            if item != '':  # 防止连续按取消接空
                self.p.paths[position].insert(0, item)
            self.p.paths[position] = list(dict.fromkeys(self.p.paths[position]))
            if len(item) == 0:
                select = not messagebox.askretrycancel('未选中', '未选中任何目录')
            else:
                select = True

        self.update_comb()

    def update_comb(self, speed: int = 5):
        is_path = True
        # 更新组合框的参数及其显示
        for i in range(3):
            if len(self.p.paths[i]) > 0:
                self.combs[i].config(values=self.p.paths[i])  # 依据类更新选项
            else:
                self.combs[i].config(values=('等待输入路径...',))
            self.combs[i].current(0)  # 显示最新的选项
            is_path = is_path and Path(self.combs[i].get()).is_dir()
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
            self.guide.update()
            time.sleep(0.1)

    def submit(self):
        # self.submit_button.config(state=tk.DISABLED)
        for i in range(3):
            if Path(self.combs[i].get()).is_dir():  # 检查输入，如果是路径
                self.message.config(text='成功提交路径！')
                # 如果是新路径，还要加入到类变量
                if self.combs[i].get() in self.p.paths[i]:
                    self.p.paths[i].remove(self.combs[i].get())
                self.p.paths[i].insert(0, self.combs[i].get())
                self.update_comb(15)
                self.p.write_path()
                workframe.WorkFrame()
            else:
                # 只要有一个错就提示错误
                self.message.config(text='请确认你的输入是否完整无误！')
                break


g = Guide()
