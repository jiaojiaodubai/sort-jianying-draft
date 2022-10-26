import winreg
from _thread import start_new_thread
from base64 import b64decode
from os import remove, walk, path
from sys import exit
from tkinter import Tk, Frame, Label, Button, filedialog, messagebox, DISABLED, HORIZONTAL, NORMAL
from tkinter.ttk import Progressbar, Combobox

# from threading import Thread
from psutil import pids, disk_partitions, Process

import workframe
from public import img, PathManager, is_match, get_key_locate

'''
这是本程序打开后的引导窗口，用于获取草稿路径、缓存路径、内置素材路径
'''


class Guide(Tk):
    message: Label
    draft_comb: Combobox
    cache_comb: Combobox
    cloud_comb: Combobox
    combs = [Combobox]  # 把复选框绑定起来
    draft_button: Button
    cache_button: Button
    cloud_button: Button
    buttons = [Button]  # 把按钮绑定起来
    progress_bar: Progressbar
    submit_button: Button
    p = PathManager()

    def __init__(self):
        super().__init__()
        self.resizable(False, False)
        # 为了适用窗口退出而设置的代码
        # https://stackoverflow.com/questions/24885827/python-tkinter-how-can-i-ensure-only-one-child-window-is-created-onclick-and-no
        self.w = None
        # with open会自动帮我们关闭文件，不需要手动关闭
        with open('tmp.ico', 'wb') as tmp:  # wb表示以覆盖写模式、二进制方式打开文件
            tmp.write(b64decode(img))  # 通过base64的decode解码图标文件
        self.iconbitmap('tmp.ico')
        remove('tmp.ico')
        self.title('导入路径')
        self.geometry('560x255+300+250')
        self.config(padx=5)  # 5px的边缘不至于太拥挤
        # 第一列标签的初始化及捆绑
        install_path_label = Label(self, text='安装路径：')
        draft_path_label = Label(self, text='草稿路径：')
        data_path_label = Label(self, text='Data路径：')
        self.labels = [install_path_label, draft_path_label, data_path_label]
        # 第二列输入框的初始化及捆绑
        self.draft_comb = Combobox(self, width=52)
        self.cache_comb = Combobox(self, width=52)
        self.cloud_comb = Combobox(self, width=52)
        self.combs = [self.draft_comb, self.cache_comb, self.cloud_comb]
        # 第三列按钮的初始化及捆绑，使用lambda表达式实现传参
        self.draft_button = Button(text='选择路径',
                                   command=lambda: self.manual_search(0),
                                   )
        self.cache_button = Button(text='选择路径',
                                   command=lambda: self.manual_search(1),
                                   )
        self.cloud_button = Button(text='选择路径',
                                   command=lambda: self.manual_search(2),
                                   )
        self.buttons = [self.draft_button, self.cache_button, self.cloud_button]
        # 一些单独的组件：状态提示和进度条
        self.message = Label(self)
        self.message.grid(row=0, column=0, columnspan=3, padx=5, pady=10)
        self.progress_bar = Progressbar(self, length=540, mode='indeterminate', orient=HORIZONTAL)
        self.progress_bar.grid(row=4, column=0, columnspan=3, padx=5, pady=10)
        # 通过采用frame容器重新打包，完成按钮之间的对齐
        bt_frame = Frame(self, width=255, height=20)
        self.submit_button = Button(bt_frame, text='下一步>>>', state=DISABLED, width=30, command=self.submit)
        self.submit_button.grid(row=0, column=1, padx=10)
        self.auto_button = Button(bt_frame, text='全盘搜索',
                                  width=30,
                                  # 注意此处lambda表达式的写法，因为开启线程要求必须传入参数，因此选定了一个不痛不痒的速度参数
                                  command=lambda: start_new_thread(self.auto_search, (50,)))
        self.auto_button.grid(row=0, column=0, padx=10)
        bt_frame.grid(row=5, column=0, columnspan=3)
        # 批量布局标签、组合框、按钮
        for i in range(3):
            self.labels[i].grid(row=i + 1, column=0, pady=5, sticky='e')
            self.combs[i].grid(row=i + 1, column=1, pady=5)
            self.buttons[i].grid(row=i + 1, column=2, pady=5)
        self.activate_button(False)
        self.check_env()
        # 不要把主功能和loop加入到else子句，否则会出现意外（因为找不到loop）
        # 执行窗体功能
        self.auto_get()
        # thread = Thread(target=self.is_have, args=(30,))
        # thread.start()
        # 窗口基本参数
        self.mainloop()

    @staticmethod
    def check_env():
        _pids = pids()
        for pid in _pids:
            # 由于pid是动态分配的，因此通过程序名称来定位进程更合理
            if Process(pid).name() == 'JianyingPro.exe':
                messagebox.showerror(title='遇到异常', message='检测到剪映正在后台运行，\n请关闭剪映后重新启动本程序！')
                exit()

    # def is_have(self, speed: int = 5):
    #     self.update_comb(speed)  # 第一次update时为了显示历史记录或者提示语句
    #     self.message.config(text='正在读取历史记录...')
    #     self.progress_bar.start()
    #     # 检查历史记录
    #     if not self.p.read_path():
    #         if messagebox.askokcancel(title='遇到异常', message='缺少历史记录，是否使用自动获取？'):
    #             self.message.config(text='正在自动搜索...')
    #             thread = Thread(target=self.auto_search, args=(True,))
    #             thread.start()
    #             thread.join()  # 主线程等待子线程结束后才结束
    #             self.message.config(text='已完成路径检索！')
    #         else:
    #             self.message.config(text='请点击相应按钮来选取路径...')
    #             self.activate_button()
    #     # 此时只检查是否有历史记录以及时给出反馈，并解锁按钮，
    # 手动输入的值在稍后提交时再验证
    # if self.update_comb(5):
    #     self.message.config(text='已找到历史记录！')
    #     self.activate_button()

    def activate_button(self, state: bool):
        if state:
            for i in range(3):
                self.combs[i].config(state=NORMAL)
                self.buttons[i].config(state=NORMAL)
            self.submit_button.config(state=NORMAL)
            self.auto_button.config(state=NORMAL)
        else:
            for i in range(3):
                self.combs[i].config(state=DISABLED)
                self.buttons[i].config(state=DISABLED)
            self.submit_button.config(state=DISABLED)
            self.auto_button.config(state=DISABLED)

    def auto_get(self):
        self.p.paths[0].insert(0, get_key_locate(winreg.HKEY_CURRENT_USER,
                                                 r'Software\Bytedance\JianyingPro',
                                                 'installDir').strip('\\'))
        self.p.paths[1].insert(0, get_key_locate(winreg.HKEY_CURRENT_USER,
                                                 r'Software\Bytedance\JianyingPro\GlobalSettings\History',
                                                 'currentCustomDraftPath'))
        # https://www.tutorialspoint.com/How-to-find-the-real-user-home-directory-using-Python
        self.p.paths[2].insert(0, fr'{path.expanduser("~")}\AppData\Local\JianyingPro\User Data')
        self.p.write_path()
        if self.update_comb(5):
            self.message.config(text='已找到历史记录！')
            self.activate_button(True)

    def auto_search(self, char_len: int):
        self.activate_button(False)
        disk_infor = disk_partitions(all=False)
        i = 0  # 遍历计数单位
        for item in disk_infor:
            # disk_info原来都是一个partitions对象，这里为了类型兼容，强制转换为list
            for super_path, sub_path, sub_files in walk(list(item)[1]):
                i += 1
                self.message.config(text='已查找{}个文件，正在检查{}...'.format(i, super_path)[:char_len] + '...')
                for j in range(3):
                    if is_match(super_path, j):
                        self.p.paths[j].insert(0, super_path)
                        # 实现在不改变列表顺序的情况下，以原来的次序作为键值进行去重
        self.update_comb()
        self.message.config(text='全盘搜索完毕!（仅提供可能的路径）')
        self.activate_button(True)
        # 全盘搜索就是依次遍历，不会出现重复，不需要去重
        self.p.write_path()

    def manual_search(self, position: int):
        # filedialog返回的是posix格式的地址字符串，为了统一，这里用os.path.abspath改写为Windows风格
        item = path.abspath(filedialog.askdirectory(parent=self, title=f'选取{self.p.path_names_display[position]}'))
        if path.isdir(item):
            self.combs[position].set(item)
        else:
            messagebox.askretrycancel('未选中', '未选中任何目录')

    def update_comb(self, speed: int = 5):
        is_path = True
        # 更新组合框的参数及其显示
        for i in range(3):
            if len(self.p.paths[i]) > 0:
                self.combs[i].config(values=self.p.paths[i])  # 依据类更新选项
            else:
                # 当该项目记录为空时显示
                self.combs[i].config(values=('等待输入路径...',))
            self.combs[i].current(0)  # 显示最新的选项
            # Python似乎会把空字符串视为路径，因此要加以排除
            is_path = is_path and path.isdir(self.combs[i].get()) and self.combs[i].get() != ''
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

    def submit(self):
        # 提交前检查路径的合法性和有效性
        is_correct = True
        for i in range(3):
            now_path = self.combs[i].get()
            if path.isdir(now_path):  # 检查输入，如果是路径
                is_correct = is_correct and is_match(now_path, i)
                if is_match(now_path, i):
                    self.p.paths[i].insert(0, self.combs[i].get())
                else:
                    # 只要有一个不匹配路径规则就提示错误
                    # https://zhuanlan.zhihu.com/p/491112809
                    messagebox.showwarning(title='路径有误',
                                           message=f'您输入的【{self.p.path_names_display[i]}】有误，\n请重新选择！')
                    break
            else:
                is_correct = False
                # 只要有一个不是路径就提示错误
                messagebox.showwarning(title='路径有误',
                                       message=f'请确认你输入的【{self.p.path_names_display[i]}】是否存在！')
                break
        if is_correct:
            self.p.write_path()
            # 这时才定义这个组件是Workframe
            self.w = workframe.WorkFrame(self)
            # 以下代码的作用是使workframe窗口仅有一个，且打开workframe时不允许过多操作主窗口
            self.w.transient(self)
            self.w.focus_set()
            self.w.grab_set()
            self.wait_window(self.w)
            self.message.config(text='路径提交成功！')


g = Guide()
