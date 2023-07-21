import winreg
from base64 import b64decode
from os import remove, path
from tkinter import filedialog, messagebox, DISABLED, HORIZONTAL, NORMAL, Toplevel
from tkinter.ttk import Progressbar, Combobox, Button, Label

from public import img, PathManager, is_match, get_key_locate


# Toplevel是子窗口，相当于带独立窗口边框的frame，且不需要loop，它会自动随着主窗口loop


class Initializer(Toplevel):
    message: Label
    labels = [Label]
    install_comb: Combobox
    drafts_comb: Combobox
    data_comb: Combobox
    combs = [Combobox]  # 把复选框绑定起来
    install_button: Button
    drafts_button: Button
    data_button: Button
    buttons = [Button]  # 把按钮绑定起来
    progress_bar: Progressbar
    submit_button: Button
    p = PathManager()

    def __init__(self):
        super().__init__()
        self.w = None
        self.resizable(False, False)
        # with open会自动帮我们关闭文件，不需要手动关闭
        with open('tmp.ico', 'wb') as tmp:  # wb表示以覆盖写模式、二进制方式打开文件
            tmp.write(b64decode(img))  # 通过base64的decode解码图标文件
        self.iconbitmap('tmp.ico')
        remove('tmp.ico')
        self.title('导入路径')
        width, height = 560, 255
        x, y = int((self.winfo_screenwidth() - width) / 2), int((self.winfo_screenheight() - height) / 2)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))  # 大小以及位置
        self.config(padx=5)  # 5px的边缘不至于太拥挤
        # 一些单独的组件：状态提示和进度条
        self.message = Label(self)
        self.message.grid(row=0, column=0, columnspan=3, padx=5, pady=5)
        self.progress_bar = Progressbar(self, length=540, mode='indeterminate', orient=HORIZONTAL)
        self.progress_bar.grid(row=1, column=0, columnspan=3, padx=5, pady=10)
        # 第一列标签的初始化及捆绑
        install_path_label = Label(self, text='安装路径：')
        draft_path_label = Label(self, text='草稿路径：')
        data_path_label = Label(self, text='Data路径：')
        self.labels = [install_path_label, draft_path_label, data_path_label]
        # 第二列输入框的初始化及捆绑
        self.install_comb = Combobox(self, width=52)
        self.drafts_comb = Combobox(self, width=52)
        self.data_comb = Combobox(self, width=52)
        self.combs = [self.install_comb, self.drafts_comb, self.data_comb]
        # 第三列按钮的初始化及捆绑，使用lambda表达式实现传参，按钮也需要显式声明父容器
        self.install_button = Button(self,
                                     text='选择路径',
                                     command=lambda: self.manual_search(0),
                                     )
        self.drafts_button = Button(self,
                                    text='选择路径',
                                    command=lambda: self.manual_search(1),
                                    )
        self.data_button = Button(self,
                                  text='选择路径',
                                  command=lambda: self.manual_search(2),
                                  )
        self.buttons = [self.install_button, self.drafts_button, self.data_button]
        # 批量布局标签、组合框、按钮
        for i in range(3):
            self.labels[i].grid(row=i + 2, column=0, pady=5, sticky='e')
            self.combs[i].grid(row=i + 2, column=1, pady=5)
            self.buttons[i].grid(row=i + 2, column=2, pady=5)

        # # 通过采用frame容器重新打包，完成按钮之间的对齐
        # bt_frame = Frame(self, width=255, height=20)
        # self.submit_button = Button(bt_frame, text='下一步>>>', state=DISABLED, width=30, command=self.submit)
        # self.submit_button.grid(row=0, column=1, padx=10)
        # bt_frame.grid(row=5, column=0, columnspan=3)

    def activate_button(self, state: bool):
        if state:
            for i in range(3):
                self.combs[i].config(state=NORMAL)
                self.buttons[i].config(state=NORMAL)
        else:
            for i in range(3):
                self.combs[i].config(state=DISABLED)
                self.buttons[i].config(state=DISABLED)

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

            self.message.config(text='路径提交成功！')
