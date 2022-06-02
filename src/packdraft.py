import json
import os
import shutil
import pathlib as pl
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import public


class PackDraft(tk.Frame):
    p = public.PathManager()
    drafts_origin = []
    drafts_todo = []
    export_path = [p.DESKTOP, ]
    draft_comb: ttk.Combobox
    export_comb: ttk.Combobox
    message: tk.Label
    export_button: tk.Button
    is_zip: tk.Checkbutton
    is_share: tk.Checkbutton
    is_open: tk.Checkbutton
    is_remember: tk.Checkbutton
    val1: tk.BooleanVar
    val2: tk.BooleanVar
    val3: tk.BooleanVar
    val4: tk.BooleanVar
    shells = win32com.client.Dispatch("WScript.Shell")

    def __init__(self, parent, label):
        super().__init__(parent, width=560, height=155)
        draft_label = tk.Label(self, text='已选草稿：')
        export_label = tk.Label(self, text='导出路径：')
        draft_label.grid(row=0, column=0, pady=10, padx=5)
        export_label.grid(row=1, column=0, pady=10, padx=5)
        self.draft_comb = ttk.Combobox(self, width=51, state='readonly')
        self.draft_comb.grid(row=0, column=1, columnspan=2, pady=10, padx=5)
        self.export_comb = ttk.Combobox(self, width=51)
        self.export_comb.grid(row=1, column=1, columnspan=2, pady=10, padx=5)
        self.export_comb.config(values=self.export_path)
        self.export_comb.current(0)
        draft_choose = tk.Button(self, text='选取草稿', command=self.choose_draft)
        draft_choose.grid(row=0, column=3, pady=5, padx=5)
        export_choose = tk.Button(self, text='选择路径', command=self.choose_export)
        export_choose.grid(row=1, column=3, pady=5, padx=5)
        box_frame = tk.Frame(self, width=520, height=20)
        self.val1, self.val2, self.val3, self.val4 = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
        # 启动guide的时候就已经检查过config.ini了，因此不再执行检查
        self.p.configer.read('config.ini', encoding='utf-8')
        if self.p.configer.has_section('pack_setting'):
            if self.p.configer.has_option('pack_setting', 'is_only'):
                self.val1.set(eval(self.p.configer.get('pack_setting', 'is_only')))
            if self.p.configer.has_option('pack_setting', 'is_zip'):
                self.val2.set(eval(self.p.configer.get('pack_setting', 'is_zip')))
            if self.p.configer.has_option('pack_setting', 'is_open'):
                self.val3.set(eval(self.p.configer.get('pack_setting', 'is_open')))
            if self.p.configer.has_option('pack_setting', 'is_remember'):
                self.val4.set(eval(self.p.configer.get('pack_setting', 'is_remember')))
            if self.p.configer.has_option('pack_setting', 'export_path'):
                self.export_path.insert(0, self.p.configer.get('pack_setting', 'export_path'))
                self.export_comb.config(values=self.export_path)
                self.export_comb.current(0)
        else:
            self.p.configer.add_section('pack_setting')
            self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
        is_only = tk.Checkbutton(box_frame, text='仅导出索引', variable=self.val1)
        is_only.grid(row=0, column=0)
        is_zip = tk.Checkbutton(box_frame, text='打包为单文件', variable=self.val2)
        is_zip.grid(row=0, column=1)
        is_open = tk.Checkbutton(box_frame, text='完成后打开文件', variable=self.val3)
        is_open.grid(row=0, column=2)
        is_remember = tk.Checkbutton(box_frame, text='记住导出路径', variable=self.val4)
        is_remember.grid(row=0, column=3)
        box_frame.grid(row=2, column=0, columnspan=4)
        self.export_button = tk.Button(self, text='一键导出', padx=40,
                                       command=lambda: threading.Thread(target=self.export).start(),
                                       )
        self.export_button.grid(row=3, column=1, columnspan=2, padx=5, pady=5)
        self.message = label
        self.p.collect_draft()
        self.p.read_path()

    def choose_export(self):
        select_temp = filedialog.askdirectory(parent=self,
                                              title='请选择导出位置',
                                              initialdir=self.p.DESKTOP,
                                              )
        # 不要把工程导出到原工程路径，否则会引发困扰
        if os.path.isdir(select_temp) and select_temp not in self.p.paths[1]:
            self.export_path.insert(0, select_temp)
            self.export_comb.config(values=self.export_path)
            self.export_comb.current(0)
            self.message.config(text='导出路径选择完毕！')
        else:
            messagebox.showwarning(title='发生错误', message='请选择合适的文件夹！')

    def choose_draft(self):
        # 每次点开选择按钮都刷新草稿列表
        self.p.collect_draft()
        links_select = filedialog.askopenfilenames(parent=self,
                                                   title='请选择你想要导出的草稿',
                                                   initialdir=r'.\draft-preview',
                                                   filetypes=[('快捷方式', '*.lnk *link')]
                                                   )
        one_todo = []
        for link in links_select:
            if pl.PurePath(link).parent == pl.Path(r'.\draft-preview').resolve():
                one_todo.insert(0, self.shells.CreateShortCut(link).Targetpath)
                self.drafts_todo.insert(0, tuple(one_todo))
                self.draft_comb.config(values=public.names2name(self.drafts_todo))
                self.draft_comb.current(0)
                self.message.config(text='草稿选取完毕！')
            else:
                messagebox.showwarning(title='操作有误', message='请在默认打开的位置内选取！')
                break

    def analyse_meta(self, draft_path: str):
        # 每次执行新项目就要清空上次的记录，否则记录会叠加
        self.p.paths[3].clear()
        self.p.paths[4].clear()
        with open('{}\\draft_meta_info.json'.format(draft_path), 'r', encoding='utf-8') as f:
            data = json.load(f)
            meta_temp = data.pop('draft_materials')
            meta_dic = meta_temp[0]
            meta_list = meta_dic.pop('value')
            for item in meta_list:
                path = item.pop('file_Path')
                if os.path.exists(path):
                    # 这些路径仅用于替换，因此顺序不重要，直接用了append
                    self.p.paths[3].append(path)
                    self.p.paths[4].append(os.path.split(path)[0])
        self.p.paths[4] = list(set(self.p.paths[4]))
        self.p.write_path()
        # 必须及时关闭，否则f时刻被占用，不可替换内部资源
        f.close()

    def export(self):
        # os.path.isdir(self.export_comb.get())防止选择拿空，self.draft_comb.get() in self.todo_history防止不选直接按空
        if os.path.isdir(self.export_comb.get()) and self.draft_comb.get() in public.names2name(self.drafts_todo):
            self.p.read_path()
            self.message.config(text='正在打包...')
            # 写入导出路径
            if self.val4.get() == 1:
                self.p.configer.set('pack_setting', 'export_path', ','.join(self.export_path))
                self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
            # 写入设置项
            self.p.configer.set('pack_setting', 'is_only', str(self.val1.get()))
            self.p.configer.set('pack_setting', 'is_zip', str(self.val2.get()))
            self.p.configer.set('pack_setting', 'is_open', str(self.val3.get()))
            self.p.configer.set('pack_setting', 'is_remember', str(self.val4.get()))
            self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
            for draft in self.drafts_todo[public.names2name(self.drafts_todo).index(self.draft_comb.get())]:
                # 不能使用冒号，否则OSError: [WinError 123] 文件名、目录名或卷标语法不正确
                suffix = time.strftime('%m.%d.%H-%M-%S', time.localtime())
                filepath = '{}/{}-收集的草稿-{}'.format(self.export_path[0], os.path.basename(draft), suffix)
                self.analyse_meta(draft)
                public.win32_shell_copy(draft, '{}\\{}'.format(filepath, os.path.basename(draft)))
                public.win32_shell_copy(os.path.abspath('config.ini'), '{}\\config.ini'.format(filepath))
                # 未选中“仅导出索引，就导出素材”
                if self.val1.get() != 1:
                    # 复制的是文件，则复制后也应当是文件
                    for path in self.p.paths[3]:
                        public.win32_shell_copy(path, '{}\\meta\\{}'.format(filepath, os.path.basename(path)))
                        # TODO:这里有风险，优势会有重名文件覆盖
                if self.val2.get() == 1:
                    self.message.config(text='正在压缩单文件...')
                    shutil.make_archive(filepath, 'zip', filepath)
                    shutil.rmtree(filepath)
                    self.message.config(text='草稿文件创建成功！')
                if self.val3.get() == 1:
                    os.startfile(self.export_path[0])
            self.message.config(text='草稿打包完毕！')
        else:
            messagebox.showwarning(title='路径无效', message='请检查路径是否存在！')
