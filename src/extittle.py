import json
import os
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
from lxml import etree
import public


class ExTittle(tk.Frame):
    p = public.PathManager()
    drafts_origin = []
    drafts_todo = []
    export_path = [p.DESKTOP, ]
    texts = []
    draft_comb: ttk.Combobox
    export_comb: ttk.Combobox
    message: tk.Label
    export_button: tk.Button
    is_zip: tk.Checkbutton
    is_share: tk.Checkbutton
    is_lrc: tk.Checkbutton
    is_font: tk.Checkbutton
    val1: tk.BooleanVar
    val2: tk.BooleanVar
    val3: tk.BooleanVar
    val4: tk.BooleanVar
    shells = win32com.client.Dispatch("WScript.Shell")

    def __init__(self, parent, label):
        super().__init__(parent, width=520, height=150)
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
        if self.p.configer.has_section('extittle_setting'):
            if self.p.configer.has_option('extittle_setting', 'is_srt'):
                self.val1.set(eval(self.p.configer.get('extittle_setting', 'is_srt')))
            if self.p.configer.has_option('extittle_setting', 'is_zip'):
                self.val2.set(eval(self.p.configer.get('extittle_setting', 'is_zip')))
            if self.p.configer.has_option('extittle_setting', 'is_lrc'):
                self.val3.set(eval(self.p.configer.get('extittle_setting', 'is_lrc')))
            if self.p.configer.has_option('extittle_setting', 'is_remember'):
                self.val4.set(eval(self.p.configer.get('extittle_setting', 'is_remember')))
            if self.p.configer.has_option('extittle_setting', 'export_path'):
                self.export_path.insert(0, self.p.configer.get('pack_setting', 'export_path'))
                self.export_comb.config(values=self.export_path)
                self.export_comb.current(0)
        else:
            self.p.configer.add_section('extittle_setting')
            self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
        is_srt = tk.Checkbutton(box_frame, text='导出SRT字幕', variable=self.val1)
        is_srt.grid(row=0, column=0)
        is_lrc = tk.Checkbutton(box_frame, text='导出LRC歌词', variable=self.val2)
        is_lrc.grid(row=0, column=1)
        is_font = tk.Checkbutton(box_frame, text='连同字体导出', variable=self.val3)
        is_font.grid(row=0, column=2)
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
        self.p.collect_draft()
        links_select = filedialog.askopenfilenames(parent=self,
                                                   title='请选择你想要导出的草稿',
                                                   initialdir='draft-preview',
                                                   filetypes=[('快捷方式', '*.lnk *link')]
                                                   )
        one_todo = []
        for link in links_select:
            # 直接split出来的斜杠方向和abspath出来的不一样，需要统一化
            if os.path.abspath(os.path.split(link)[0]) == os.path.abspath('../draft-preview'):
                one_todo.insert(0, self.shells.CreateShortCut(link).Targetpath)
                self.drafts_todo.insert(0, tuple(one_todo))
                self.draft_comb.config(values=public.names2name(self.drafts_todo))
                self.draft_comb.current(0)
                self.message.config(text='草稿选取完毕！')
            else:
                messagebox.showwarning(title='操作有误', message='你选择的文件有误！')
                break

    def analyse_meta(self, draft_path: str):
        self.texts.clear()
        transer = 60000000
        # 每次执行新项目就要清空上次的记录，否则记录会叠加
        with open('{}\\draft_content.json'.format(draft_path), 'r', encoding='utf-8') as f:
            data = json.load(f)
            materials = data.get('materials')
            texts_temp = materials.get('texts')
            i = 1
            has_tittle = False
            for row_dic in texts_temp:
                dic_type = row_dic.get('type')
                if dic_type in ['subtitle', 'lyrics']:
                    has_tittle = True
                    dict_id = row_dic.get('id')
                    simp_dic = {'serial': i,
                                'content': row_dic.get('content'),
                                'font_path': row_dic.get('font_path'),
                                }
                    for track in data.get('tracks'):
                        if track.get('type') == 'text':
                            for seg in track.get('segments'):
                                if seg.get('material_id') == dict_id:
                                    st = int(seg.get('target_timerange').get('start'))
                                    du = int(seg.get('target_timerange').get('duration'))
                                    ed = st + du
                                    simp_dic['start'] = '{:02d},{:02d},{:02d},{:02d}'.format(
                                        st // (60 * transer),
                                        st % (60 * transer) // transer,
                                        st % (60 * transer) % transer // 1000000,
                                        st % (60 * transer) % transer % 1000000 // 1000
                                    ).split(',')
                                    simp_dic['end'] = '{:02d},{:02d},{:02d},{:02d}'.format(
                                        ed // (60 * transer),
                                        ed % (60 * transer) // transer,
                                        ed % (60 * transer) % transer // 1000000,
                                        ed % (60 * transer) % transer % 1000000 // 1000
                                    ).split(',')
                    self.texts.append(simp_dic)
                    i += 1
        # 必须及时关闭，否则f时刻被占用，不可替换内部资源
        f.close()
        return has_tittle

    def export(self):
        # os.path.isdir(self.export_comb.get())防止选择拿空，self.draft_comb.get() in self.todo_history防止不选直接按空
        if os.path.isdir(self.export_comb.get()) and self.draft_comb.get() in public.names2name(self.drafts_todo):
            self.p.read_path()
            self.message.config(text='正在导出...')
            for draft in self.drafts_todo[public.names2name(self.drafts_todo).index(self.draft_comb.get())]:
                have_tittle = self.analyse_meta(draft)
                if have_tittle:
                    # 写入导出路径
                    if self.val4.get() == 1:
                        self.p.configer.set('extittle_setting', 'export_path', ','.join(self.export_path))
                        self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
                    # 写入设置项
                    self.p.configer.set('extittle_setting', 'is_srt', str(self.val1.get()))
                    self.p.configer.set('extittle_setting', 'is_lrc', str(self.val2.get()))
                    self.p.configer.set('extittle_setting', 'is_font', str(self.val3.get()))
                    self.p.configer.set('extittle_setting', 'is_remember', str(self.val4.get()))
                    self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
                    # 不能使用冒号，否则OSError: [WinError 123] 文件名、目录名或卷标语法不正确
                    suffix = time.strftime('%m.%d.%H-%M-%S', time.localtime())
                    filepath = '{}/{}-导出的字幕-{}'.format(self.export_path[0], os.path.basename(draft), suffix)
                    os.mkdir(filepath)
                    self.write_txt(file_name=os.path.basename(draft), file_path=filepath)
                    # 写入导出路径
                    if self.val1.get() == 1:
                        self.write_srt(file_name=os.path.basename(draft), file_path=filepath)
                    if self.val2.get() == 1:
                        self.write_lrc(file_name=os.path.basename(draft), file_path=filepath)
                    if self.val3.get() == 1:
                        self.write_font(file_name=os.path.basename(draft), file_path=filepath)
                else:
                    messagebox.showwarning(title='缺少内容',
                                           message='你选择的草稿{}中没有找到字幕！'.format(os.path.basename(draft)))
        else:
            messagebox.showwarning(title='路径无效', message='请检查路径是否存在！')

    def write_txt(self, file_name: str, file_path: str):
        self.message.config(text='正在导出纯文本文件...')
        txt = []
        for dic in self.texts:
            txt.append(etree.HTML(text=dic.get('content')).xpath('string(.)').strip('[').strip(']'))
        with open('{}\\{}-字幕.txt'.format(file_path, file_name), 'w', encoding='utf-8') as f:
            s = '\n'.join(txt)
            f.write(s)
        f.close()
        self.message.config(text='纯文本文件导出成功！')

    def write_srt(self, file_name: str, file_path: str):
        self.message.config(text='正在导出st字幕文件...')
        txt = []
        for dic in self.texts:
            st = dic.get('start')
            ed = dic.get('end')
            s = '{}\n{} --> {}\n{}\n'.format(
               dic.get('serial'),
               '{}:{}:{},{}'.format(st[0], st[1], st[2], st[3]),
               '{}:{}:{},{}'.format(ed[0], ed[1], ed[2], ed[3]),
               dic.get('content')
            )
            txt.append(s)
        with open('{}\\{}-字幕.srt'.format(file_path, file_name), 'w', encoding='utf-8') as f:
            s = '\n'.join(txt)
            f.write(s)
        f.close()
        self.message.config(text='srt字幕文件导出成功！')

    def write_lrc(self,  file_name: str, file_path: str):
        self.message.config(text='正在导出lrc歌词文件...')
        txt = []
        for dic in self.texts:
            st = dic.get('start')
            s = '[{}]{}'.format(
                '{}:{}.{:.2}'.format(st[1], st[2], st[3]),
                etree.HTML(text=dic.get('content')).xpath('string(.)')
            )
            txt.append(s)
        with open('{}\\{}-字幕.lrc'.format(file_path, file_name), 'w', encoding='utf-8') as f:
            s = '\n'.join(txt)
            f.write(s)
        f.close()
        self.message.config(text='lrc歌词文件导出成功！...')

    def write_font(self, file_name, file_path: str):
        self.message.config(text='正在导出字体...')
        font_paths = set()
        for dic in self.texts:
            font_paths.add(dic.get('font_path'))

        for font in font_paths:
            if os.path.exists(font):
                public.win32_shell_copy(font, '{}\\{}-字体\\{}'.format(file_path, file_name, os.path.basename(font)))
        self.message.config(text='导出字体成功！')
