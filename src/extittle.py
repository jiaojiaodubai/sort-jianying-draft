from json import load
from os import mkdir
from os.path import basename, exists
from time import strftime, localtime
from tkinter import messagebox, Checkbutton, BooleanVar

from lxml import etree

import template
from public import names2name, win32_shell_copy


class ExTittle(template.Template):
    checks: list[Checkbutton] = []
    vals: list[BooleanVar] = []
    checks_names = ['is_srt', 'is_lrc', 'is_font', 'is_remember']
    checks_names_display = ['导出SRT字幕', '导出LRC字幕', '连同字体导出', '记住导出路径']
    module_name = 'extittle'
    module_name_display = '导出字幕'
    texts = []

    def __init__(self, parent, label):
        super().__init__(parent, label)

    def analyse_meta(self, draft_path: str):
        self.texts.clear()
        transer = 60000000
        # 每次执行新项目就要清空上次的记录，否则记录会叠加
        with open('{}\\draft_content.json'.format(draft_path), 'r', encoding='utf-8') as f:
            data = load(f)
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
                                        # 微秒格式化，微秒到毫秒进率是1000，毫秒到秒进率也是1000
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

    def main_fun(self):
        for draft in self.draft_todo[names2name(self.draft_todo).index(self.draft_comb.get())]:
            have_tittle = self.analyse_meta(draft)
            if have_tittle:
                # 写入导出路径
                if self.vals[3].get() == 1:
                    self.p.configer.set('extittle_setting', 'export_path', ','.join(self.export_path))
                    self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
                # 不能使用冒号，否则OSError: [WinError 123] 文件名、目录名或卷标语法不正确
                suffix = strftime('%m.%d.%H-%M-%S', localtime())
                filepath = '{}/{}-导出的字幕-{}'.format(self.export_path[0], basename(draft), suffix)
                mkdir(filepath)
                self.write_txt(file_name=basename(draft), file_path=filepath)
                # 根据复选框确定导出操作
                if self.vals[0].get() == 1:
                    self.write_srt(file_name=basename(draft), file_path=filepath)
                if self.vals[1].get() == 1:
                    self.write_lrc(file_name=basename(draft), file_path=filepath)
                if self.vals[2].get() == 1:
                    self.write_font(file_name=basename(draft), file_path=filepath)
            else:
                messagebox.showwarning(title='缺少内容',
                                       message='你选择的草稿{}中没有找到字幕！'.format(basename(draft)))

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

    def write_lrc(self, file_name: str, file_path: str):
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
            if exists(font):
                win32_shell_copy(font, '{}\\{}-字体\\{}'.format(file_path, file_name, basename(font)))
        self.message.config(text='导出字体成功！')
