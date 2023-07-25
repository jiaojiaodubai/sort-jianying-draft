from json import load
from os import startfile
from os.path import basename
from time import strftime, localtime
from tkinter import messagebox, Checkbutton, BooleanVar

import template
from public import names2name, win32_shell_copy


class ExVoice(template.Template):
    checks: list[Checkbutton] = []
    vals: list[BooleanVar] = []
    checks_names = ['is_open', 'is_remember']
    checks_names_display = ['完成后打开文件', '记住导出路径']
    module_name = 'exvoice'
    module_name_display = '导出配音'
    audios = []

    def __init__(self, parent, label):
        super().__init__(parent, label)

    def analyse_meta(self, draft_path: str):
        self.audios.clear()
        with open('{}\\draft_content.json'.format(draft_path), 'r', encoding='utf-8') as f:
            data = load(f)
            materials = data.get('materials')
            audios = materials.get('audios')
            i = 0
            has_audio = False
            for item in audios:
                if item.get('type') == 'text_to_audio':
                    has_audio = True
                    i += 1
                    a_audio = {'serial': i,
                               'txt': item.get('name'),
                               'path': '{}\\textReading\\{}'.format(draft_path, item.get('path').split('/')[-1]),
                               'tune': item.get('tone_type')
                               }
                    self.audios.append(a_audio)
        # 必须及时关闭，否则f时刻被占用，不可替换内部资源
        f.close()
        return has_audio

    def main_fun(self):
        # 写入导出路径
        if self.vals[1].get() == 1:
            self.p.configer.set('exvoice_setting', 'export_path', ','.join(self.export_path))
            self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
        for draft in self.draft_todo[names2name(self.draft_todo).index(self.draft_comb.get())]:
            have_audio = self.analyse_meta(draft)
            if have_audio:
                # 不能使用冒号，否则OSError: [WinError 123] 文件名、目录名或卷标语法不正确
                suffix = strftime('%m.%d.%H-%M-%S', localtime())
                filepath = '{}\\{}-收集的配音-{}'.format(self.export_path[0], basename(draft), suffix)
                for audio in self.audios:
                    win32_shell_copy(audio.get('path'),
                                     '{}\\{}-{}-{}.wav'.format(filepath,
                                                               # 序号既能防止重复，也能标示先后顺序
                                                               audio.get('serial'),
                                                               audio.get('txt'),
                                                               audio.get('tune')
                                                               )
                                     )

                if self.vals[0].get() == 1:
                    startfile(self.export_path[0])
            else:
                messagebox.showwarning(title='缺少内容',
                                       message='你选择的草稿{}中没有找到配音！'.format(basename(draft)))
