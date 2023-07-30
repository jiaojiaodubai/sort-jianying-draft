from json import load
from os import startfile
from os.path import basename, isdir
from time import strftime, localtime
from tkinter import messagebox, Checkbutton, BooleanVar

from lib import DESKTOP, win32_shell_copy, names2name
from public import PathX, Template


class ExVoice(Template):
    # 模块级属性
    module_name = 'ex_voice'
    module_name_display = '导出配音'
    audios = []

    # 目标行
    t_comb_name = '导出路径：'
    target_path = PathX(module=module_name, name='target_path', display='导出路径', content=[DESKTOP])

    # 复选行
    checks: list[Checkbutton] = []
    checks_names = ['完成后打开文件']
    vals: list[BooleanVar] = []
    vals_names = ['is_open']

    def __init__(self, parent, label):
        super().__init__(parent, label)
        self.target_comb.config(values=self.target_path.content)
        self.target_comb.current(0)

    def choose_draft(self):
        super().choose_draft()

    def choose_export(self):
        super().choose_export()

    def analyse_meta(self, draft_path: str):
        self.audios.clear()
        with open(fr'{draft_path}\draft_content.json', 'r', encoding='utf-8') as f:
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
                               'path': fr'{draft_path}\textReading\{item.get("path").split("/")[-1]}',
                               'tune': item.get('tone_type')
                               }
                    self.audios.append(a_audio)
        # 必须及时关闭，否则f时刻被占用，不可替换内部资源
        f.close()
        return has_audio

    def patch_fun(self):
        if isdir(self.target_comb.get()):
            # 写入导出路径
            self.target_path.add(self.target_comb.get())
            self.target_path.save(parser=self.p.configer, update=True)
            for draft in self.draft_todo[names2name(self.draft_todo).index(self.draft_comb.get())]:
                have_audio = self.analyse_meta(draft)
                if have_audio:
                    # 不能使用冒号，否则OSError: [WinError 123] 文件名、目录名或卷标语法不正确
                    suffix = strftime('%m.%d.%H-%M-%S', localtime())
                    filepath = '{}\\{}-收集的配音-{}'.format(self.target_path.now(), basename(draft), suffix)
                    for audio in self.audios:
                        win32_shell_copy(audio.get('path'),
                                         '{}\\{}-{}-{}.wav'.format(filepath,
                                                                   # 序号既能防止重复，也能标示先后顺序
                                                                   audio.get('serial'),
                                                                   audio.get('txt'),
                                                                   audio.get('tune')
                                                                   )
                                         )

                else:
                    messagebox.showwarning(title='缺少内容',
                                           message=f'您选择的草稿{basename(draft)}中没有找到配音！')
            if self.vals[0].get() == 1:
                startfile(self.target_path.now())
        else:
            messagebox.showwarning(title='路径无效', message='请检查路径是否存在！')
