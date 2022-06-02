import psutil


def judge_process(process_name):
    pl = psutil.pids()
    for pid in pl:
        if psutil.Process(pid).name() == process_name:
            print(pid)
            break
    else:
        print("not found")


if judge_process('JianyingPro.exe') == 0:
    print('success')
else:
    pass
