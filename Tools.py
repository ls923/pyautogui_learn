import os
import datetime as dt
import time
import Auto
import Enums

LOG_METHOD = 1
LOG_FILE = 'Auto.log'
today = time.strftime("%Y%m%d", time.localtime())


def get_dir_list(dir_path):
    if dir_path == "":
        return []
    dir_path = dir_path.replace("/", "\\")
    if dir_path[-1] != "\\":
        dir_path = dir_path + "\\"
    alist = os.listdir(dir_path)
    dirlist = [x for x in alist if os.path.isdir(dir_path + x)]
    return dirlist


#  @ 功能：日志记录 调试时可以选择输出控制台
#  @ 参数：[I] :*BUF 输入的内容
def log(*buf):
    #  输出到文件
    if LOG_METHOD == 1:
        with open(LOG_FILE, 'a') as logfile:
            print(dt.datetime.now().strftime('%F %T:%f'), file=logfile, end=' ')
            for i in buf:
                print(i, file=logfile, end=' ')
            print('', file=logfile)

        print(dt.datetime.now().strftime('%F %T:%f'), end=' ')
        for i in buf:
            print(i, end=' ')
        print(end='\n')

    # 输出到控制台
    elif LOG_METHOD == 2:
        print(dt.datetime.now().strftime('%F %T:%f'), end=' ')
        for i in buf:
            print(i, end=' ')
        print(end='\n')
