import asyncio
import os
import re
import sys
import threading
import time

import cv2
import glob2
import keyboard
import pyautogui
import win32gui
import xlrd
from PIL import ImageGrab

import Enums
import Tools as tool

# params define
Running = Enums.RunType.Wait  # 是否运行标志
Key_Value_pair = 0  # 键值对数
NowRowKey = []
NowRowValue = []
ColumnValue = ['', 'B', 'C', 'D', 'E', 'F', 'G']
CurrentRow = 1
TotalTaskList = ['']
LoopCount = 1  # 循环次数
StartKey = 'ctrl+alt+s'  # 开始热键
StopKey = 'ctrl+alt+q'  # 结束热键
ListCfg = ['loopcounter', 'starthotkey', 'stophotkey']  # 下拉栏是独立的
XlsSource = None
WorkPath = ''
DIR = os.path.dirname(__file__)  # 运行路径
StatusText = ''
MSGWindowName = 'AutoWorkMessage'
PicRoot = '.\\Pic\\'


def log(*buf):
    tool.log(buf)


def get_xy(template_img_path):
    # Pillow 截图
    hwnd = get_window("依露希尔：星晓")
    x_start, y_start, x_end, y_end = win32gui.GetWindowRect(hwnd)
    # 坐标信息
    box = (x_start, y_start, x_end, y_end)
    shot = ImageGrab.grab(box)
    shot.save("./Pic/screenshot.png")

    # copy一份图像资源用于debug操作
    shot_origin = cv2.imread("./Pic/screenshot.png")
    shot_copy = shot_origin.copy()

    img = cv2.imread("./Pic/screenshot.png")

    # 加载模板图像，转换灰度图，检测边缘
    # 使用边缘而不是原始图像进行模板匹配可以大大提高模板匹配的精度。
    img_temp = cv2.imread(template_img_path)
    # img_temp = cv2.cvtColor(img_temp, cv2.COLOR_BGR2GRAY)
    # img_temp = cv2.Canny(img_temp, 50, 200)

    # 模版图片的高宽
    height, width, channel = img_temp.shape
    # 进行模版匹配 当前method = 5
    result = cv2.matchTemplate(img, img_temp, cv2.TM_CCOEFF_NORMED)
    # 可能匹配多个符合条件的
    has_match = False
    has_best = False
    # 筛选大于一定匹配值的点
    val, result_best = cv2.threshold(result, 0.95, 1.0, cv2.THRESH_BINARY)
    best_list = cv2.findNonZero(result_best)
    has_best = best_list is not None

    # for x in range(len(result)):
    #     if has_match:
    #         break
    #     else:
    #         for y in range(len(result[x])):
    #             # print(f'{result[x][y]}')
    #             if result[x][y] > 0.99:
    #                 print(f'====== find => {result[x][y]} =======')
    #                 # has_match = True
    #                 # break
    #             elif result[x][y] > 0.5:
    #                 print(f'{result[x][y]}')
    #                 print(f'====== print => {(x, y)} =======')
    #                 cv2.rectangle(shot_copy, (x, y), (x + width, y + height), (0, 0, 255), thickness=2)
    #                 # has_match = True

    if not has_best:
        # 如果没有最佳匹配 就寻找其他匹配点
        val2, result_mutil = cv2.threshold(result, 0.75, 1.0, cv2.THRESH_BINARY)
        match_list = cv2.findNonZero(result_mutil)
        has_match = match_list is not None

    if not has_match and not has_best:
        print(f'未找到匹配图像')
        return None
    elif has_best or has_match:
        # 计算最佳匹配位置
        # method = 0|1 得到最小位置 index = 2
        # method = 其他 得到最大位置出 index = 3
        upper_left = cv2.minMaxLoc(result)[3]
        lower_right = (upper_left[0] + width, upper_left[1] + height)
        star = (upper_left[0], upper_left[1])
        end = (lower_right[0], lower_right[1])

        center = cal_center(upper_left, lower_right)
        print(f'star{star}')
        print(f'end{end}')

        draw_rectangle(shot_copy, star, end)
        # cv2.imshow('shot_copy', cv2.resize(shot_copy, None, fx=0.6, fy=0.6))
        # cv2.waitKey()
        # cv2.destroyAllWindows()

        print(f'找到了最佳匹配{center}')
        return center
    elif has_match and not has_best:
        for r in match_list:
            # result_f 是个三维数组 0=> 1=>固定长度1
            upper_left = r[0]
            lower_right = (upper_left[0] + width, upper_left[1] + height)
            star = (upper_left[0], upper_left[1])
            end = (lower_right[0], lower_right[1])

            center = cal_center(upper_left, lower_right)
            print(f'star{star}')
            print(f'end{end}')

            draw_rectangle(shot_copy, star, end)
        # cv2.imshow('shot_copy', cv2.resize(shot_copy, None, fx=0.6, fy=0.6))
        # cv2.waitKey()
        # cv2.destroyAllWindows()

        # print(f'找到了最佳匹配{center}')
        return None


# 计算中心点
def cal_center(upper_left, lower_right):
    center = (int((upper_left[0] + lower_right[0]) / 2), int((upper_left[1] + lower_right[1]) / 2))
    return center


# 画矩形
def draw_rectangle(image, star, end):
    cv2.rectangle(image, star, end, (0, 0, 255), thickness=2)


# pyautogui 自动点击
def auto_click(center):
    hwnd = get_window("依露希尔：星晓")
    x_start, y_start, x_end, y_end = win32gui.GetWindowRect(hwnd)
    # 坐标信息
    pyautogui.click(x_start + center[0], y_start + center[1], button='left')
    time.sleep(1)


# 主循环
def routine(img_model_path, name, hwnd):
    avg = get_xy(img_model_path, hwnd)
    print(f'正在点击{name}')
    if avg is not None:
        auto_click(avg)


def get_window(title):
    hwnd = win32gui.FindWindow(0, title)  # 寻找窗口
    if not hwnd:
        print("找不到该窗口")
        return None
    else:
        win32gui.SetForegroundWindow(hwnd)  # 前置窗口
        return hwnd


def analysis(img_path, location):
    print(f'正在点击{img_path}')

    auto_click(location)


# 找图并执行动作
def find_pic_and_click(pic_name, timeout, timeout_action, interval):
    img = f'{pic_name}.png'
    img_path = f'{PicRoot}{img}'
    if img_path != '' and os.path.exists(img_path) is True and Running == Enums.RunType.Running:
        log(img_path, '图片有效')

        location = get_xy(img_path)
        view_log = True

        if location is not None:
            log(img_path, 'location is available, Quick run')
            analysis(img_path, location)
        else:
            star_time = time.time()
            while timeout >= 0 and location is None and Running == Enums.RunType.Running:
                if view_log:
                    log(img_path, 'is not appear,waiting..(timeout > 0)', timeout)
                    view_log = False

                location = get_xy(img_path)
                # await asyncio.sleep(interval)
                # loca = pyautogui.locateCenterOnScreen(img_path, confidence=0.9)
                time.sleep(interval)

                if time.time() - star_time > timeout:
                    log(img_path, 'waiting timeout !!!!')
                    log('超时方法： ' + timeout_action)
                    if timeout_action == '弹窗':  # pyautogui.alert和通知有冲突 通知后无法弹窗
                        # pyautogui.alert(text=ImgPath + '查找超时', title=MSGWindowName)
                        # tkinter.messagebox.showinfo(title='PyRPA: ', message=str(ImgPath + '查找超时'), icon='error')
                        return timeout_action
                    else:
                        return timeout_action
            while timeout == -1 and location is None and Running == Enums.RunType.Running:  # 一直找图 热键停止
                if view_log:
                    log(img_path, 'timeout = -1, is not appear,waiting..(timeout = -1)')
                    view_log = False

                # location = pyautogui.locateCenterOnScreen(img_path, confidence=0.9)
                location = get_xy(img_path)
                time.sleep(interval)
                # await asyncio.sleep(interval)

            log(img_path, 'appear,waiting succecs,run')
            analysis(img_path, location)
    elif pic_name == '':
        log(WorkPath, ' Excel中的图片名为空\n【以非找图模式运行】')
        analysis('None', None)
    else:
        log(img_path, '！！图片无效，无法继续运行')


def data_check(sheet):
    log('excel 数据校验')

    def show_error():
        log('！第 ' + str(now_row + 1) + ' 行 ' + ColumnValue[column] + ' 列 数据有问题，程序无法继续运行')
        exit(-1)

    def action_check(n_row):
        action = str(sheet.row(n_row)[6].value)
        action = (action.replace(',', '，')
                  .replace('“', '"')
                  .replace('”', '"'))
        if action.count('=') == action.count('，') + 1:
            return True
        else:
            log('！第 ' + str(n_row + 1), '行 动作队列异常')
            return False

    for now_row in range(1, sheet.nrows):
        for column in range(1, 7):  # 判断各行的各列
            # 单元格数据值类型
            stype = sheet.cell(now_row, column).ctype
            # 是否为启用列
            is_enable_col = Enums.ColType(column) == Enums.ColType.Enable
            # 是否为图片列
            is_pic_col = Enums.ColType(column) == Enums.ColType.PicName
            # 是否为超时时间列
            is_time_out_col = Enums.ColType(column) == Enums.ColType.TimeOut
            # 是否为超时行为列
            is_time_out_action_col = Enums.ColType(column) == Enums.ColType.TimeOutAction
            # 是否为间隔列
            is_interval_col = Enums.ColType(column) == Enums.ColType.Interval
            # 是否为动作列
            is_action_col = Enums.ColType(column) == Enums.ColType.Action

            # 启用列数据是否错误标志
            enable_type_err = is_enable_col and Enums.STYPE(stype) != Enums.STYPE.Num
            # 图片列数据是否错误标志
            pic_type_err = is_pic_col and Enums.STYPE(stype) != Enums.STYPE.Str
            # 超时列数据是否错误标志
            time_out_type_err = is_time_out_col and Enums.STYPE(stype) != Enums.STYPE.Num
            # 间隔列数据是否错误标志
            interval_type_err = is_interval_col and Enums.STYPE(stype) != Enums.STYPE.Num
            # 动作列数据是否错误标志
            action_type_err = is_action_col and Enums.STYPE(stype) != Enums.STYPE.Str

            # 是否启用
            enable = Enums.EnableType(sheet.row(now_row)[1].value) == Enums.EnableType.Enable
            # 图片名
            pic_name = sheet.row(now_row)[2].value
            # 超时时间
            time_out = sheet.row(now_row)[3].value
            # 超时行为
            time_out_action = sheet.row(now_row)[4].value

            if enable_type_err:
                show_error()
                return False
            elif enable and pic_name != '':  # 检查启用并且为找图模式的
                if interval_type_err or action_type_err:
                    show_error()
                    return False
                if pic_type_err:  # 图片列 值类型不为str
                    if Enums.STYPE(stype) != Enums.STYPE.Empty:
                        show_error()
                        return False
                if time_out_type_err:  # 超时列 值类型不为数字
                    show_error()
                    return False
                elif is_time_out_action_col:
                    # 超时行为列 & 超时时间!=-1
                    if int(time_out) != -1:
                        temp_str = str(time_out_action)
                        if not (temp_str == '弹窗'
                                or temp_str == '跳过'
                                or temp_str == '退出'
                                or (temp_str.find("跳转") != -1)):
                            show_error()
                            return False
            else:
                break
        if action_check(now_row) is False:
            return False
    return True


# 校验表数据
def work_space(sheet):
    global NowRowKey, NowRowValue, StatusText, Key_Value_pair, JumpLine, CurrentRow
    StatusText = '工作'
    if sheet is None:
        return

    if data_check(sheet) is True:
        log('数据校验通过')
    else:
        return

    # 开始遍历表数据
    CurrentRow = 1

    while CurrentRow < sheet.nrows and Running == Enums.RunType.Running:
        enable = sheet.row(CurrentRow)[1].value == Enums.EnableType.Enable.value
        if enable:
            log('------------ start work ------------')
            log(f'excel row {CurrentRow + 1}')
            action_str = sheet.row(CurrentRow)[6].value
            log(f'excel str {action_str}')
            replace_str = action_str.replace(',', '，').replace('，', '=')
            split = re.split('=', replace_str)

            _i = 0
            _count: int = 0
            while _count < len(split):
                NowRowKey.append(split[_count])
                NowRowValue.append(split[_count + 1])
                _count += 2
                _i += 1
            Key_Value_pair = _i
            ret = str(find_pic_and_click(pic_name=sheet.row(CurrentRow)[2].value,
                                         timeout=sheet.row(CurrentRow)[3].value,
                                         timeout_action=sheet.row(CurrentRow)[4].value,
                                         interval=sheet.row(CurrentRow)[5].value))
            log("FindPicAndClick ret=", ret)
            NowRowKey.clear()
            NowRowValue.clear()
            if ret == '退出':
                log('查找超时,退出整个查找')
                return ret

            if ret.find("跳转") != -1:
                temp_list = re.split('=', ret)
                if len(temp_list) > 0:
                    CurrentRow = int(temp_list[1]) - 2  # 针对程序
                    log("由超时行为触发的跳转到第 ", int(temp_list[1]), "行")  # 针对用户
                    if CurrentRow > sheet.nrows or CurrentRow < 0:
                        log("！请检查跳转参数")
                        return -1
        else:
            log("excel row", CurrentRow + 1, '未启用操作')

            if JumpLine != -1:
                log("由动作触发的跳转到第 ", JumpLine, "行")
                CurrentRow = int(JumpLine) - 2
                JumpLine = -1
                if CurrentRow > sheet.nrows or CurrentRow < 0:
                    log("！请检查跳转参数")
                    return -1
        CurrentRow += 1
    log('works end')


# 初始化
def initial():
    log('Run path:', DIR)
    log('Execute File:', sys.argv[0])

    if os.path.exists('Source') is not True:
        log('! Source文件夹不存在，程序无法继续运行')
    else:
        total_task_list = tool.get_dir_list('Source')
        log('当前Source文件夹内容(可选任务列表):', total_task_list)


# 加载表数据
def thread_init_file():
    global TotalTaskList, LoopCount, StartKey, StopKey, XlsSource

    def update_current_xls():  # 表重定向
        global XlsSource
        global WorkPath
        WorkPath = '.\\Source\\'

        file_list = os.listdir(WorkPath)
        for f in range(len(file_list)):
            file_list[f] = os.path.splitext(file_list[f])[1]

        if '.xls' not in file_list:
            log('！路径' + WorkPath + ' 下可能没有任务表，程序无法继续运行，\n请添加表格或者删除任务文件夹')

        now_dir_xls_path = glob2.glob(WorkPath + '\\*.xls')[0]
        if os.path.exists(now_dir_xls_path) is not True:
            log('！' + now_dir_xls_path + ' 不存在，程序无法继续运行')
        else:
            log('任务路径更新:', now_dir_xls_path)
            XlsSource = xlrd.open_workbook(filename=now_dir_xls_path).sheet_by_index(0)

    update_current_xls()


# 开始热键绑定的事件
def begin_working():
    log('热键按下 ：begin_working')
    global Running
    # WindowCtrl(ClassWindow, WindowName, 0)
    # mutex.acquire()
    Running = Enums.RunType.Running
    # start_up()
    # mutex.release()


# 结束热键绑定的事件
def finished_working():
    global Running
    # mylog('热键按下 ：finished_working')
    # WindowCtrl(ClassWindow, WindowName, 1)
    # mutex.acquire()
    Running = Enums.RunType.Stop
    # mutex.release()
    # WindowCtrl(None, MSGWindowName, -1)


#

# ----------------- MAIN -----------------
if __name__ == '__main__':
    initial()
    RunCount = 0
    threading.Thread(target=thread_init_file).start()
    log(' ————————————————————————————————————————————')
    log('|欢迎使用自动化软件！  <程序版本V0.0.1>')
    log('|作者: ** 838164707@qq.com')
    log('|鸣谢: **')
    log(' ————————————————————————————————————————————\n')


    def app_start():
        # async def app_start():
        global Running
        global RunCount
        global XlsSource
        if StartKey != '' and StopKey != '':
            keyboard.add_hotkey(StartKey, begin_working)
            keyboard.add_hotkey(StopKey, finished_working)
        log('等待热键按下,或点击开始')

        while Running == Enums.RunType.Wait:
            # await asyncio.sleep(1)
            time.sleep(1)
            log('准备')

        Running = Enums.RunType.Running
        RunCount = int(LoopCount)
        if RunCount == -1:
            log("循环模式")
            while Running == Enums.RunType.Running:
                if work_space(XlsSource) == '退出':
                    break
        else:
            count = 0
            total_count = RunCount
            while RunCount > 0 and Running == Enums.RunType.Running:
                count += 1
                log(f'\n [运行 {count}/{total_count} 次]')
                if work_space(XlsSource) == '退出':
                    break
                RunCount -= 1
        log("Excel 遍历结束")


    # asyncio.run(app_start())
    app_start()


def start_up():
    pass
