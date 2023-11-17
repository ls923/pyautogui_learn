import time
import cv2
import pyautogui
import win32gui
from PIL import ImageGrab


def get_xy(template_img_path, hwnd):
    # Pillow 截图
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
        cv2.imshow('shot_copy', cv2.resize(shot_copy, None, fx=0.6, fy=0.6))
        cv2.waitKey()
        cv2.destroyAllWindows()

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
        cv2.imshow('shot_copy', cv2.resize(shot_copy, None, fx=0.6, fy=0.6))
        cv2.waitKey()
        cv2.destroyAllWindows()

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
    pyautogui.click(center[0], center[1], button='left')
    time.sleep(1)


# 主循环
def routine(img_model_path, name, hwnd):
    avg = get_xy(img_model_path, hwnd)
    print(f'正在点击{name}')
    if avg is not None:
        auto_click(avg)


def get_window(title):
    # hwnd = win32gui.FindWindow(lpClassName=None, lpWindowName=None)  # 查找窗口，不找子窗口，返回值为0表示未找到窗口
    # hwnd = win32gui.FindWindowEx(hwndParent=0, hwndChildAfter=0, lpszClass=None, lpszWindow=None)  # 查找子窗口，返回值为0表示未找到子窗口
    # win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
    # GetDesktopWindow 获得代表整个屏幕的一个窗口（桌面窗口）句柄
    # hd = win32gui.GetDesktopWindow()
    #
    # # 获取所有子窗口
    # hwndChildList = []
    #
    # win32gui.EnumChildWindows(hd, lambda hwnd, param: param.append(hwnd), hwndChildList)
    #
    # for hwnd in hwndChildList:
    #     print("句柄：", hwnd, "标题：", win32gui.GetWindowText(hwnd))

    hwnd = win32gui.FindWindow(0, title)  # 寻找窗口
    if not hwnd:
        print("找不到该窗口")
        return None
    else:
        win32gui.SetForegroundWindow(hwnd)  # 前置窗口
        time.sleep(1)
        return hwnd


hwnd = get_window("依露希尔：星晓")

routine("./Pic/img.png", "~", hwnd)
