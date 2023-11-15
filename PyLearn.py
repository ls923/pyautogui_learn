import time

import cv2
import pyautogui
from PIL import ImageGrab


def get_xy(img_model_path):
    # pyautogui._screenshot_win32

    shot = ImageGrab.grab()
    shot.save("./Pic/screenshot.png")

    shot_origin = cv2.imread("./Pic/screenshot.png")
    shot_copy = shot_origin.copy()

    img = cv2.imread("./Pic/screenshot.png")
    img_temp = cv2.imread(img_model_path)

    height, width, channel = img_temp.shape
    result = cv2.matchTemplate(img, img_temp, cv2.TM_CCOEFF_NORMED)

    has_match = False
    for x in range(len(result)):
        if has_match:
            break
        else:
            for y in range(len(result[x])):
                # print(f'{result[x][y]}')
                if result[x][y] > 0.99:
                    print(f'====== find => {result[x][y]} =======')
                    has_match = True
                    break
                elif result[x][y] > 0.75:
                    print(f'{result[x][y]}')
                    print(f'====== print => {(x, y)} =======')
                    cv2.rectangle(shot_copy, (x, y), (x + width, y + height), (0, 0, 255), thickness=2)
                    has_match = True
    if not has_match:
        print(f'未找到匹配图像')
        return None
    else:
        upper_left = cv2.minMaxLoc(result)[3]
        lower_right = (upper_left[0] + width, upper_left[1] + height)
        star = (upper_left[0], upper_left[1])
        end = (lower_right[0], lower_right[1])

        print(f'star{star}')
        print(f'end{end}')
        cv2.rectangle(shot_copy, star, end, (255, 255, 255), thickness=2)
        cv2.imshow('shot_copy', cv2.resize(shot_copy, None, fx=0.6, fy=0.6))
        # cv2.waitKey()
        time.sleep(5)
        cv2.destroyAllWindows()
        center = (int((upper_left[0] + lower_right[0]) / 2), int((upper_left[1] + lower_right[1]) / 2))

        return center


def auto_click(center):
    pyautogui.click(center[0], center[1], button='left')
    time.sleep(1)


def routine(img_model_path, name):
    avg = get_xy(img_model_path)
    print(f'正在点击{name}')
    if avg is not None:
        auto_click(avg)


routine("./Pic/kaizhan.png", "任务")
