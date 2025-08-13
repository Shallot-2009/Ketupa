# pip install pyautogui

from time import sleep

import pyautogui
import time

def bbox_to_coords(bbox, screen_width, screen_height):
    """将 bbox 坐标转换为屏幕坐标."""
    xmin, ymin, xmax, ymax = bbox
    x_center = int((xmin + xmax) / 2 * screen_width)
    y_center = int((ymin + ymax) / 2 * screen_height)
    return x_center, y_center


def click_bbox(bbox):
    """点击指定的 bbox."""
    screen_width, screen_height = pyautogui.size()
    x, y = bbox_to_coords(bbox, screen_width, screen_height)

    # 移动鼠标到指定位置
    pyautogui.moveTo(x, y, duration=0.1)  # duration 是移动时间，单位为秒

    # 点击鼠标
    pyautogui.doubleClick()

    print(f"点击了坐标: x={x}, y={y}")

if __name__ == '__main__':

    sleep(5)

    # 示例 bbox (来自您提供的数据)
    bbox = [0.0028251188341528177, 0.010840130038559437, 0.04837050661444664, 0.11918634921312332] # chrome

    # 点击 bbox
    click_bbox(bbox)




def click_bbox(bbox):
    """点击指定的 bbox."""
    screen_width, screen_height = pyautogui.size()
    x, y = bbox_to_coords(bbox, screen_width, screen_height)

    # 移动鼠标到指定位置
    pyautogui.moveTo(x, y, duration=0.1)  # duration 是移动时间，单位为秒

    # 点击鼠标
    pyautogui.doubleClick()

    print(f"点击了坐标: x={x}, y={y}")

if __name__ == '__main__':

    sleep(5)

    # 示例 bbox (来自您提供的数据)
    bbox = [0.18070517480373383, 0.2619055509567261, 0.2084973156452179, 0.31362271308898926] # chrome

    # 点击 bbox
    click_bbox(bbox)


    def click_bbox(bbox):
        """点击指定的 bbox."""
        screen_width, screen_height = pyautogui.size()
        x, y = bbox_to_coords(bbox, screen_width, screen_height)

        # 移动鼠标到指定位置
        pyautogui.moveTo(x, y, duration=0.1)  # duration 是移动时间，单位为秒

        # 点击鼠标
        pyautogui.doubleClick()

        print(f"点击了坐标: x={x}, y={y}")


    if __name__ == '__main__':
        sleep(5)

        # 示例 bbox (来自您提供的数据)
        bbox = [0.1779809296131134, 0.36977070569992065, 0.29686635732650757, 0.3965611457824707]  # chrome

        # 点击 bbox
        click_bbox(bbox)