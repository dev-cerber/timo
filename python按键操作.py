import win32api
import time
#鼠标移动
import win32con

from config import VK_CODE


def mouse_move(x,y):
    win32api.SetCursorPos([x,y])

#鼠标点击，默认左键
def mouse_click(click_type="right"):
    if click_type=="left":
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    else:
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP | win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    time.sleep(0.01)

#鼠标双击击，默认左键
def mouse_double_click(click_type="left"):
    if click_type=="left":
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.01)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

    else:
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP | win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        time.sleep(0.01)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP | win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    time.sleep(0.01)


def key_input( input_words=''):
    for word in input_words:
        win32api.keybd_event(VK_CODE[word], 0, 0, 0)
        win32api.keybd_event(VK_CODE[word], 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)


# 根据config中的键进行选择(组合选择)
def key_even(input_key):
    win32api.keybd_event(VK_CODE[input_key], 0, 0, 0)
    time.sleep(0.01)
    win32api.keybd_event(VK_CODE[input_key], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.3)


if __name__ == '__main__':
    while True:
        mouse_click()
        key_even('e')