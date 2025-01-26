import json
import sys
from pywinauto import Application
import pyautogui
import time
from PIL import ImageGrab
import easyocr
import pyperclip
import win32gui
import os
# 定义常量
# 搜索框相对于微信窗口左上角的X轴偏移量
SEARCH_X_OFFSET = 100
# 搜索框相对于微信窗口左上角的Y轴偏移量
SEARCH_Y_OFFSET = 40
# 输入文本框相对于微信窗口右下角的X轴偏移量
INPUT_TEXT_LOCATION_X_OFFSET = -100
# 输入文本框相对于微信窗口右下角的Y轴偏移量
INPUT_TEXT_LOCATION_Y_OFFSET = -70
# 截图区域的左上角X坐标
RECTANGLE_LEFT = 0
# 截图区域的左上角Y坐标
RECTANGLE_TOP = 0
# 截图区域相对于搜索框的X轴偏移量
RECTANGLE_X_OFFSET = -40
# 截图区域相对于搜索框的Y轴偏移量
RECTANGLE_Y_OFFSET = 30
# 截图区域的宽度
RECTANGLE_WIDTH = 200
# 截图区域的高度
RECTANGLE_HEIGHT = 90

def open_wx():
    """
    启动微信应用程序并将其窗口置于最前面。
    """
    global app
    wechat_path = 'c:\\program files (x86)\\Tencent\\Wechat\\WeChat.exe'
    if not os.path.exists(wechat_path):
        raise Exception(f"微信应用程序路径不存在: {wechat_path}")
    app = Application(backend='uia').start(wechat_path)        
    app.wait_cpu_usage_lower(threshold=5, timeout=30, usage_interval=1)
    hwnd = win32gui.FindWindow(None, '微信')
    if hwnd:
        if app is not None:
            app.window(handle=hwnd).set_focus()
            rect = app.window(handle=hwnd).rectangle()
            left, top = rect.left, rect.top            
            right, bottom = left + rect.width(), top + rect.height()            
            init_wx(left, top, right, bottom)
        else:
            raise Exception("Application对象未正确初始化")
    else:
        raise Exception("未找到微信窗口")

def init_wx(left, top, right, bottom):
    """
    初始化微信窗口的位置信息。
    """
    global input_text_location, search_location, rectangle_left, rectangle_right
    input_text_location = [right + INPUT_TEXT_LOCATION_X_OFFSET, bottom + INPUT_TEXT_LOCATION_Y_OFFSET]
    search_location = [left + SEARCH_X_OFFSET, top + SEARCH_Y_OFFSET]
    rectangle_left = search_location[0] + RECTANGLE_X_OFFSET
    rectangle_right = search_location[1] + RECTANGLE_Y_OFFSET


def send_msg():
    """
    读取配置文件并发送消息给好友。
    """
    with open('msg.json', 'r', encoding='utf-8') as f:
        msg_dict = json.load(f)
    for name, msg in msg_dict.items():
        send_message_to_user(name, msg)

def capture_and_ocr(x, y, width=RECTANGLE_WIDTH, height=RECTANGLE_HEIGHT):
    """
    截取指定区域的图像并进行OCR识别。
    """
    screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))
    screenshot.save('captured_image.png')
    reader = easyocr.Reader(['ch_sim', 'en'])
    result = reader.readtext('captured_image.png')
    text = ' '.join([res[1] for res in result])
    print("识别结果:", text)
    return text

def send_message_to_user(name, msg):
    """
    向指定好友发送消息。
    """
    print("正在发送消息给好友：", name, msg)
    pyautogui.click(search_location)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('delete')
    pyperclip.copy(name)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1.5)
    # 截取搜索结果区域的图像并进行OCR识别,判断是否存在好友
    text = capture_and_ocr(rectangle_left, rectangle_right)
    if text.find(name) != -1:
        pyautogui.click(rectangle_left + 20, rectangle_right + 40)
        time.sleep(0.5)
        pyperclip.copy(msg)        
        pyautogui.click(input_text_location)
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        print("发送成功")
    else:
        print("未找到好友：", name)

if __name__ == "__main__":
    open_wx()
    send_msg()
