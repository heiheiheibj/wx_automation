import json
import os
import subprocess
import sys
from pywinauto import Application
import pyautogui
import time
from PIL import ImageGrab
import easyocr
import pyperclip
import psutil

# 定义常量
# 搜索框相对于微信窗口左上角的X轴偏移量
SEARCH_BOX_X_OFFSET = 100
# 搜索框相对于微信窗口左上角的Y轴偏移量
SEARCH_BOX_Y_OFFSET = 40
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
RECTANGLE_Y_OFFSET = 25
# 截图区域的宽度
RECTANGLE_WIDTH = 200
# 截图区域的高度
RECTANGLE_HEIGHT = 90
# 是否发送群消息
SEND_GROUP_MSG = True

def is_wechat_running():
    """
    检查微信是否正在运行。
    
    返回:
        bool: 如果微信正在运行返回True，否则返回False。
    """
    for proc in psutil.process_iter(['name']):
        if proc.info['name'].lower().find('wechat') >= 0:
            return True
    return False

def open_wx():
    """
    打开微信应用程序。
    
    如果微信未运行，则启动微信并等待窗口加载。
    获取微信主窗口的坐标，并初始化相关位置信息。
    
    抛出:
        Exception: 如果连接微信窗口失败。
    """
    
    wx_path = r'C:\Program Files (x86)\Tencent\WeChat\WeChat.exe'
    if not is_wechat_running():
        # 启动微信
        subprocess.Popen(wx_path)
        # 等待微信启动
        time.sleep(5)  # 等待一段时间以便窗口加载
    try:
        # 使用pywinauto连接到微信窗口
        app = Application(backend="uia").connect(path=wx_path)
        # 获取微信主窗口
        main_window = app.window(title="微信")
        main_window.set_focus()  # 将微信窗口置顶
        # 获取窗口坐标
        rect = main_window.rectangle()
        left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
        # print(f"微信窗口坐标：左={left}, 上={top}, 右={right}, 下={bottom}") 
        init_wx(left, top, right, bottom)
    except Exception as e:
        raise Exception("连接微信窗口失败，请激活微信窗口后再运行程序。", e)

def init_wx(left, top, right, bottom):
    """
    初始化微信窗口的位置信息。
    
    参数:
        left (int): 微信窗口的左边界坐标。
        top (int): 微信窗口的上边界坐标。
        right (int): 微信窗口的右边界坐标。
        bottom (int): 微信窗口的下边界坐标。
    """
    global input_text_location, search_location, rectangle_left, rectangle_top
    input_text_location = [right + INPUT_TEXT_LOCATION_X_OFFSET, bottom + INPUT_TEXT_LOCATION_Y_OFFSET]
    search_location = [left + SEARCH_BOX_X_OFFSET, top + SEARCH_BOX_Y_OFFSET]
    rectangle_left = search_location[0] + RECTANGLE_X_OFFSET
    rectangle_top = search_location[1] + RECTANGLE_Y_OFFSET
    # print("input_text_location:", input_text_location)
    # print("search_location:", search_location)
    # print("rectangle_left:", rectangle_left)
    # print("rectangle_top:", rectangle_top)

def send_msg():
    """
    读取配置文件并发送消息给好友。
    
    从'msg.json'文件中读取消息配置，遍历每个好友并发送相应的消息。
    """
    with open('msg.json', 'r', encoding='utf-8') as f:
        msg_dict = json.load(f)
    for name, msg in msg_dict.items():
        send_message_to_user(name, msg)

def capture_and_ocr(x, y, width=RECTANGLE_WIDTH, height=RECTANGLE_HEIGHT):
    """
    截取指定区域的图像并进行OCR识别。
    
    参数:
        x (int): 截图区域的左上角X坐标。
        y (int): 截图区域的左上角Y坐标。
        width (int): 截图区域的宽度，默认为RECTANGLE_WIDTH。
        height (int): 截图区域的高度，默认为RECTANGLE_HEIGHT。
    
    返回:
        str: OCR识别结果。
    """
    screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))
    screenshot.save('captured_image.png')
    reader = easyocr.Reader(['ch_sim', 'en'])
    result = reader.readtext('captured_image.png')
    text = ' '.join([res[1] for res in result])
    print("识别结果:", text)
    return text

def is_user_or_group(text,name):
    """
    检查搜索的结果是用户还是群组。

    """
    if text.find('包含:') != -1 :
        return 'group'

    if text.find(text.lower()) != -1 and '天记' != text[1:3] and text.find('网络查找微信号') == -1:
        return 'user'
    
    return 'unknown'
    

def send_message_to_user(name, msg):
    """
    向指定好友发送消息。
    
    参数:
        name (str): 好友的名称。
        msg (str): 要发送的消息内容。
    
    打印发送消息的状态，并在发送成功后输出提示信息。
    """
    print("正在发送消息给好友：", name, msg)
    pyautogui.click(search_location)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('delete')
    pyperclip.copy(name)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1.5)
    # 截取搜索结果区域的图像并进行OCR识别,判断是否存在好友
    text = capture_and_ocr(rectangle_left, rectangle_top).lower()
    #判断text中第2个和第3个字符是否是'天记'.因为ocr的原因，会把聊天记录识别为职天记灵
    
    #if text.find(name.lower()) != -1 and '天记' != text[1:3] and text.find('网络查找微信号') == -1:
    user_or_group = is_user_or_group(text, name)
    if (user_or_group == 'user') or (user_or_group == 'group' and SEND_GROUP_MSG):
        if is_user_or_group(text,name) == 'group':
            msg = f'@{name} ' + msg
        pyautogui.click(rectangle_left + 20, rectangle_top + 40)
        time.sleep(0.5)
        pyperclip.copy(msg)
        # print('input_text_location', input_text_location)
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
