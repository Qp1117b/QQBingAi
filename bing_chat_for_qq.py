# -*- coding:utf-8 -*-

import argparse
import asyncio
import atexit
import glob
import json
import os
import re
import time
import traceback
from io import BytesIO

import BingImageCreator
import pyperclip
import uiautomation as auto
import win32clipboard
import win32con
import win32gui
from BingImageCreator import ImageGen
from EdgeGPT import Chatbot, ConversationStyle
from PIL import Image
from pygtrans import Translate
from requests import RequestException
from uiautomation import WindowControl

parser = argparse.ArgumentParser(description="A script to process files")
# 发送窗口
parser.add_argument("-sender", "--sender", help="QQ title", default='12')
# 凭证
parser.add_argument("-c", "--cookies", help="cookies_.json path", default='./cookies.json')
# 需要检查的机器人在群里的名称
parser.add_argument("-bot", "--bot", help="check @bot", default="bot")

parser.add_argument("--quiet", type=str)

args = parser.parse_args()

# 各种临时文件
path = os.path.split(os.path.realpath(__file__))[0] + "/temp/" + args.sender + "/"
if not os.path.exists(path):
    os.makedirs(path)
source = path + args.sender + "_source.png"
idea_ = path + args.sender + "_idea.txt"
say = path + args.sender + "_say.txt"
save_img_path = path + args.sender + "_img_path"
record_ = path + args.sender + "_record.txt"
orgin_idea = idea_ + ".txt"

print("发送窗口", args.sender)
print("图片源", source)
print("结果源", idea_)
print("原始源", orgin_idea)
print("@源", say)
print("cookies源", args.cookies)
print("检测@对象", args.bot)
print("生成图片的文件夹", save_img_path)


def sit_unlock():
    if os.path.exists(sit_path):
        os.remove(sit_path)


def task(fl_):
    global init
    global bot

    if init == 0:
        send_info_to_qq("Bing Chat已上线，@bot提问，结尾可用：.c灵活，.b平衡，.p精确，.d绘画，.r重置，默认.c")
        init = 1

    m_r = get_refresh_messages()
    m_u = get_unread_messages(m_r)
    m_f = format_messages(m_u)
    print(m_f)

    for message in m_f:
        if message['@'] == args.bot:
            user = message['sender']
            info = message['@i']
            print(user, info)

            idea = edge_loop.run_until_complete(edgegpt(info))

            if idea is None:
                continue

            print(idea)

            if user == args.bot:
                with open(idea_, 'w', encoding='utf-8') as file_:
                    file_.write(idea)
            else:
                with open(idea_, 'w', encoding='utf-8') as file_:
                    file_.write("@" + user + "\n\n" + idea)
            file_.close()

            send(is_sort=False)

    with open(say, 'a') as file_:
        file_.write(str(time.strftime("%H:%M:%S",
                                      time.localtime())) + " " + str(m_u) + '\n')
        file_.close()

    if fl_ is not None:
        fl_.release()
        sit_unlock()


def cn_en(info):
    client = Translate()
    try:
        text = client.translate(info, "en", "zh-CN")
    except RequestException:
        traceback.print_exc()
        send_info_to_qq("描述翻译失败，网络异常")
        return None

    return text.translatedText


def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()


def send_img_windows_and_delete():
    files = glob.glob(os.path.join(save_img_path, "*"))
    for file_ in files:
        image = Image.open(file_)

        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        send_to_clipboard(win32clipboard.CF_DIB, data)
        win32gui.SendMessage(handle_sender, 770, 0, 0)
        os.remove(file_)


def generator_image(prompt):
    prompt_en = cn_en(prompt)

    if prompt_en is None:
        return None

    image_generator = ImageGen(cookies_U, args.quiet)

    send_info_to_qq("开始作画: " + prompt_en)

    try:
        image_generator.save_images(image_generator.get_images(prompt_en), output_dir=save_img_path)
    except Exception as e:
        traceback.print_exc()
        if e.args[0] == BingImageCreator.error_blocked_prompt:
            send_info_to_qq(e.args[0])
            return None

    return "success"


# BingChat模型
async def edgegpt(info):
    global bot
    info_s = str(info)
    if info_s.endswith(".d") and len(info[:-2]) > 0:

        if generator_image(info[:-2]) is None:
            send_info_to_qq("作画失败")
            return None

        send_info_to_qq("作画完成: " + info_s[:-2])
        send_img_windows_and_delete()
        return str(info[:-2])
    elif info_s.endswith(".r"):
        await bot.reset()
        send_info_to_qq("重置对话完成")
        return None
    else:
        if info_s[len(info_s) - 2] != '.':
            info_s += ".c"

        if info_s.endswith(".c"):
            print("灵活")
            which_model = "灵活"
        elif info_s.endswith(".b"):
            print("平衡")
            which_model = "平衡"
        elif info_s.endswith(".p"):
            print("精确")
            which_model = "精确"
        else:
            print("灵活")
            which_model = "灵活"

        send_info_to_qq("正在回答(" + which_model + ")：" + info_s[:-2])

        if which_model == "灵活":
            result = await bot.ask(prompt=info_s[:-2], conversation_style=ConversationStyle.creative,
                                   wss_link="wss://sydney.bing.com/sydney/ChatHub")
        elif which_model == "平衡":
            result = await bot.ask(prompt=info_s[:-2], conversation_style=ConversationStyle.balanced,
                                   wss_link="wss://sydney.bing.com/sydney/ChatHub")
        else:
            result = await bot.ask(prompt=info_s[:-2], conversation_style=ConversationStyle.precise,
                                   wss_link="wss://sydney.bing.com/sydney/ChatHub")

        orgin_idea_.write(str(result) + "\n\n")
        orgin_idea_.flush()

        try:
            idea = result['item']['messages'][1]['text']
        except Exception:
            traceback.print_exc()
            if "messages" in result['item']:
                try:
                    idea = result['item']['messages'][1]['hiddenText']
                except Exception:
                    idea = "请使用.r重置"
            else:
                idea = result

        try:
            sourceAttributions = result['item']['messages'][1]['sourceAttributions']
        except Exception:
            traceback.print_exc()
            sourceAttributions = None

        if sourceAttributions is not None and len(sourceAttributions) != 0:
            idea += "\n\n更多：\n"
            i = 1
            for item in sourceAttributions:
                idea += str(i) + ". " + item['providerDisplayName'] + "【" + item['seeMoreUrl'] + "】\n"
                i += 1
            idea = idea[:-1]
        try:
            suggestedResponses = result['item']['messages'][1]['suggestedResponses']
        except Exception:
            traceback.print_exc()
            suggestedResponses = None

        if suggestedResponses is not None and len(sourceAttributions) != 0:
            idea += "\n\n建议：\n"
            i = 1
            for item in suggestedResponses:
                idea += str(i) + ". " + item['text'] + "\n"
                i += 1
            idea = idea[:-1]
        return idea


# 自动化发送
def send(is_sort=True):
    with open(idea_, encoding='utf-8', errors='ignore', mode='r+') as file_:
        text = file_.read()
        file_.truncate(0)
        file_.close()

    if text != '':

        send_w = auto.WindowControl(searchDepth=1, Name=args.sender)
        active_window(send_w)

        text = re.sub(r"\[\^(\d+)\^\]", r"[\1]", text)

        if not is_sort:
            if text[0] == '@':
                at_ = text.split('\n')[0]
                user = at_[1:]
                text = text[len(user) + 1:]
                At(user)
            text = text + "\n\n内容来自Bing Chat，@" + args.bot + "提问"

        print("send", text)
        pyperclip.copy(text)
        win32gui.SendMessage(handle_sender, 770, 0, 0)
        win32gui.SendMessage(handle_sender, win32con.WM_KEYDOWN,
                             win32con.VK_RETURN, 0)


def send_info_to_qq(msg):
    # return
    if msg is None or len(msg) == 0:
        return
    with open(idea_, 'w', encoding='utf-8') as file_:
        file_.write(msg)
        file_.close()
        send()


@atexit.register
def quit_func():
    if os.path.exists(sit_path):
        os.remove(sit_path)

    send_info_to_qq("bot已下线")


class FileLock:
    def __init__(self, file_path):
        self.file_path = file_path
        self.lock_file_path = file_path + '.lock'
        self.lock_file = None

    def acquire(self):
        while True:
            try:
                self.lock_file = os.open(self.lock_file_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                break
            except FileExistsError:
                time.sleep(0.1)

    def release(self):
        if self.lock_file is not None:
            os.close(self.lock_file)
            os.remove(self.lock_file_path)
            self.lock_file = None

    def __enter__(self):
        self.acquire()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.release()


def to_special_group(group):
    chat_list = \
        message_manager_main_w.GetChildren()[2].GetChildren()[0].GetChildren()[1].GetChildren()[0].GetChildren()[
            1].GetChildren()[0]
    special_chat = chat_list.ListItemControl(searchDepth=1, Name=group)
    # special_chat.Click()
    # time.sleep(0.1)
    special_chat.DoubleClick()

    message_infos_w = message_manager_main_w.ListControl(searchDepth=10, Name="IEMsgView")
    message_infos_w_p = message_infos_w.GetParentControl()
    reset_but = message_infos_w_p.GetChildren()[2].GetChildren()[1].GetChildren()[4]
    reset_but.Click()

    while len(get_no_refresh_messages()) >= 60:
        next_messages_page()


def active_window(w: WindowControl):
    w.SetActive()
    w.SetTopmost(True)


def get_no_refresh_messages():
    active_window(message_manager_main_w)

    message_infos_w = message_manager_main_w.ListControl(searchDepth=10, Name="IEMsgView")

    messages = []
    for message_info in message_infos_w.GetChildren():
        messages.append(message_info.Name)

    return messages


def get_refresh_messages():
    active_window(message_manager_main_w)

    message_infos_w = message_manager_main_w.ListControl(searchDepth=10, Name="IEMsgView")

    index_pre = message_manager_main_w.TextControl(searchDepth=10)
    refresh_but = index_pre.GetParentControl().GetChildren()[1].GetChildren()[0]
    refresh_but.Click()

    messages = []
    for message_info in message_infos_w.GetChildren():
        messages.append(message_info.Name)

    return messages


def get_unread_messages(messages: list):
    """

    :param messages: list
    """
    unread_messages = []
    last_id = record_file.readline()

    # print("last_id", last_id)
    # print("messages", len(messages))

    if len(messages) >= 60:
        next_messages_page()
        return unread_messages[59]

    if last_id is not None and last_id != '':
        last_id_num = int(last_id)
        unread_messages = messages[last_id_num:]

    record_file.seek(0)
    record_file.truncate(0)
    record_file.write(str(len(messages)))
    record_file.seek(0)
    record_file.flush()

    return unread_messages


def next_messages_page():
    message_infos_w = message_manager_main_w.ListControl(searchDepth=10, Name="IEMsgView")
    message_infos_w_p = message_infos_w.GetParentControl()

    next_but = message_infos_w_p.GetChildren()[2].GetChildren()[1].GetChildren()[2]

    # record_file.seek(0)
    # record_file.truncate(0)
    # record_file.write("0")
    # record_file.seek(0)
    # record_file.flush()

    next_but.Click()


def format_messages(messages: list):
    """
    :param messages: list
    """

    format_messages_ = []

    for message in messages:

        pattern = r"(.+)\((\d{10})\)(\d?\d:\d{2}:\d{2})(.+)"
        sender, qq, time_, info = re.split(pattern, message)[1:-1]
        info_r = None
        info_i = None
        if info[0] == '@':
            info_r = info.split(' ')[0][1:]
            info_i = info.split(' ')[1]

        message_dict = {'sender': sender, 'qq': qq, 'info': info, 'time': time_, "@": info_r, "@i": info_i}
        format_messages_.append(message_dict)

    return format_messages_


def open_message_manager_to_group(Title):
    sub_main_window = auto.WindowControl(searchDepth=1, Name=Title)
    active_window(sub_main_window)

    group_chat = sub_main_window.TabItemControl(searchDepth=10, Name="群聊")
    group_chat.Click()

    return sub_main_window


def open_qq_from_tools_bar(QQ_Tool_Name_):
    # 从右下角打开QQ
    qq_t_name = QQ_Tool_Name_

    tool_w = auto.ToolBarControl(searchDepth=10, Name="用户提示通知区域")
    qq_t_w = tool_w.ButtonControl(searchDepth=2, Name=qq_t_name)
    qq_t_w.DoubleClick()


def open_chat_window_from_qq_chat_list(Chat_Name):
    # 从QQ列表打开会话窗口
    _window = auto.PaneControl(searchDepth=15, Name='会话列表')
    sub_window = _window.PaneControl(searchDepth=5, Name='会话列表')

    chat_list_w = sub_window.GetChildren()[0].GetChildren()[0].GetChildren()[0]
    need_w = chat_list_w.ListItemControl(searchDepth=1, Name=Chat_Name)
    need_w.DoubleClick()


def At(who):
    sub_main_window = auto.WindowControl(searchDepth=1, Name=args.sender)
    sub_main_window.SetActive()
    a = sub_main_window.ListItemControl(searchDepth=16, Name=who)
    a.MoveCursorToMyCenter()
    time.sleep(0.2)
    a.RightClick()
    b = auto.MenuControl(searchDepth=1, Name="TXMenuWindow", ClassName="TXGuiFoundation")
    c = b.MenuItemControl(searchDepth=1, Name="@ TA")
    c.Click()


# 初始化
auto.uiautomation.SetGlobalSearchTimeout(10)

# open_qq_from_tools_bar('QQ: robot(QQ号)\r\n声音: 开启\r\n消息提醒框: 关闭\r\n会话消息: 任务栏头像闪动')
# open_chat_window_from_qq_chat_list(args.sender)

message_manager_main_w = open_message_manager_to_group("消息管理器")
to_special_group(args.sender)

edge_loop = asyncio.get_event_loop()

handle_sender = win32gui.FindWindow("TXGuiFoundation", args.sender)
if handle_sender == 0:
    print("未找到句柄，请确保窗口名合适,或者重新打开窗口")
    exit(-1)

with open(args.cookies, 'r', encoding='utf-8-sig') as f:
    cookies = json.load(f)
    for cookie in cookies:
        if cookie.get("name") == "_U":
            cookies_U = cookie.get("value")
            break

bot = Chatbot(cookies=cookies)

init = 0

# 抢占文件，保证多个机器人运行
sit_path = path + "bingchat_sit.txt"
sit_path_lock = path + "d:/bingchat_sit.txt.lock"

if not os.path.exists(save_img_path):
    os.makedirs(save_img_path)
open(idea_, 'w')
open(say, 'w')
open(record_, 'w')

record_file = open(record_, "r+")
orgin_idea_ = open(orgin_idea, 'w', encoding='utf-8')
if not os.path.exists(record_):
    open(record_, "w")

if __name__ == "__main__":
    isMutil = False

    while True:
        if isMutil:
            time.sleep(0.1)
            to_special_group(args.sender)
            if not os.path.exists(sit_path):
                open(sit_path, "w")
                with FileLock(sit_path) as fl:
                    task(fl)
        else:
            time.sleep(0.1)
            task(None)
