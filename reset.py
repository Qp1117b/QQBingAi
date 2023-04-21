# 抢占文件，保证多个机器人运行
import os

sit_path = "d:/bingchat_sit.txt"
sit_path_lock = "d:/bingchat_sit.txt.lock"

if os.path.exists(sit_path):
    os.remove(sit_path)
if os.path.exists(sit_path_lock):
    os.remove(sit_path_lock)