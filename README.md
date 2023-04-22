# qq机器人 for BingChat

### 介绍
qq bingchat，qq群检查@机器人信息（使用截图加文字识别的方式），提交给bingchat，自动返回信息

### 安装
```
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```
[EdgeGPT](https://github.com/acheong08/EdgeGPT)  

### 运行
+ 使用cmd，参数意义参考源代码
```
start cmd /c "python bing_chat_for_qq.py -sender 123 -bot bot & pause"
```
+ 也可以使用ide
### 使用说明
 + #### 需要消息管理打开
 + #### 机器人建议使用虚拟机，多个就用多个虚拟机
 + #### 由于自动发送信息需要使用粘贴板因此在使用期间不能粘贴复制，且鼠标也会强制移动
 + #### -bot参数指定当有 @(-bot) 的消息出现时，对消息进行处理
 + #### 需要cookies.json等文件，获取方法参考 [EdgeGPT](https://github.com/acheong08/EdgeGPT)
 + #### 可能需要魔法