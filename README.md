# Bilibili小工具
## 功能
1. 支持下载Bilibili视频和弹幕
2. 弹幕分析并导出xlsx文件
3. 支持将弹幕硬嵌入视频中并单独保存
4. 将本地弹幕嵌入本地视频
5. 解密被CRC32加密的用户UID，并获取用户信息
6. 批量转换视频格式
***
## 模式
1. 下载视频及弹幕，并分析弹幕
2. 下载视频及弹幕，并嵌入弹幕单独保存及分析弹幕
3. 本地视频嵌入本地弹幕，并分析弹幕
4. 仅分析在线弹幕
5. 仅分析本地弹幕
6. 解密用户UID，并获取用户信息
7. 批量转换本地视频到MP4格式
***
## 使用方法（仅支持Windows系统）
1. 安装<a href="https://www.python.org/">Python</a>
2. 将ffmpeg文件夹下的ffmpeg.7z解压，得到ffmpeg.exe
3. 安装库openpyxl,requests,you-get,easygui,filetype<br>`pip install openpyxl requests you-get easygui filetype`
4. 获取视频BV号
5. 运行main.py并按照提示使用
***
## 使用工具
1. <a href="https://www.python.org/">Python</a>
2. <a href="https://github.com/FFmpeg/FFmpeg">FFmpeg</a>
3. <a href="https://github.com/m13253/danmaku2ass">Danmaku2ass</a>
