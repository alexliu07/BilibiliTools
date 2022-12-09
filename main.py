import os
import tkinter
from tkinter import filedialog

import filetype
import requests
from easygui import choicebox, enterbox

from tools import crc2uid
from tools.bulletchat import bulletchat

chats = []
if not os.path.exists('download'):
    os.mkdir('download')


def getBulletChat(bv, p):
    print('获取弹幕中...')
    # 获取cid
    cid_result = requests.get('https://api.bilibili.com/x/player/pagelist?bvid=' + bv).json()
    cid = cid_result['data'][p - 1]['cid']
    # 写入弹幕文件
    xml_b = requests.get('https://comment.bilibili.com/' + str(cid) + '.xml').content
    xml_f = open('download/tmp.xml', 'wb')
    xml_f.write(xml_b)
    xml_f.close()


def getName(bv, p):
    print('获取视频标题中...')
    # 获取视频标题
    title_result = requests.get('https://api.szfx.top/bilibili/api.php?bv=' + bv).json()
    name = title_result['title']
    # 如果是分p的还要进一步改名
    # 获取分p名
    pname_result = requests.get('https://api.bilibili.com/x/player/pagelist?bvid=' + bv).json()
    if len(pname_result['data']) != 1:
        name = name + ' (P' + str(p) + '. ' + pname_result['data'][p - 1]['part'] + ')'
    name = name.replace('/', '-').replace('[', '(').replace(']', ')')
    return name


def downloadVideo(bv, p, name):
    print('下载视频中...')
    if not os.path.exists('download/' + name + '.mp4'):
        os.system('you-get -o download https://www.bilibili.com/video/' + bv + '?p=' + str(p))
        os.system(r'del "download\\' + name + '.cmt.xml"')
        # 如果是flv格式的直接改成mp4格式的
        if os.path.exists('download/' + name + '.flv'):
            os.chdir('download')
            os.system('ren "' + name + '.flv" "' + name + '.mp4"')
            os.chdir('..')
    else:
        print('文件已存在')


# ----------------------------------------------------------------
# 隐藏窗口
win = tkinter.Tk()
win.withdraw()
# 选择模式
mode = choicebox("请选择工作模式", "工作模式",
                 ['下载视频和弹幕(单独存放)，并进行弹幕分析', '下载视频并嵌入弹幕，并进行弹幕分析', '将本地弹幕嵌入本地视频，并分析弹幕', '仅分析弹幕', '分析本地弹幕',
                  '解密用户uid(被CRC32加密)', 'FFMPEG将任意格式批量转换为MP4'])
if mode == '下载视频和弹幕(单独存放)，并进行弹幕分析':
    # 输入BV号
    bv = enterbox("请输入视频的BV号：", "BV号")
    # 输入分集
    part = enterbox("请输入视频分集的集数(留空代表第一集)：", "集数")
    if part == None or part == '':
        p = 1
    else:
        p = int(part)
    # 获取视频名
    name = getName(bv, p)
    # 下载视频及弹幕
    downloadVideo(bv, p, name)
    getBulletChat(bv, p)
    # 分析弹幕
    print('分析弹幕中...')
    if not os.path.exists('download/' + name + ' 弹幕信息.xlsx'):
        bulletchat('download/tmp.xml', name)
    else:
        print('文件已存在')
    # 播放
    os.system(r'del download\\tmp.xml')
    print('程序执行完毕！')
elif mode == '下载视频并嵌入弹幕，并进行弹幕分析':
    # 输入BV号
    bv = enterbox("请输入视频的BV号：", "BV号")
    # 输入分集
    part = enterbox("请输入视频分集的集数(留空代表第一集)：", "集数")
    if part == None or part == '':
        p = 1
    else:
        p = int(part)
    # 获取视频标题
    name = getName(bv, p)
    # 下载视频及弹幕
    downloadVideo(bv, p, name)
    getBulletChat(bv, p)
    # 分析弹幕
    print('分析弹幕中...')
    if not os.path.exists('download/' + name + ' 弹幕信息.xlsx'):
        bulletchat('download/tmp.xml', name)
    else:
        print('文件已存在')
    # 转换弹幕
    os.chdir('download')
    print("转换弹幕中...")
    if not os.path.exists('tmp.ass'):
        os.system(r'..\ffmpeg\xml2ass.exe tmp.xml')
    else:
        print('文件已存在')
    # 嵌入弹幕
    print("嵌入弹幕中...")
    if not os.path.exists(name + '(弹幕嵌入).mp4'):
        os.system(
            r'..\ffmpeg\ffmpeg.exe -i "' + name + '.mp4" -vf subtitles=tmp.ass -vcodec libx264 "' + name + '(弹幕嵌入).mp4"')
    else:
        print('文件已存在')
    # 删除临时文件
    print("删除临时文件中...")
    os.system('del tmp.ass')
    # 播放
    os.chdir('..')
    os.system(r'del download\\tmp.xml')
    print("程序执行完毕！")
elif mode == '将本地弹幕嵌入本地视频，并分析弹幕':
    chatpath = filedialog.askopenfilename(title='请选择弹幕文件', filetype=[('XML File', '.xml')])
    videopath = filedialog.askopenfilename(title='请选择视频文件', filetype=[('MP4 File', '.mp4')])
    name = os.path.splitext(os.path.split(videopath)[1])[0]
    chatname = os.path.splitext(os.path.split(chatpath)[1])[0]
    # 分析弹幕
    print('分析弹幕中...')
    if not os.path.exists('download/' + name + ' 弹幕信息.xlsx'):
        bulletchat(chatpath, name)
    else:
        print('文件已存在')
    # 复制弹幕及视频
    print('复制弹幕及视频中...')
    os.chdir('download')
    cmdchat = 'copy "' + chatpath + '" tmp.xml'
    cmdchat = cmdchat.replace('/', r'\\')
    cmdvideo = 'copy "' + videopath + '" tmp.mp4'
    cmdvideo = cmdvideo.replace('/', r'\\')
    os.system(cmdchat)
    os.system(cmdvideo)
    # 转换弹幕
    print("转换弹幕中...")
    if not os.path.exists('tmp.ass'):
        os.system(r'..\ffmpeg\xml2ass.exe tmp.xml')
    else:
        print('文件已存在')
    # 嵌入弹幕
    print("嵌入弹幕中...")
    if not os.path.exists(name + '(弹幕嵌入).mp4'):
        os.system(r'..\ffmpeg\ffmpeg.exe -i tmp.mp4 -vf subtitles=tmp.ass -vcodec libx264 "' + name + '(弹幕嵌入).mp4"')
    else:
        print('文件已存在')
    # 删除临时文件
    print("删除临时文件中...")
    os.system('del tmp.ass')
    os.system('del tmp.mp4')
    # 播放
    os.chdir('..')
    os.system(r'del download\\tmp.xml')
    print('程序执行完毕！')
elif mode == '仅分析弹幕':
    # 输入BV号
    bv = enterbox("请输入视频的BV号：", "BV号")
    # 输入分集
    part = enterbox("请输入视频分集的集数(留空代表第一集)：", "集数")
    if part == None or part == '':
        p = 1
    else:
        p = int(part)
    # 获取标题
    name = getName(bv, p)
    # 获取弹幕
    getBulletChat(bv, p)
    # 分析弹幕
    print('分析弹幕中...')
    if not os.path.exists('download/' + name + ' 弹幕信息.xlsx'):
        bulletchat('download/tmp.xml', name)
    else:
        print('文件已存在')
    os.system(r'del download\\tmp.xml')
elif mode == '分析本地弹幕':
    # 选择文件
    path = filedialog.askopenfilename(title='请选择弹幕文件', filetype=[('XML File', '.xml')])
    pather = os.path.split(path)[0] + '/'
    name = os.path.splitext(os.path.split(path)[1])[0]
    # 分析弹幕
    print('分析弹幕中...')
    if not os.path.exists('download/' + name + ' 弹幕信息.xlsx'):
        bulletchat(path, name)
    else:
        print('文件已存在')
elif mode == '解密用户uid(被CRC32加密)':
    crc_uid = enterbox("请输入被crc32加密的用户uid：", "输入加密数据")
    uid = crc2uid.crc2uid(crc_uid)
    print(uid)
    # 获取uid信息
    user_result = requests.get('https://tenapi.cn/bilibili/?uid=' + uid).json()
    user_name = user_result['data']['name']
    user_level = user_result['data']['level']
    user_sex = user_result['data']['sex']
    user_description = user_result['data']['description']
    text = "-----------用户信息-----------\nuid：{0}\n主页：https://space.bilibili.com/{0}\n名称：{1}\n等级：Lv.{2}\n性别：{3}\n简介：{4}\n------------------------------".format(
        uid, user_name, user_level, user_sex, user_description)
    print(text)
elif mode == 'FFMPEG将任意格式批量转换为MP4':
    # 询问目录
    path = filedialog.askdirectory(title='请选择要批量转换的目录(不会转换子目录)')
    # 记录当前目录
    curPath = os.getcwd()
    os.chdir(path)
    # 遍历
    for root, dirs, files in os.walk('.'):
        if root == '.':
            for i in files:
                # 获取文件类型
                predict = filetype.guess(i)
                if not predict:
                    continue
                type = predict.mime
                exts = os.path.splitext(i)[1]
                # 判断是否为视频并转换
                if type[:5] == 'video' and exts != '.mp4':
                    print('正在转换', i)
                    filename = os.path.splitext(i)[0]
                    os.system(curPath + r'\ffmpeg\ffmpeg.exe -i ' + i + ' -vcodec libx264 ' + filename + '.mp4')
            break
