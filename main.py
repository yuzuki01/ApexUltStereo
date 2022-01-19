from PyQt5 import QtWidgets
from PIL import Image, UnidentifiedImageError
from win32gui import FindWindow
from playsound import playsound
import os
import sys
import time
import json
import random


with open('screen_dpi.json', 'r') as fp:
    js_data = json.load(fp)
    width, height = js_data['width'], js_data['height']
random.seed(int(time.time() * 1000))
INTERVAL = 0.5
MUSIC_LIST = []
for fp in os.listdir('assets'):
    if fp.split('.')[-1] == 'wav':
        MUSIC_LIST.append(fp)
if len(MUSIC_LIST) == 0:
    raise FileNotFoundError('没有找到可以播放的音乐')
last_play_num = random.randint(0, len(MUSIC_LIST) - 1)
PLAY_LIST = [last_play_num]
with open('data/valkyrie.json', 'r') as fp:
    js_data = json.load(fp)
if '%sx%s' % (width, height) in js_data:
    dpi = 1, 1
else:
    print('游戏窗口分辨率并不是1920x1080，图像识别可能出现错位导致程序不发挥功能')
    dpi = width/1920, height/1080

# PyQt5
app = QtWidgets.QApplication(sys.argv)
screen = QtWidgets.QApplication.primaryScreen()


def value_in_range(value, value_range):
    if value_range[0] <= value <= value_range[1]:
        return True
    else:
        return False


def image_check(state, image, js):
    for x, y, r_range, g_range, b_range in js["1920x1080"][state]:
        x, y = int(x * dpi[0]), int(y * dpi[1])
        try:
            r, g, b = image.getpixel((x, y))
        except IndexError as e:
            print('游戏窗口分辨率可能改变，请重新设置')
            return exit(-1)
        if not (value_in_range(r, r_range), value_in_range(g, g_range), value_in_range(b, b_range)) == (
                True, True, True):
            return False
    return True


def get_game_window():
    hwnd = FindWindow(None, 'Apex Legends')
    return hwnd


def play_music():
    global PLAY_LIST
    if len(PLAY_LIST) <= 1:
        create_play_num_list()
    else:
        PLAY_LIST = PLAY_LIST[1:]
    music_path = '%s\\assets\\%s' % (os.path.abspath('.'), MUSIC_LIST[PLAY_LIST[0]])
    playsound(music_path)


def create_play_num_list():
    music_num = len(MUSIC_LIST)
    while True:
        if len(PLAY_LIST) < len(MUSIC_LIST):
            n = random.randint(0, music_num - 1)
            if n in PLAY_LIST:
                pass
            else:
                PLAY_LIST.append(n)
        else:
            break


def init():
    while True:
        game_window = get_game_window()
        if game_window != 0:
            break
        print('Apex未启动')
        time.sleep(INTERVAL * 10)
    print('主循环运行')
    return game_window


def main():
    game_window = init()
    ult_ready = 0
    ult_charge = 0
    interval = INTERVAL
    while True:
        try:
            img = Image.fromqimage(screen.grabWindow(game_window).toImage())
        except UnidentifiedImageError:
            return init()
        if image_check('ult_is_ready', img, js_data):
            ult_ready = 100
        else:
            ult_ready = ult_ready / 2
        if ult_ready >= 1.5:
            if image_check('ult_charge_ready', img, js_data):
                ult_ready = 100
                ult_charge = 10000
                interval = INTERVAL / 4
            else:
                ult_charge = ult_charge / 1.1
        else:
            ult_charge = ult_charge / 1.1
        if ult_charge >= 50.00:
            if image_check('ult_is_flying', img, js_data):
                # play
                play_music()
                ult_ready = 0
                ult_charge = 0
                img.close()
                del img
                interval = INTERVAL
                continue
        else:
            interval = INTERVAL

        time.sleep(interval)
        img.close()
        del img


if __name__ == '__main__':
    main()
