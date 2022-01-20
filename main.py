import threading
import random
import time
import os
from win32com.client import Dispatch

# global
MEDIA_PLAYER = Dispatch('WMPlayer.OCX')
random.seed(int(time.time() * 1000))


class Window:
    import json
    import tkinter as tk
    from tkinter import ttk

    INTERVAL_LENGTH = (0.5, 0.75, 1.0)
    JS_DICT = {}
    for fn in os.listdir('data'):
        if fn.split('.')[-1] == 'json':
            with open('data/%s' % fn, 'r') as fp:
                js = json.load(fp)
            if js['id'] == fn.split('.')[0]:
                JS_DICT[js['id']] = 'data/%s' % fn

    def __init__(self):
        self.window = self.tk.Tk()
        self.window.iconbitmap('data/icon.ico')
        self.window.title('ApexUltStereo')
        self.window.minsize(330, 350)  # 最小尺寸
        self.window.maxsize(330, 350)
        self.font = {
            'normal': ('Arial', 12),
            'small': ('Arial', 10)
        }
        self.label_text = {}
        self.label = {}
        for status in ('RUNTIME_STATUS', 'ULT_STATUS'):
            self.label_text[status] = self.tk.StringVar()
            self.label[status] = self.tk.Label(self.window, textvariable=self.label_text[status],
                                               font=self.font['normal'])
            self.label[status].pack()
        self.tk.Label(self.window, text='音量(%)', font=self.font['small']).pack()
        self.VOLUME_SCALE = self.tk.Scale(self.window, from_=0, to=100, tickinterval=10, command=self.SET_VOLUME,
                                          orient=self.tk.HORIZONTAL, length=220)
        self.VOLUME_SCALE.set(80)
        self.VOLUME_SCALE.pack()
        self.label_text['INTERVAL_TEXT'] = self.tk.StringVar()
        self.tk.Label(self.window, textvariable=self.label_text['INTERVAL_TEXT'], font=self.font['small']).pack()
        self.INTERVAL_SCALE = self.tk.Scale(self.window, from_=0, to=2, tickinterval=1, command=self.SET_INTERVAL,
                                            orient=self.tk.HORIZONTAL, length=150)
        self.INTERVAL_SCALE.set(1)
        self.label_text['INTERVAL_TEXT'].set('检测间隔: %.2f s' % self.INTERVAL_LENGTH[1])
        self.INTERVAL_SCALE.pack()
        self.tk.Label(self.window, text='检测间隔过小可能会导致卡顿\n过大可能会导致不灵敏',
                      font=self.font['small'], relief='groove').pack()
        self.label_text['SELECTOR'] = self.tk.StringVar()
        self.tk.Label(self.window, textvariable=self.label_text['SELECTOR'], font=self.font['normal']).pack()
        self.label_text['SELECTOR'].set('选择识别的英雄 (当前: %s)' % [i for i in self.JS_DICT][0])

        self.selector = self.ttk.Combobox(self.window, state='readonly')
        self.selector['value'] = [i for i in self.JS_DICT]
        self.selector.current(0)
        self.JS_PATH = self.JS_DICT[[i for i in self.JS_DICT][0]]
        self.selector.bind("<<ComboboxSelected>>", self.SELECT_JS_DATA)
        self.selector.pack()

        self.watch_dog = WatchDog(GUI=self)
        self.app = App(GUI=self, WATCH_DOG=self.watch_dog, JS_DATA=self.GET_JSON(self.JS_PATH))
        self.watch_dog.bind_watch_thread(self.app)

    def run(self):
        self.app.setDaemon(True)
        self.app.start()
        self.watch_dog.setDaemon(True)
        self.watch_dog.start()
        self.window.mainloop()

    def SET_VOLUME(self, value):
        MEDIA_PLAYER.settings.volume = int(value)
        self.app.MEDIA_VOLUME = int(value)

    def SET_INTERVAL(self, value):
        value = int(value)
        self.app.INTERVAL = self.INTERVAL_LENGTH[value]
        self.label_text['INTERVAL_TEXT'].set('检测间隔: %.2f s' % self.INTERVAL_LENGTH[value])

    def SET_LABEL(self, state, text):
        self.label_text[state].set(text)

    def GET_JSON(self, path):
        with open(path, 'r') as fp:
            js = self.json.load(fp)
        return js

    def SELECT_JS_DATA(self, event):
        if self.JS_PATH == self.JS_DICT[self.selector.get()]:
            return
        self.JS_PATH = self.JS_DICT[self.selector.get()]
        self.label_text['SELECTOR'].set('选择识别的英雄 (当前: %s)' % self.selector.get())
        self.watch_dog.set_keep_watch(False)
        self.app.flag = False
        while not self.app.thread_is_sleeping:
            time.sleep(0.1)
        self.app.flag = True
        self.app.restart(self.GET_JSON(self.JS_PATH))
        time.sleep(0.5)


class WatchDog(threading.Thread):
    MAX_WAIT_TIME = 0.040
    last_feed = None
    watch_thread = None
    keep_watch = True
    flag = True
    thread_is_sleeping = False

    def __init__(self, GUI):
        GUI: window
        watch_thread: App
        self.GUI = GUI
        threading.Thread.__init__(self, name='WatchDog_Thread')

    def run(self):
        self.main()

    def main(self):
        while self.flag:
            time.sleep(1)
        return self.thread_sleep()

    def thread_sleep(self):
        self.thread_is_sleeping = True
        while not self.thread_is_sleeping:
            time.sleep(0.1)
        return self.main()

    def thread_awake(self):
        self.thread_is_sleeping = False

    def stop(self):
        self.flag = False

    def bind_watch_thread(self, handler):
        self.watch_thread = handler

    def set_keep_watch(self, value):
        if value:
            self.feed(time.time())
        self.keep_watch = value

    def feed(self, t):
        self.last_feed = t

    def active(self):
        # print('WATCH DOG ACTIVE')
        self.watch_thread.thread_sleep()
        self.last_feed = None
        self.GUI.SET_LABEL('ULT_STATUS', '触发看门狗\n')
        self.watch_thread.flag = False
        time.sleep(1)
        while not self.watch_thread.thread_is_sleeping:
            time.sleep(0.1)
        self.GUI.SET_LABEL('ULT_STATUS', '线程重启\n')
        self.watch_thread.flag = True
        self.watch_thread.restart(JS_DATA=self.watch_thread.JS_DATA)

    def check(self, WAIT_TIME=None):
        if not WAIT_TIME:
            WAIT_TIME = self.MAX_WAIT_TIME
        if self.last_feed is None:
            pass
        else:
            if time.time() - self.last_feed >= WAIT_TIME and self.watch_thread and self.keep_watch:
                self.active()


class App(threading.Thread):
    import win32gui
    import win32con
    import win32ui
    from PIL import Image

    flag = True
    thread_is_sleeping = False
    MEDIA_PLAYER.settings.volume = 80
    MEDIA_VOLUME = 80
    INTERVAL = 0.50
    MUSIC_LIST = {
        'ult': [],
        'para': [],
        'uav': []
    }
    PLAY_LIST = {}
    for STATE in ('ult', 'para', 'uav'):
        for fp in os.listdir('assets/%s' % STATE):
            if fp.split('.')[-1] in ('wav', 'mp3'):
                MUSIC_LIST[STATE].append(fp)
        if len(MUSIC_LIST[STATE]) > 0:
            last_play_num = random.randint(0, len(MUSIC_LIST[STATE]) - 1)
            PLAY_LIST[STATE] = [last_play_num]
        else:
            PLAY_LIST[STATE] = []

    def __init__(self, GUI, WATCH_DOG, JS_DATA):
        GUI: Window
        WATCH_DOG: WatchDog
        self.GUI = GUI
        self.WATCH_DOG = WATCH_DOG
        self.JS_DATA = JS_DATA
        threading.Thread.__init__(self, name='Func_thread')
        self.GUI.SET_LABEL('ULT_STATUS', 'Null\n')
        self.GUI.SET_LABEL('RUNTIME_STATUS', 'Null\n')
        self.FUNC_DICT = {
            'valkyrie': self.VALKYRIE,
            'crypto': self.CRYPTO
        }

    def run(self):
        self.mainloop()

    def getHWND(self):
        while True:
            hwnd = self.win32gui.FindWindow(None, "Apex Legends")
            if hwnd != 0:
                break
            self.GUI.SET_LABEL('RUNTIME_STATUS', '未捕获到 Apex 窗口')
            time.sleep(self.INTERVAL)
        rect = self.win32gui.GetWindowRect(hwnd)
        width, height = rect[2] - rect[0], rect[3] - rect[1]
        self.GUI.SET_LABEL('RUNTIME_STATUS', '捕获到Apex窗口:  (cost: NaN s)\nid=%s, size=(%d, %d)' % (hwnd, width, height))
        return hwnd, (width, height)

    def getImage(self, HWND, ScreenSize):
        hWndDC = self.win32gui.GetWindowDC(HWND)
        mfcDC = self.win32ui.CreateDCFromHandle(hWndDC)
        BitMap = self.win32ui.CreateBitmap()
        BitMap.CreateCompatibleBitmap(mfcDC, *ScreenSize)
        saveDC = mfcDC.CreateCompatibleDC()
        saveDC.SelectObject(BitMap)
        saveDC.BitBlt((0, 0), ScreenSize, mfcDC, (0, 0), self.win32con.SRCCOPY)
        BitMap_info = BitMap.GetInfo()
        BitMap_str = BitMap.GetBitmapBits(True)
        img_PIL = self.Image.frombuffer('RGB', (BitMap_info['bmWidth'], BitMap_info['bmHeight']), BitMap_str, 'raw',
                                        'BGRX', 0, 1)
        del BitMap, saveDC, BitMap_str, BitMap_info, hWndDC, mfcDC
        return img_PIL

    @staticmethod
    def value_in_range(value, value_range):
        if value_range[0] <= value <= value_range[1]:
            return True
        else:
            return False

    def image_check(self, state, image, js, dpi):
        for x, y, r_range, g_range, b_range in js[dpi[0]][state]:
            x, y = int(x * dpi[1]), int(y * dpi[2])
            try:
                r, g, b = image.getpixel((x, y))
            except IndexError:
                self.GUI.SET_LABEL('RUNTIME_STATUS', '游戏窗口分辨率可能改变\n请重新启动本程序')
                return exit(0)
            if not (self.value_in_range(r, r_range), self.value_in_range(g, g_range),
                    self.value_in_range(b, b_range)) == (
                           True, True, True):
                return False
        return True

    def play_music(self, state, hwnd=None, window_size=None, dpi=None, js_data=None):
        if len(self.PLAY_LIST[state]) <= 1:
            if len(self.PLAY_LIST[state]) == 0:
                self.GUI.SET_LABEL('ULT_STATUS', '无音乐播放\n')
                time.sleep(1)
                return
            self.create_play_num_list(state)
        else:
            self.PLAY_LIST[state] = self.PLAY_LIST[state][1:]
        music_name = self.MUSIC_LIST[state][self.PLAY_LIST[state][0]]
        music_path = '%s\\assets\\%s\\%s' % (
            os.path.abspath('.'), state, music_name)
        play_mode = 0
        if state == 'ult':
            music_str = '大招'
        elif state == 'para':
            music_str = '跳伞'
        elif state == 'uav':
            music_str = '无人机'
            play_mode = 1
        else:
            raise ValueError(state)
        if play_mode == 0:
            # valkyrie
            self.GUI.SET_LABEL('ULT_STATUS', '播放%s音乐\nassets\\%s\\%s' % (music_str, state, music_name))
            media = MEDIA_PLAYER.newMedia(music_path)
            MEDIA_PLAYER.currentPlaylist.appendItem(media)
            MEDIA_PLAYER.controls.play()
            time.sleep(self.INTERVAL)
            MEDIA_PLAYER.controls.playItem(media)
            time.sleep(media.duration)
            self.GUI.SET_LABEL('ULT_STATUS', '播放器冷却中\n')
            time.sleep(7.5)
            MEDIA_PLAYER.currentPlaylist.removeItem(media)
        elif play_mode == 1:
            # crypto
            player_thread = PlayerThread(app=self)
            player_thread.play()
            time.sleep(self.INTERVAL)
            while True:
                st = time.time()
                self.WATCH_DOG.set_keep_watch(True)
                time.sleep(self.INTERVAL)
                img = self.fetch_window_image(hwnd, window_size)
                if self.image_check('uav', img, js_data, dpi):
                    # enter uav playing thread
                    img.close()
                    del img
                    cost = time.time() - st
                    self.GUI.SET_LABEL('RUNTIME_STATUS',
                                       '捕获到Apex窗口:  (cost: %.3f s)\nid=%s, size=(%d, %d)' % (cost, hwnd, *window_size))
                    self.WATCH_DOG.check(1.0)
                    time.sleep(self.INTERVAL)
                    continue
                player_thread.stop()
                self.WATCH_DOG.set_keep_watch(False)
                img.close()
                del img
                break
            time.sleep(self.INTERVAL)
            self.GUI.SET_LABEL('ULT_STATUS', '播放器冷却中\n')
            time.sleep(1.5)

    def create_play_num_list(self, state):
        music_num = len(self.MUSIC_LIST[state])
        while True:
            if len(self.PLAY_LIST[state]) < len(self.MUSIC_LIST[state]):
                n = random.randint(0, music_num - 1)
                if n in self.PLAY_LIST[state]:
                    pass
                else:
                    self.PLAY_LIST[state].append(n)
            else:
                break

    def init(self):
        hwnd, window_size = self.getHWND()
        dpi_str = '%dx%d' % window_size[:2]
        if dpi_str in self.JS_DATA:
            dpi = dpi_str, 1, 1
        else:
            dpi = '1920x1080', window_size[0] / 1920, window_size[1] / 1080
        return hwnd, window_size, dpi, self.JS_DATA

    def mainloop(self):
        hwnd, window_size, dpi, js_data = self.init()
        self.FUNC_DICT[self.JS_DATA['id']](hwnd, window_size, dpi, js_data)
        return self.thread_sleep()

    def thread_sleep(self):
        self.thread_is_sleeping = True
        while self.thread_is_sleeping:
            time.sleep(0.5)
        return self.mainloop()

    def thread_awake(self):
        self.thread_is_sleeping = False

    def restart(self, JS_DATA):
        MEDIA_PLAYER.controls.stop()
        self.flag = True
        self.JS_DATA = JS_DATA
        self.thread_awake()

    def fetch_window_image(self, hwnd, window_size):
        try:
            img = self.getImage(hwnd, window_size)
        except Exception as e:
            self.GUI.SET_LABEL('RUNTIME_STATUS', 'Error: %s\n' % e.strerror)
            time.sleep(self.INTERVAL)
            return self.mainloop()
        return img

    def VALKYRIE(self, hwnd, window_size, dpi, js_data):
        ult_charge = 0
        interval = self.INTERVAL
        self.GUI.SET_LABEL('ULT_STATUS', '大招未使用\n')
        while self.flag:
            coff = 0.5 / self.INTERVAL
            st = time.time()
            self.WATCH_DOG.set_keep_watch(True)
            img = self.fetch_window_image(hwnd, window_size)

            if self.image_check('ult_charge_ready', img, js_data, dpi):
                ult_charge = 10000 * coff
                interval = self.INTERVAL / 2
                self.GUI.SET_LABEL('ULT_STATUS', '大招使用中\n')
            else:
                ult_charge = ult_charge / 1.1
            if ult_charge >= 40.0:
                if self.image_check('ult_is_flying', img, js_data, dpi):
                    # ult-play
                    self.WATCH_DOG.set_keep_watch(False)
                    self.play_music('ult')
                    ult_charge = 0
                    img.close()
                    del img
                    interval = self.INTERVAL
                    time.sleep(interval)
                    continue
            else:
                self.GUI.SET_LABEL('ULT_STATUS', '大招未使用\n')
                if self.image_check('ult_is_flying', img, js_data, dpi):
                    # para-play
                    self.WATCH_DOG.set_keep_watch(False)
                    self.play_music('para')
                    ult_charge = 0
                    img.close()
                    del img
                    interval = self.INTERVAL
                    time.sleep(interval)
                    continue

            img.close()
            del img
            cost = time.time() - st
            self.GUI.SET_LABEL('RUNTIME_STATUS',
                               '捕获到Apex窗口:  (cost: %.3f s)\nid=%s, size=(%d, %d)' % (cost, hwnd, *window_size))
            self.WATCH_DOG.check()
            self.WATCH_DOG.set_keep_watch(False)
            time.sleep(interval)

    def CRYPTO(self, hwnd, window_size, dpi, js_data):
        self.GUI.SET_LABEL('ULT_STATUS', '无人机未使用\n')
        while self.flag:
            st = time.time()
            self.WATCH_DOG.set_keep_watch(True)
            img = self.fetch_window_image(hwnd, window_size)

            if self.image_check('uav', img, js_data, dpi):
                # enter uav playing thread
                self.WATCH_DOG.set_keep_watch(False)
                self.play_music('uav', hwnd, window_size, dpi, js_data)
                img.close()
                del img
                continue
            img.close()
            del img
            cost = time.time() - st
            self.GUI.SET_LABEL('RUNTIME_STATUS',
                               '捕获到Apex窗口:  (cost: %.3f s)\nid=%s, size=(%d, %d)' % (cost, hwnd, *window_size))
            self.WATCH_DOG.check()
            self.WATCH_DOG.set_keep_watch(False)
            time.sleep(self.INTERVAL)


class PlayerThread(threading.Thread):
    flag = True
    is_playing = False

    def __init__(self, app):
        self.app = app
        self.GUI = self.app.GUI
        self.media = None
        threading.Thread.__init__(self)

    def play(self):
        self.flag = True
        self.is_playing = True
        self.setDaemon(True)
        self.start()

    def run(self):
        self.main()

    def main(self):
        state = 'uav'
        while self.flag:
            self.is_playing = True
            if len(self.app.PLAY_LIST[state]) <= 1:
                if len(self.app.PLAY_LIST[state]) == 0:
                    self.GUI.SET_LABEL('ULT_STATUS', '无音乐播放\n')
                    time.sleep(1)
                    return 0
                self.app.create_play_num_list(state)
            else:
                self.app.PLAY_LIST[state] = self.app.PLAY_LIST[state][1:]
            music_name = self.app.MUSIC_LIST[state][self.app.PLAY_LIST[state][0]]
            music_path = '%s\\assets\\%s\\%s' % (
                os.path.abspath('.'), state, music_name)
            self.GUI.SET_LABEL('ULT_STATUS', '播放%s音乐\nassets\\%s\\%s' % ('无人机', state, music_name))
            self.media = MEDIA_PLAYER.newMedia(music_path)
            MEDIA_PLAYER.currentPlaylist.appendItem(self.media)
            MEDIA_PLAYER.controls.play()
            time.sleep(self.app.INTERVAL)
            MEDIA_PLAYER.controls.playItem(self.media)
            time.sleep(self.media.duration)
            time.sleep(self.app.INTERVAL)
            MEDIA_PLAYER.currentPlaylist.removeItem(self.media)
            self.media = None
        return self.thread_close()

    def thread_close(self):
        self.is_playing = False
        return 0

    def stop(self):
        self.flag = False
        if self.media is None:
            pass
        else:
            t = 1
            volume = self.app.MEDIA_VOLUME
            while int(volume) > 2:
                volume = self.app.MEDIA_VOLUME // (t + 1)
                MEDIA_PLAYER.settings.volume = volume
                t += 1
                time.sleep(0.1)
            MEDIA_PLAYER.settings.volume = 0
        while self.is_playing:
            time.sleep(0.1)
        MEDIA_PLAYER.controls.stop()
        MEDIA_PLAYER.settings.volume = self.app.MEDIA_VOLUME
        del self


if __name__ == "__main__":
    window = Window()
    window.run()
