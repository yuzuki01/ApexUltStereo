import threading


class Window:
    import tkinter as tk

    def __init__(self):
        self.window = self.tk.Tk()
        self.window.iconbitmap('data/icon.ico')
        self.window.title('ApexUltStereo')
        self.window.minsize(330, 150)  # 最小尺寸
        self.window.maxsize(330, 150)
        self.font = ('Arial', 12)
        self.label_text = {}
        self.label = {}
        for status in ('RUNTIME_STATUS', 'ULT_STATUS'):
            self.label_text[status] = self.tk.StringVar()
            self.label[status] = self.tk.Label(self.window, textvariable=self.label_text[status], font=self.font)
            self.label[status].pack()
        self.VOLUME_SCALE = self.tk.Scale(self.window, from_=0, to=100, tickinterval=10, command=self.SET_VOLUME,
                                          orient=self.tk.HORIZONTAL, length=220)
        self.VOLUME_SCALE.set(80)
        self.VOLUME_SCALE.pack()

        self.app = App(self)

    def run(self):
        self.app.setDaemon(True)
        self.app.start()
        self.window.mainloop()

    def SET_VOLUME(self, value):
        self.app.MEDIA_PLAYER.settings.volume = int(value)

    def SET_LABEL(self, state, text):
        self.label_text[state].set(text)


class App(threading.Thread):
    import json
    import time
    import random
    import os
    import win32gui
    import win32con
    import win32ui
    from win32com.client import Dispatch
    from PIL import Image
    random.seed(int(time.time() * 1000))
    MEDIA_PLAYER = Dispatch('WMPlayer.OCX')
    MEDIA_PLAYER.settings.volume = 80
    INTERVAL = 0.5
    MUSIC_LIST = {
        'ult': [],
        'para': []
    }
    PLAY_LIST = {}
    for STATE in ('ult', 'para'):
        for fp in os.listdir('assets/%s' % STATE):
            if fp.split('.')[-1] in ('wav', 'mp3'):
                MUSIC_LIST[STATE].append(fp)
        if len(MUSIC_LIST[STATE]) == 0:
            raise FileNotFoundError('assets/%s 没有找到可以播放的音乐' % STATE)
        last_play_num = random.randint(0, len(MUSIC_LIST[STATE]) - 1)
        PLAY_LIST[STATE] = [last_play_num]

    def __init__(self, GUI):
        GUI: Window
        self.GUI = GUI
        threading.Thread.__init__(self)
        self.GUI.SET_LABEL('ULT_STATUS', 'Null\n')
        self.GUI.SET_LABEL('RUNTIME_STATUS', 'Null\n')

    def run(self):
        self.mainloop()

    def getJson(self, path):
        with open(path, 'r') as fp:
            js = self.json.load(fp)
        return js

    def getHWND(self):
        while True:
            hwnd = self.win32gui.FindWindow(None, "Apex Legends")
            if hwnd != 0:
                break
            self.GUI.SET_LABEL('RUNTIME_STATUS', '未捕获到 Apex 窗口')
            self.time.sleep(self.INTERVAL)
        rect = self.win32gui.GetWindowRect(hwnd)
        width, height = rect[2] - rect[0], rect[3] - rect[1]
        self.GUI.SET_LABEL('RUNTIME_STATUS', '捕获到Apex窗口:\nid=%s, size=(%d, %d)' % (hwnd, width, height))
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
                                        'BGRX',
                                        0, 1)
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

    def play_music(self, state):
        if len(self.PLAY_LIST[state]) <= 1:
            self.create_play_num_list(state)
        else:
            self.PLAY_LIST[state] = self.PLAY_LIST[state][1:]
        music_name = self.MUSIC_LIST[state][self.PLAY_LIST[state][0]]
        music_path = '%s\\assets\\%s\\%s' % (
            self.os.path.abspath('.'), state, music_name)
        if state == 'ult':
            music_str = '大招'
        else:
            music_str = '跳伞'
        self.GUI.SET_LABEL('ULT_STATUS', '播放%s音乐\nassets\\%s\\%s' % (music_str, state, music_name))
        media = self.MEDIA_PLAYER.newMedia(music_path)
        self.MEDIA_PLAYER.currentPlaylist.appendItem(media)
        self.MEDIA_PLAYER.controls.play()
        self.time.sleep(self.INTERVAL)
        self.MEDIA_PLAYER.controls.playItem(media)
        self.time.sleep(media.duration)
        self.GUI.SET_LABEL('ULT_STATUS', '播放器冷却中\n')
        self.time.sleep(self.INTERVAL * 10)
        self.MEDIA_PLAYER.currentPlaylist.removeItem(media)

    def create_play_num_list(self, state):
        music_num = len(self.MUSIC_LIST[state])
        while True:
            if len(self.PLAY_LIST[state]) < len(self.MUSIC_LIST[state]):
                n = self.random.randint(0, music_num - 1)
                if n in self.PLAY_LIST[state]:
                    pass
                else:
                    self.PLAY_LIST[state].append(n)
            else:
                break

    def init(self):
        hwnd, window_size = self.getHWND()
        js_data = self.getJson('data/valkyrie.json')
        dpi_str = '%dx%d' % window_size[:2]
        if dpi_str in js_data:
            dpi = dpi_str, 1, 1
        else:
            dpi = '1920x1080', window_size[0] / 1920, window_size[1] / 1080
        return hwnd, window_size, dpi, js_data

    def mainloop(self):
        hwnd, window_size, dpi, js_data = self.init()
        ult_charge = 0
        interval = self.INTERVAL
        self.GUI.SET_LABEL('ULT_STATUS', '大招未使用\n')
        while True:
            try:
                img = self.getImage(hwnd, window_size)
            except Exception as e:
                self.GUI.SET_LABEL('RUNTIME_STATUS', 'Error: %s\n' % e.strerror)
                self.time.sleep(interval)
                return self.mainloop()

            if self.image_check('ult_charge_ready', img, js_data, dpi):
                ult_charge = 10000
                interval = self.INTERVAL / 4
                self.GUI.SET_LABEL('ULT_STATUS', '大招使用中\n')
            else:
                ult_charge = ult_charge / 1.1
            if ult_charge >= 40.0:
                if self.image_check('ult_is_flying', img, js_data, dpi):
                    # ult-play
                    self.play_music('ult')
                    ult_charge = 0
                    img.close()
                    del img
                    interval = self.INTERVAL
                    continue
            else:
                self.GUI.SET_LABEL('ULT_STATUS', '大招未使用\n')
                if self.image_check('ult_is_flying', img, js_data, dpi):
                    # para-play
                    self.play_music('para')
                    ult_charge = 0
                    img.close()
                    del img
                    interval = self.INTERVAL
                    continue

            self.time.sleep(interval)
            img.close()
            del img


if __name__ == "__main__":
    window = Window()
    window.run()
