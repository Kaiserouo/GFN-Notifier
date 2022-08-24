from flask import Flask, jsonify
import shell
import atexit
import win32gui, win32con, win32ui, win32api
import time
from PIL import Image
import pytesseract
import collections
import re
import os
import datetime
from ctypes import windll
from pathlib import Path

# --- config ---

# VPN settings
TOGGLE_VPN = True   
VPN_NAME = 'Kaiserouo'
VPN_SERVER_IP = '25.8.47.6'
VPN_SERVER_PORT = 12000
HAMACHI_PATH = r'C:\Program Files (x86)\LogMeIn Hamachi\x64\hamachi-2.exe'

# GFN settings
GFN_SEARCH_STRING = '在 GeForce NOW 上玩'
GFN_START_COMMAND = r'"C:\Users\user\AppData\Local\NVIDIA Corporation\GeForceNOW\CEF\GeForceNOWStreamer.exe"  --url-route="#?cmsId=100901811&launchSource=External&shortName=apex_legends_steam_ww&parentGameId=cb2b1b5f-54ba-45fd-9839-96bbfe1376cd"'

# TeamViewer
TEAMVIEWER_PATH = r'C:\Program Files\TeamViewer\TeamViewer.exe'

# -------------


app = Flask(__name__)

class GFNHwndManager:
    """
        Manages everything regarding hwnd. Especially GFN
    """
    def __init__(self):
        windll.user32.SetProcessDPIAware()

    def _filterWindow(self, filter_fn):
        # given filter function: (hwnd, name) -> bool,
        # return all hwnd that satisfies this condition
        class Handler:
            def __init__(self):
                self.ls = []
            def __call__(self, hwnd, ctx):
                if filter_fn(hwnd, win32gui.GetWindowText(hwnd)):
                    self.ls.append(hwnd)

        hdl = Handler()
        win32gui.EnumWindows(hdl, None)
        return hdl.ls

    def _getScreenshot(self, hwnd):
        # Change the line below depending on whether you want the whole window
        # or just the client area. 
        #left, top, right, bot = win32gui.GetClientRect(hwnd)
        left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top

        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

        saveDC.SelectObject(saveBitMap)

        # Change the line below depending on whether you want the whole window
        # or just the client area. 
        #result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
        PW_RENDERFULLCONTENT = 0x00000002
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), PW_RENDERFULLCONTENT)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        return im

    def getHwnd(self, substr):
        # if is not executing then return -1
        hwnd_ls = self._filterWindow(lambda x,y: substr in y)
        print([win32gui.GetWindowText(hwnd) + f'({hwnd})' for hwnd in hwnd_ls])
        if len(hwnd_ls) != 1:
            return -1
        return hwnd_ls[0]

    def getGFNHwnd(self):
        return self.getHwnd(GFN_SEARCH_STRING)

class GFNViewerTesseract:
    """
        Snapshot the window and use OCR to work out the queue number;
        Not really direct, a bit slow, and may have error.
    """
    def __init__(self):
        self.hwnd_manager = GFNHwndManager()
    
    def getQueueCount(self):
        hwnd = self.hwnd_manager.getGFNHwnd()
        if hwnd < 0:
            # not executing
            return {
                "count": -2,
                "message": "GFN not executing now or not in queue."
            }

        # if it is minimize, make it... not minimized
        if win32gui.IsIconic(hwnd):
            foreground_hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            if foreground_hwnd != 0:
                win32gui.SetForegroundWindow(hwnd)
                
        # get image
        img = self.hwnd_manager._getScreenshot(hwnd)
        w, h = img.size
        # img = img.crop((w // 3, h // 3, 2 * w // 3, 2 * h // 3))
        # img = img.resize((w, h))
        img.save('a.png')

        # OCR
        s = pytesseract.image_to_string(img, lang='chi_tra')
        print(f"result:\n----------------\n{s}\n--------")
        search_result = re.search(': [\\d ]+', s)
        if search_result is None:
                return {
                    "count": -1,
                    "message": "GFN is now in or OCR error."
                }

        d = int(search_result.group(0)[2:].replace(' ', ''))
        return {
            "count": d,
            "message": ""
        }

    def click(self, x, y):
        hwnd = self.hwnd_manager.getGFNHwnd()
        if hwnd < 0:
            return
        lParam = win32api.MAKELONG(x, y)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, None, lParam)

class GFNViewerDebugFile:
    """
        Use GFN debug file to work out the queue count.
        Very direct (should've known this sooner), no error, but will open a 1MB file and process
        everytime a request come in.

        Should close GFN everytime your time had finished, since the debug file will only be cleaned
        when GFN launches. Otherwise it will build up.
    """

    def __init__(self):
        super().__init__()
        self.debug_fpath = Path(os.getenv('LOCALAPPDATA')) / 'NVIDIA Corporation' / 'GeForceNOW' / 'debug.log'
    
    def _tail(self, n=10, fpath=None):
        # Return the last n lines of a file. 
        # For file with 7000 lines, faster than reading the file backward and manually finding `\n`,
        # which is weird considering deque actually reads the whole file.
        # Power of C code I guess.
        if fpath is None:
            fpath = self.debug_fpath
        with fpath.open('r') as fin:
            return collections.deque(fin, n)

    def _readAllLines(self, fpath=None):
        if fpath is None:
            fpath = self.debug_fpath
        with fpath.open('r') as fin:
            return fin.readlines()

    def _parseQueueLine(self, s):
        # Some lines in debug.log look like the following:
        # [2022-08-24/ 10:54:53.659:INFO:simple_grid_app.cc(1422)] onSessionSetupProgress(state: 1, queue: 32, eta: 30000)
        # [2022-08-24/ 10:54:53.659:INFO:simple_grid_app.cc(1422)] onSessionSetupProgress(state: 2, queue: 0, eta: 20000)

        result = re.search("\\[(.+)/[ ]*(.+):INFO:.+\\].+\\(state: (.*), queue: (.*), eta: (.*)\\)", s)
        return {
            "datetime": datetime.datetime.strptime(
                result.group(1) + ' ' + result.group(2),
                "%Y-%m-%d %H:%M:%S.%f"
            ),
            "state": int(result.group(3)),
            "queue": int(result.group(4)),
            "eta": int(result.group(5)),
        }

    def _filterQueueLine(self, s_ls):
        return [s for s in s_ls if "onSessionSetupProgress" in s]
    
    def _parseQueue(self, s_ls):
        return [self._parseQueueLine(s) for s in s_ls]

    def getQueueCount(self):
        # get queue (tail 200 lines, roughly 5.7ms for 7500 line file)
        queue_ls = self._parseQueue(self._filterQueueLine(self._tail(n=200)))
        if len(queue_ls) == 0:
            # just fall back to reading the whole file (roughly 40ms for 7500 line file)
            queue_ls = self._parseQueue(self._filterQueueLine(self._readAllLines()))

        if len(queue_ls) == 0:
            return {
                "count": -2,
                "message": "GFN not in queue, or wait longer before next request."
            }

        queue = queue_ls[-1]
        if queue["state"] != 1:
            return {
                "count": -1,
                "message": 
                    "Should already be in.\n" 
                    f"(time={queue['datetime'].strftime('%Y-%m-%d %H:%M:%S.%f')}, "
                    f"state={queue['state']}, "
                    f"queue={queue['queue']})"
            }
        return {
            "count": queue['queue']-1,
            "message": 
                f"(time={queue['datetime'].strftime('%Y-%m-%d %H:%M:%S.%f')}, "
                f"state={queue['state']}, "
                f"queue={queue['queue']})"
        }
        
gfnviewer = GFNViewerDebugFile()

@app.route('/gfnviewer', methods=['GET', 'POST'])
def requestQueue():
    d = gfnviewer.getQueueCount()
    print(d["count"])
    return jsonify(d)

@app.route('/gfnopener', methods=['GET', 'POST'])
def requestOpen():
    if GFNHwndManager().getGFNHwnd() > 0:
        return jsonify({
            "code": 1,
            "message": "GFN is already opened."
        })
    cmd = r'cmd /c ' + GFN_START_COMMAND
    sh = shell.shell(cmd)
    return jsonify({
        "code": 0,
        "message": ""
    })

# @app.route('/gfncloser', methods=['GET', 'POST'])
def requestClose():
    hwnd = GFNHwndManager().getGFNHwnd()
    if hwnd < 0:
        return jsonify({
            "code": 1,
            "message": "GFN is already closed."
        })
    
    # actually send mouse event for it to close
    # since WM_CLOSE won't close it and forced thread termination breaks it
    # and I can't make click event happen what the hell
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    time.sleep(1)
    gfnviewer.click(864, 508)

    return jsonify({
        "code": 0,
        "message": ""
    })

@app.route('/tvopener', methods=['GET', 'POST'])
def requestTVOpen():
    cmd = r'cmd /c "' + TEAMVIEWER_PATH + '"'
    sh = shell.shell(cmd)
    return jsonify({
        "code": 0,
        "message": ""
    })

def toggleVPN(do_open):
    # toggle VPN on and off.
    # Uses LogMeIn Hanachi

    print(f"[toggleVPN] {'opening' if do_open else 'closing'}...")
    open_vpn = r'cmd /c "' + HAMACHI_PATH + '" --cli go-online ' + VPN_NAME 
    close_vpn = r'cmd /c "' + HAMACHI_PATH + '" --cli go-offline ' + VPN_NAME 
    sh = shell.shell(
        open_vpn if do_open else close_vpn
    )
    print(f"[toggleVPN] output: {sh.output()}")
    print(f"[toggleVPN] errors: {sh.errors()}")
    

if __name__ == '__main__':
    if TOGGLE_VPN:
        toggleVPN(do_open=True)
        atexit.register(lambda: toggleVPN(do_open=False))
    app.run(host=VPN_SERVER_IP, port=VPN_SERVER_PORT)