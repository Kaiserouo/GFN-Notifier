from flask import Flask, jsonify
import shell
import atexit
import win32gui, win32con
import time

from gfnviewer import *

# --- config ---

# - VPN settings -

# if want to toggle VPN on start, set to true and make sure 
# VPN_NAME and HAMACHI_PATH is correct (only supports LogMeIn Hamachi)
# else just make sure your IP & port is correct
TOGGLE_VPN = True   

# web server IP & port. Set to this computer's VPN's IP
VPN_SERVER_IP = '25.8.47.6'
VPN_SERVER_PORT = 12000

# VPN network name for hamachi
VPN_NAME = 'Kaiserouo'

# Hamachi path for CLI
HAMACHI_PATH = r'C:\Program Files (x86)\LogMeIn Hamachi\x64\hamachi-2.exe'

# - GFN settings -

# the command to execute when there are request for `/gfnopener`
# refer to the game link made by GFN to know what command they use to open your game.
GFN_START_COMMAND = r'"C:\Users\user\AppData\Local\NVIDIA Corporation\GeForceNOW\CEF\GeForceNOWStreamer.exe"  --url-route="#?cmsId=100901811&launchSource=External&shortName=apex_legends_steam_ww&parentGameId=cb2b1b5f-54ba-45fd-9839-96bbfe1376cd"'

# TeamViewer path for request for `/tvopener`
TEAMVIEWER_PATH = r'C:\Program Files\TeamViewer\TeamViewer.exe'

# -------------


app = Flask(__name__)

# change this to GFNViewerTesseract for old detection method
gfnviewer_pc = GFNViewerDebugFile()
gfnviewer_chrome = GFNViewerChromeTesseract()

@app.route('/gfnviewer', methods=['GET', 'POST'])
def requestQueue():
    d_pc = gfnviewer_pc.getQueueCount()
    print(d_pc["count"])
    return jsonify(d_pc)

@app.route('/gfnviewerc', methods=['GET', 'POST'])
def requestQueueChrome():
    d_ch = gfnviewer_chrome.getQueueCount()
    return jsonify(d_ch)

@app.route('/gfnviewerb', methods=['GET', 'POST'])
def requestQueueBoth():
    d_pc = gfnviewer_pc.getQueueCount()
    d_ch = gfnviewer_chrome.getQueueCount()
    print(f'PC={d_pc["count"]} | Chrome={d_ch["count"]}')
    return jsonify({
        "count": min(d_pc["count"], d_ch["count"]),
        "message": f'PC={d_pc["count"]}, Chrome={d_ch["count"]}\n'
            f'----- PC -----\n{d_pc["message"]}\n--------------\n'
            f'--- Chrome ---\n{d_ch["message"]}\n--------------'
    })

@app.route('/gfnopener', methods=['GET', 'POST'])
def requestOpen():
    if GFNHwndManager().getGFNDesktopHwnd() > 0:
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
def requestGFNClose():
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
    gfnviewer_pc.click(864, 508)

    return jsonify({
        "code": 0,
        "message": ""
    })

@app.route('/tvcloser', methods=['GET', 'POST'])
def requestTVClose():
    hwnd = GFNHwndManager().getHwndUnique(lambda h, s: s == 'TeamViewer')
    if hwnd < 0:
        return jsonify({
            "code": 1,
            "message": "TeamViewer is already closed."
        })
    
    # actually send mouse event for it to close
    # since WM_CLOSE won't close it and forced thread termination breaks it
    # and I can't make click event happen what the hell
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

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