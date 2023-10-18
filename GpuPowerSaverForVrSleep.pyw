# ==============================================================================================================
# 作成者:dimebag29 作成日:2023年10月18日 バージョン:v0.3
# (Author:dimebag29 Creation date:October 18, 2023 Version:v0.3)
#
# このプログラムのライセンスはLGPLv3です。pynputライブラリのライセンスを継承しています。
# (This program is licensed to LGPLv3. Inherits the license of the pynput library.)
# https://www.gnu.org/licenses/lgpl-3.0.html.en
#
# 開発環境 (Development environment)
# ･python 3.7.5
# ･auto-py-to-exe 2.36.0 (used to create the .exe file)
# ==============================================================================================================

# python 3.7.5の標準ライブラリ (Libraries included as standard in python 3.7.5)
from tkinter import *
import subprocess
import os
import time
import math
import threading
import socket
import json

# 外部ライブラリ (External libraries)
import win32gui                                                                 # Version:306 (Included in pywin32)
from pynput import keyboard                                                     # Version:1.7.6



# ================================================= 関数定義 ====================================================
# GUI関係 ----------------------------------------------------------------------
# ボタンが押された時の処理
def PushSW():
    global SW                                                                   # 関数内で変数の値を変更したい場合はglobalにする

    # パワーセーブが無効になってたら有効にする処理
    if False == SW:
        SW = True                                                               # パワーセーブの状態を"有効"で保存
        os.system(MasterBatPath + " " + str(MinimumPowerLimit))                 # 引数に電力制限下限を指定してRunPowerSaveCommand.batを実行
    
    # パワーセーブが有効になってたら無効にする処理
    else:
        SW = False                                                              # パワーセーブの状態を"無効"で保存
        os.system(MasterBatPath + " " + str(DefaultPowerLimit))                 # 引数に電力制限初期値を指定してRunPowerSaveCommand.batを実行
    
    root.after(3500, Update)                                                    # 少し待ってからにUpdate()を実行する


# ウィンドウの閉じるボタンが押された時の処理
def PushClose():
    os.system(MasterBatPath + " " + str(DefaultPowerLimit))                     # 引数に電力制限初期値を指定してRunPowerSaveCommand.batを実行

    # 表示内容更新
    Info  = "元の状態に戻して\n"
    Info += "　 終了します"
    button01["text"] = Info                                                     # ボタンのテキスト更新
    root.update_idletasks()                                                     # GUIの更新はこのPushClose()関数が終わってからなので、今強制的にGUI更新
    time.sleep(3.5)                                                             # 少し待ってからプログラムを停止させる
    listener.stop()
    root.destroy()                                                              # プログラム終了


# 表示内容更新
def Update():
    # 現在の電力制限値(int)を取得
    cmd = "nvidia-smi --query-gpu=power.limit --format=csv"
    NowPowerLimit = math.floor(float(subprocess.check_output(cmd).decode().splitlines()[1].split()[0]))

    # 表示内容更新
    Info = "今の状態：\n"
    Info += "パワーセーブ　："
    if True == SW:  Info += "ON\n"
    else:           Info += "OFF\n"
    Info += "電力制限設定値：" + str(NowPowerLimit) + " W\n"
    Info += "\n"
    Info += "グラボ仕様：\n"
    Info += "電力制限初期値：" + str(DefaultPowerLimit) + " W\n"
    Info += "電力制限下限値：" + str(MinimumPowerLimit) + " W"
    button01["text"] = Info                                                     # ボタンのテキスト更新


# XSOverlayからの操作関係 -------------------------------------------------------
# メディアキー入力監視スレッド
def StartMediakeyLoggingThread():
    global listener                                                             # 関数内で変数の値を変更したい場合はglobalにする

    with keyboard.Listener(on_press=None, on_release=None, win32_event_filter=win32_event_filter, suppress=False) as listener:
        listener.join()
    # https://stackoverflow.com/questions/54394219/pynput-capture-keys-prevent-sending-them-to-other-applications
    # https://github.com/moses-palmer/pynput/issues/170#issuecomment-602743287


# メディアキー(前のトラック)が押された時の処理
def win32_event_filter(msg, data):
    if 0xB1 == data.vkCode and True == VRChatRunning:
        if 256 == msg:
            ViewXSOverlayNotification(not SW)                                   # 現在の電力制限ON/OFF設定値ではなく、切り替え予定の電力制限ON/OFF設定値を送る
            PushSW()                                                            # ウィンドウのボタンが押された時の処理を実行
        listener.suppress_event()                                               # メディアキー(前のトラック)入力が他のプログラムに伝わらないようにここで殺す
    # https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    #   data.vkCode >> 再生/一時停止:0xB3、次のトラック:0xB0、前のトラック:0xB1
    # https://github.com/moses-palmer/pynput/issues/170#issuecomment-602743287
    #   msg >> キー下げ:257、キー上げ:256


# XSOverlayに通知を出す処理
def ViewXSOverlayNotification(NotificationInput):
    global Message                                                              # 関数内で変数の値を変更したい場合はglobalにする

    MySocket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)                 # ソケット通信用インスタンスを生成

    # XSOverlay通知文更新
    if True == NotificationInput:
        Message["content"] = "電力制限 : ON"
    else:
        Message["content"] = "電力制限 : OFF"
    SendData = json.dumps(Message).encode("utf-8")

    MySocket.sendto(SendData, ("127.0.0.1", 42069))                             # XSOverlayに通知依頼を送信
    MySocket.close()                                                            # ソケット通信終了
    # https://zenn.dev/eeharumt/scraps/95f49a62dd809a
    # https://gist.github.com/nekochanfood/fc8017d8247b358154062368d854be9c


# VRChatが実行されているか監視するスレッド
def StartVRChatLoggingThread():
    global VRChatRunning                                                        # 関数内で変数の値を変更したい場合はglobalにする

    # VRChatが起動されているか監視ループ。
    while True:
         # VRChat検出用のウィンドウ名と完全一致するウィンドウを取得してみる。なかった場合は0(int)、あった場合はウィンドウハンドル(int)が返ってくる
        WindowHandle = win32gui.FindWindow(None, VRChatWindowName)
        
        # VRChatを検知できた
        if 0 != WindowHandle:
            VRChatRunning = True
        # VRChatを検知できなかった
        else:
            VRChatRunning = False

        time.sleep(5)                                                           # 少し待つ



# ================================================== 初期化 ====================================================
SW = False                                                                      # パワーセーブ有効:True、無効:False
WindowName = "GpuPowerSaver"                                                    # ウィンドウ名定義

VRChatWindowName = "VRChat"                                                     # VRChat検出用のウィンドウ名定義
VRChatRunning = False                                                           # VRChatが起動していたらTrueになるフラグ。VRChatが起動してなかったらメディアキーによる電力制限の制御をしないようにする

# XSOverlay共通通知文 https://xiexe.github.io/XSOverlayDocumentation/#/NotificationsAPI?id=xsoverlay-message-object
Message = {
    "messageType" : 1,          # 1 = Notification Popup, 2 = MediaPlayer Information, will be extended later on.
    "index" : 0,                # Only used for Media Player, changes the icon on the wrist.
    "timeout" : 2.0,            # How long the notification will stay on screen for in seconds
    "height" : 100.0,           # Height notification will expand to if it has content other than a title. Default is 175
    "opacity" : 1.0,            # Opacity of the notification, to make it less intrusive. Setting to 0 will set to 1.
    "volume" : 0.5,             # Notification sound volume.
    "audioPath" : "default",    # File path to .ogg audio file. Can be "default", "error", or "warning". Notification will be silent if left empty.
    "title" : "V睡節電ツール",   # Notification title, supports Rich Text Formatting
    "useBase64Icon" : False,    # Set to true if using Base64 for the icon image
    "icon" : "default",         # Base64 Encoded image, or file path to image. Can also be "default", "error", or "warning"
    "sourceApp" : "TEST_App"    # Somewhere to put your app name for debugging purposes
    }


# =========================================== ウィンドウ多重起動防止 =============================================
# 定義したウィンドウ名と完全一致するウィンドウを取得してみる。なかった場合は0(int)、あった場合はウィンドウハンドル(int)が返ってくる
WindowHandle = win32gui.FindWindow(None, WindowName)
if 0 != WindowHandle : exit()                                                   # もし既にウィンドウがあったら終了する


# ============================================= batファイルの生成 ===============================================
# AppData\Localの中にGpuPowerSaverというフォルダを生成。すでにあったらそのまま-------
MyLocalPath = os.path.expanduser("~\AppData\Local\GpuPowerSaver")
os.makedirs(MyLocalPath, exist_ok=True)

# PowerSaveCommand.batを管理者権限で実行するbatファイル ---------------------------
MasterBatName = "RunPowerSaveCommand.bat"                                       # ファイル名
MasterBatPath = MyLocalPath + "\\" +  MasterBatName                             # ファイルパス
MasterBat = open(MasterBatPath, 'w')                                            # 上書きモードでファイル生成
# batのコマンド生成
Command  = '@rem なにも表示しないようにする\n'
Command += '@echo off\n\n'
Command += '@rem 最小化状態で実行しなおす\n'
Command += '@if not "%~0"=="%~dp0.\%~nx0" start /min cmd /c,"%~dp0.\%~nx0" %* & goto :eof\n\n'
Command += '@rem Powershellを使って管理者権限でPowerSaveCommand.batを実行する。%1はPythonから渡される引数。電力値[W]が入ってくる\n'
Command += 'powershell.exe -Command Start-Process %~dp0\PowerSaveCommand.bat """%1""" -Verb Runas\n\n'
Command += '@rem これを入れないとなぜかPythonから動かせなかった\n'
Command += 'powershell sleep 0.1'
MasterBat.write(Command)                                                        # ファイル書き込み
MasterBat.close()                                                               # ファイルを閉じる

# nvidia-smiを使って電力制限を行うbatファイル -------------------------------------
SlaveBatName  = "PowerSaveCommand.bat"                                          # ファイル名
SlaveBatPath  = MyLocalPath + "\\" +  SlaveBatName                              # ファイルパス
SlaveBat = open(SlaveBatPath, 'w')                                              # 上書きモードでファイル生成
# batのコマンド生成
Command  = '@rem 最小化状態で実行しなおす\n'
Command += '@if not "%~0"=="%~dp0.\%~nx0" start /min cmd /c,"%~dp0.\%~nx0" %* & goto :eof\n\n'
Command += '@rem 電力制限実行。 %1にはRunPowerSaveCommand.batから渡された電力値[W]が入ってる\n'
Command += 'nvidia-smi -pl %1'
SlaveBat.write(Command)                                                         # ファイル書き込み
SlaveBat.close()                                                                # ファイルを閉じる


# ============================================= グラボの仕様を取得 ===============================================
# 電力制限初期値(int)を取得
cmd = "nvidia-smi --query-gpu=power.default_limit --format=csv"
DefaultPowerLimit = math.floor(float(subprocess.check_output(cmd).decode().splitlines()[1].split()[0]))
# 電力制限下限(int)を取得
cmd = "nvidia-smi --query-gpu=power.min_limit --format=csv"
MinimumPowerLimit = math.floor(float(subprocess.check_output(cmd).decode().splitlines()[1].split()[0]))
# 現在の電力制限値(int)を取得
cmd = "nvidia-smi --query-gpu=power.limit --format=csv"
NowPowerLimit     = math.floor(float(subprocess.check_output(cmd).decode().splitlines()[1].split()[0]))

# 電力制限初期値よりも現在の電力制限値が小さかったら、電力制限ON状態にしておく
if DefaultPowerLimit > NowPowerLimit:
    SW = True


# =============================================== スレッド開始 ==================================================
# メディアキー入力監視開始 (daemon=Trueでデーモン化しないと、メインスレッドが終了しても生き残り続けちゃう)
MediakeyLoggingThread = threading.Thread(target=StartMediakeyLoggingThread, daemon=True)
MediakeyLoggingThread.start()

# VRChatが実行されているか監視開始 (daemon=Trueでデーモン化しないと、メインスレッドが終了しても生き残り続けちゃう)
VRChatLoggingThread = threading.Thread(target=StartVRChatLoggingThread, daemon=True)
VRChatLoggingThread.start()


# ================================================= GUI生成 ====================================================
# ウィンドウ設定
root = Tk()                                                                     # Tkクラスのインスタンス化
root.title(WindowName)                                                          # タイトル設定
root.geometry("300x200")                                                        # 画面サイズ設定
root.resizable(False, False)                                                    # リサイズ不可に設定
root.protocol("WM_DELETE_WINDOW", PushClose)                                    # ウィンドウの閉じるボタンが押されたらPushClose()を実行するようにする

# フレーム設定
frame01 = Frame(root, width=300, height=200)                                    # フレームサイズ設定
frame01.grid(row=0, column=0)                                                   # 左上ぴったりに配置
frame01.propagate(False)                                                        # ボタンに合わせて大きさ変わらないようにする
frame01.configure(background="#333333")                                         # 背景色

#ボタン設定
button01 = Button(frame01, text='Starting', width=240, height=240, font=("Helvetica", 16, "bold"), justify="left", command=PushSW)
button01.pack(padx=5, pady=5)                                                   # 余白
button01.configure(foreground="#CCCCCC")                                        # 文字色
button01.configure(background="#333333")                                        # ボタン通常色
button01.configure(activebackground="#555555")                                  # ボタン押し込み時の色

Update()                                                                        # Update()を実行
root.mainloop()                                                                 # destroy()されるまでここで待機。(GUIが表示され続ける)
