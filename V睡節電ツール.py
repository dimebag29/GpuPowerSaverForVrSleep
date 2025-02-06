# ==============================================================================================================
# 作成者:dimebag29 作成日:2025年2月6日 バージョン:v0.4
# (Author:dimebag29 Creation date:February 6, 2025 Version:v0.4)
#
# このプログラムのライセンスはLGPLv3です。pynputライブラリのライセンスを継承しています。
# (This program is licensed to LGPLv3. Inherits the license of the pynput library.)
# https://www.gnu.org/licenses/lgpl-3.0.html.en
#
# 開発環境 (Development environment)
# ･python 3.7.5
# ･auto-py-to-exe 2.36.0 (used to create the exe file)
#
# exe化時のauto-py-to-exeの設定
# ･ひとつのディレクトリにまとめる (--onedir)
# ･ウィンドウベース (--windowed)
# ･exeアイコン設定 (--icon) (GpuPowerSaveExeIcon.ico)
# ･追加ファイルで電力制限ONとOFFのタスクトレイアイコン追加 (--add-data) (GpuPowerSaveOnIcon.ico, GpuPowerSaveOffIcon.ico)
# ･高度な設定で管理者権限実行に設定 (--uac-admin)
# ==============================================================================================================

# python 3.7.5の標準ライブラリ (Libraries included as standard in python 3.7.5)
import subprocess
import os
import sys
import time
import math
import threading
import socket
import json

# 外部ライブラリ (External libraries)
import win32gui                                                                 # Included in pywin32 Version:306
from PIL import Image                                                           # Version:9.5.0
from pystray import Icon, Menu, MenuItem                                        # Version:0.19.5
from pynput import keyboard                                                     # Version:1.7.6
import psutil                                                                   # Version:5.9.6


# ================================================= 関数定義 ====================================================
# タスクトレイからの操作関係 -----------------------------------------------------
# 終了ボタンが押された時の動作
def Push_Exit():
    global Icon                                                                 # 関数内で変数の値を変更したい場合はglobalにする
    Icon.icon = GpuPowerSaveOffIcon                                             # タスクトレイアイコン変更
    cmd = MasterBatPath + " " + str(DefaultPowerLimit)                          # コマンド生成。batの引数に電力制限初期値を指定
    subprocess.run(cmd, startupinfo=StartupInfo)                                # コマンド実行
    time.sleep(3.5)                                                             # 少し待ってからプログラムを停止させる
    Icon.stop()                                                                 # タスクトレイ常駐終了

# 電力制限ONボタンが押された時の動作
def Push_PowerSaveOn():
    global Icon                                                                 # 関数内で変数の値を変更したい場合はglobalにする
    Icon.icon = GpuPowerSaveOnIcon                                              # タスクトレイアイコン変更
    cmd = MasterBatPath + " " + str(CustomPowerSaveValue)                       # コマンド生成。batの引数に電力制限下限を指定
    subprocess.run(cmd, startupinfo=StartupInfo)                                # コマンド実行

# 電力制限OFFボタンが押された時の動作
def Push_PowerSaveOff():
    global Icon                                                                 # 関数内で変数の値を変更したい場合はglobalにする
    Icon.icon = GpuPowerSaveOffIcon                                             # タスクトレイアイコン変更
    cmd = MasterBatPath + " " + str(DefaultPowerLimit)                          # コマンド生成。batの引数に電力制限初期値を指定
    subprocess.run(cmd, startupinfo=StartupInfo)                                # コマンド実行


# XSOverlayからの操作関係 -------------------------------------------------------
# 電力制限切り替え関数
def Push_SW():
    global SW                                                                   # 関数内で変数の値を変更したい場合はglobalにする

    # 電力制限が無効になってたら有効にする処理
    if False == SW:
        SW = True                                                               # 電力制限の状態を"有効"で保存
        # スレッド化して実行 (スレッド化しなしとメディアキー(前のトラック)入力を抑止できなかった)
        PowerSaveOnThread = threading.Thread(target=Push_PowerSaveOn, daemon=True)
        PowerSaveOnThread.start()
    
    # 電力制限が有効になってたら無効にする処理
    else:
        SW = False                                                              # 電力制限の状態を"無効"で保存
        # スレッド化して実行 (スレッド化しなしとメディアキー(前のトラック)入力を抑止できなかった)
        PowerSaveOffThread = threading.Thread(target=Push_PowerSaveOff, daemon=True)
        PowerSaveOffThread.start()


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
            Push_SW()                                                           # 電力制限切り替え関数実行
        listener.suppress_event()                                               # メディアキー(前のトラック)入力が他のプログラムに伝わらないようにここで抑止
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


# その他 -----------------------------------------------------------------------
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
# subprocessでコマンド実行したときにコマンドプロンプトのウインドウが表示されないようにする設定 (https://chichimotsu.hateblo.jp/entry/20140712/1405147421)
StartupInfo = subprocess.STARTUPINFO()
StartupInfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
StartupInfo.wShowWindow = subprocess.SW_HIDE

SW = False                                                                      # 電力制限有効:True、無効:False

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


# ============================================= 多重起動してたら終了 =============================================
MyExeName = os.path.basename(sys.argv[0])                                       # 自分のexe名を取得 (拡張子付き)

ProcessHitCount = 0                                                             # 自分を同じ名前のexeがプロセス内に何個あるかカウントする用
for MyProcess in psutil.process_iter():                                         # プロセス一覧取得
    try:
        if MyExeName == os.path.basename(MyProcess.exe()):                      # 自分を同じ名前のexeだったら
            ProcessHitCount = ProcessHitCount + 1                               # カウントアップ
    except:
        pass

# 単一起動時はexeが2つある(なぜかはわからない)。それを超えていたら多重起動しているということなのでここで終了
if 2 < ProcessHitCount:
    sys.exit(0)


# ============================================== ファイルの生成 =================================================
# AppData\Localの中にGpuPowerSaverというフォルダを生成。すでにあったらそのまま-------
MyLocalPath = os.path.expanduser("~\AppData\Local\GpuPowerSaver")
os.makedirs(MyLocalPath, exist_ok=True)

# PowerSaveCommand.batを管理者権限で実行するbatファイル ---------------------------
MasterBatName = "RunPowerSaveCommand.bat"                                       # ファイル名
MasterBatPath = MyLocalPath + "\\" +  MasterBatName                             # ファイルパス
MasterBat = open(MasterBatPath, 'w')                                            # 上書きモードでファイル生成
# batのコマンド生成
Command = '@rem Powershellを使って管理者権限でPowerSaveCommand.batを実行する。%1はPythonから渡される引数。電力値[W]が入ってくる\n'
Command += 'powershell.exe -Command Start-Process %~dp0\PowerSaveCommand.bat """%1""" -Verb Runas -WindowStyle Hidden\n'
MasterBat.write(Command)                                                        # ファイル書き込み
MasterBat.close()                                                               # ファイルを閉じる

# nvidia-smiを使って電力制限を行うbatファイル -------------------------------------
SlaveBatName  = "PowerSaveCommand.bat"                                          # ファイル名
SlaveBatPath  = MyLocalPath + "\\" +  SlaveBatName                              # ファイルパス
SlaveBat = open(SlaveBatPath, 'w')                                              # 上書きモードでファイル生成
# batのコマンド生成
Command = '@rem 電力制限実行。 %1にはRunPowerSaveCommand.batから渡された電力値[W]が入ってる\n'
Command += 'nvidia-smi -pl %1\n'
SlaveBat.write(Command)                                                         # ファイル書き込み
SlaveBat.close()                                                                # ファイルを閉じる

# カスタム電力制限値が書かれたファイル --------------------------------------------
CustomPowerSaveValuePath = os.path.expanduser("~\AppData\Local\GpuPowerSaver\CustomPowerSaveValue.txt")
if not os.path.isfile(CustomPowerSaveValuePath):
    CustomPowerSaveValueTxt = open(CustomPowerSaveValuePath, 'w')               # 上書きモードでファイル生成
    CustomPowerSaveValueTxt.write("0")                                          # ファイル書き込み(初期値は0にしておく)
    CustomPowerSaveValueTxt.close()                                             # ファイルを閉じる



# ============================================= グラボの仕様を取得 ===============================================
# 電力制限初期値(int)を取得
cmd = "nvidia-smi --query-gpu=power.default_limit --format=csv"
DefaultPowerLimit = math.floor(float(subprocess.run(cmd, startupinfo=StartupInfo, stdout=subprocess.PIPE).stdout.decode().splitlines()[1].split()[0]))
# 電力制限下限(int)を取得
cmd = "nvidia-smi --query-gpu=power.min_limit --format=csv"
MinimumPowerLimit = math.floor(float(subprocess.run(cmd, startupinfo=StartupInfo, stdout=subprocess.PIPE).stdout.decode().splitlines()[1].split()[0]))
# 現在の電力制限値(int)を取得
cmd = "nvidia-smi --query-gpu=power.limit --format=csv"
NowPowerLimit     = math.floor(float(subprocess.run(cmd, startupinfo=StartupInfo, stdout=subprocess.PIPE).stdout.decode().splitlines()[1].split()[0]))
#print(DefaultPowerLimit, MinimumPowerLimit, NowPowerLimit)

# 電力制限初期値よりも現在の電力制限値が小さかったら、電力制限ON状態にしておく
if DefaultPowerLimit > NowPowerLimit:
    SW = True


# ========================================== カスタム電力制限値を取得 ============================================
# カスタム電力制限値(int)を取得。取得できなかったら0にする
with open(CustomPowerSaveValuePath, "r") as f:
    try:
        CustomPowerSaveValue = math.floor(float(f.readlines()[0]))
    except:
        CustomPowerSaveValue = 0

# 電力制限初期値よりもカスタム電力制限値に大きい値が設定されていたら電力制限初期値にする
if DefaultPowerLimit <= CustomPowerSaveValue:
    CustomPowerSaveValue = DefaultPowerLimit
# 電力制限下限よりもカスタム電力制限値に小さい値が設定されていたら電力制限下限にする
if MinimumPowerLimit >= CustomPowerSaveValue:
    CustomPowerSaveValue = MinimumPowerLimit


# =============================================== スレッド開始 ==================================================
# メディアキー入力監視開始 (無限ループさせてる関数なので、daemon=Trueでデーモン化しないとメインスレッドが終了しても生き残り続けてしまう)
MediakeyLoggingThread = threading.Thread(target=StartMediakeyLoggingThread, daemon=True)
MediakeyLoggingThread.start()

# VRChatが実行されているか監視開始 (無限ループさせてる関数なので、daemon=Trueでデーモン化しないとメインスレッドが終了しても生き残り続けてしまう)
VRChatLoggingThread = threading.Thread(target=StartVRChatLoggingThread, daemon=True)
VRChatLoggingThread.start()


# ============================================== タスクトレイ生成 ================================================
# タスクトレイのアイコン設定
IconBasePath = os.path.abspath(".")                                             # exeが置かれている場所取得
GpuPowerSaveOffIcon = Image.open(os.path.join(IconBasePath, "GpuPowerSaveOffIcon.ico")) # 電力制限OFFアイコン
GpuPowerSaveOnIcon  = Image.open(os.path.join(IconBasePath, "GpuPowerSaveOnIcon.ico" )) # 電力制限ONアイコン

# タスクトレイアイコンを右クリックしたときのメニュー設定
Menu = Menu(
    MenuItem("電力制限 ON",  Push_PowerSaveOn ),
    MenuItem("電力制限 OFF", Push_PowerSaveOff),
    MenuItem("終了",         Push_Exit        ))

# タスクトレイに常駐
Icon = Icon(name="IconName", icon=GpuPowerSaveOffIcon, title="V睡節電ツール", menu=Menu)
Icon.run()
