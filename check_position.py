from pywinauto import Desktop, Application
import time

# 先启动或连接 APP
app = Application(backend="uia").connect(path="ViGEMCCA-Converter.exe")
pid = app.process

# 等个几秒，保证窗口已经弹出来
time.sleep(2)

# 枚举这个进程的所有顶层窗口，打印它们的标题
for w in Desktop(backend="uia").windows(process=pid):
    print(repr(w.window_text()))
