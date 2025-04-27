import time
from pywinauto import Application, findwindows

# 1. 先连接进程
app = Application(backend="uia").connect(process=17316)

# 2. 列出所有顶层窗口，确认标题
print("=== Top-level windows ===")
for w in app.windows():
    print(" >", repr(w.window_text()))

# 3. 如果看到你想要的窗口名称，比如 "ViGEM CCA-Converter v1.2"，就用它
#    这里暂用模糊匹配
dlg = app.window(title_re=".*CCA-Converter.*")

# 4. 等待窗口就绪
dlg.wait("exists enabled visible ready", timeout=15)

# 5. 打印控件树
dlg.print_control_identifiers()
