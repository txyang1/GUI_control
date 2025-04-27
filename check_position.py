from pywinauto import Application

app = Application(backend="uia").connect(path="YourApp.exe")
dlg = app.window(title_re=".*Your App Window Title.*")

# 打印出这个窗口下所有子控件的层级和属性
dlg.print_control_identifiers()

