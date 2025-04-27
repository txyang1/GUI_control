from pywinauto import Application, timings
import os
import time

# 1. 启动或连接到你的 APP
# 如果 APP 已经在运行，可以用 .connect；否则用 .start 启动
app = Application(backend="uia").start(r"C:\Path\To\YourApp.exe")
# 或者：app = Application(backend="uia").connect(path="YourApp.exe")

# 2. 获取主窗口（根据标题或 class_name）
dlg = app.window(title_re=".*Your App Window Title.*")

# 等待窗口就绪
dlg.wait("visible enabled ready", timeout=20)

# 3. 在“路径”输入框中输入你想要的目录
# 假设该输入框的 AutomationId 是 "PathEdit"，你可以用 Inspect.exe 等工具查看
path_edit = dlg.child_window(auto_id="PathEdit", control_type="Edit")
path_edit.set_edit_text(r"C:\Users\you\output_folder")

# 4. 点击路径下某个文件（如果是列表/ListView）
# 假设是一个 ListView，AutomationId="FileList"
file_list = dlg.child_window(auto_id="FileList", control_type="List")
# 选择名字为 "data.csv" 的那一行
file_list.get_item("data.csv").select()

# 5. 点击“运行”按钮
# 假设按钮的文字是 Run，或 AutomationId="RunButton"
run_btn = dlg.child_window(title="Run", control_type="Button")
run_btn.click_input()

# 6. 等待程序执行完毕并生成输出
# 可以通过轮询输出文件是否出现，或检测某个控件的状态
output_file = r"C:\Users\you\output_folder\result.txt"
timeout = 60
t0 = time.time()
while not os.path.exists(output_file):
    if time.time() - t0 > timeout:
        raise RuntimeError("等待输出文件超时")
    time.sleep(1)

# 7. 重命名生成的文件
new_name = r"C:\Users\you\output_folder\result_renamed.txt"
os.replace(output_file, new_name)

print("自动化流程完成，文件已重命名为：", new_name)
