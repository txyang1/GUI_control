import os

import time
import psutil
import pyautogui
import subprocess
import pyperclip
# 定义要执行的exe文件的路径
tfc_exe_path = r"C:\Users\qxz32ql\test\Carmen4_from_windows\bin\TFC.exe"
# 要处理的文件路径
address_file = r"C:\Users\qxz32ql\test\jiar_count\python_qt5"
# 结果路径
address_result = r"C:\Users\qxz32ql\test\jiar_count"
if address_result:
    result = os.path.join(address_result, "result")
else:
    result = os.path.join(address_file, "result")


if not os.path.exists(result):  # 文件夹不存在则创建
    os.makedirs(result)

file_list = os.listdir(address_result)
matching_files = []


def cat_exe():
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if "TFC.exe" in proc.name():
            return True
    else:
        return False


for file in file_list:
    if file.endswith(".MF4"):
        matching_files.append(file)
subprocess.Popen([tfc_exe_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

time.sleep(10)  # 10秒钟启动TFC.exe
while True:
    time.sleep(3)
    values = cat_exe()
    if values:
        break
pyautogui.moveTo(1500, 20)
# 执行点击操作
pyautogui.click()
time.sleep(1)
pyautogui.moveTo(1500, 20)
time.sleep(1)
pyautogui.dragTo(10, 20)
time.sleep(1)
# 如果需要确保鼠标左键释放，可以使用
pyautogui.mouseUp()
pyautogui.moveTo(1550, 45)
time.sleep(1)
pyautogui.click()

for file_name in matching_files:
    pyautogui.moveTo(100, 300, duration=1)  # 添加文件
    time.sleep(1)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(1010, 65, duration=1)  # 设置文件地址
    time.sleep(1)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyperclip.copy(address_result)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    # 模拟回车键
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.moveTo(700, 705, duration=1)  # 搜素文件
    pyautogui.click()
    time.sleep(2)
    pyperclip.copy(file_name)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    # 模拟回车键
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.moveTo(1100, 290, duration=1)
    time.sleep(1)
    pyautogui.click()
    pyperclip.copy(result)
    time.sleep(1)
    pyautogui.moveTo(1010, 65, duration=1)
    time.sleep(1)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    # 模拟回车键
    pyautogui.press('enter')
    pyautogui.moveTo(1200, 750, duration=1)  # 替换为程序窗口的坐标
    time.sleep(1)
    pyautogui.click()
    pyautogui.moveTo(150, 50, duration=1)  # 替换为程序窗口的坐标
    time.sleep(1)
    pyautogui.click()
    time.sleep(5)  # 生成结果的时间
    file_name_1 = file_name.replace(".MF4", ".pcap")
    try:
        os.rename(os.path.join(result, "conversion_output.pcap"), os.path.join(result, file_name_1))
    except Exception as e:
        print(e)
        pass
    pyautogui.moveTo(1900, 1100, duration=1)  # 替换为程序窗口的坐标
    time.sleep(1)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(880, 420, duration=1)  #
    time.sleep(1)
    pyautogui.click()

for proc in psutil.process_iter(attrs=['pid', 'name']):
    if "TFC.exe" in proc.name():
        proc.terminate()