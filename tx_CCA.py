import time
from pywinauto import Application

# 1. 启动并连接
app = Application(backend="uia").start(r"C:\Path\To\ViGEM CCA-Converter.exe")
dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
dlg.wait("visible enabled ready", timeout=20)

# 2. 定位 TabControl 和 Start Jobs 按钮
tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
start_btn = dlg.child_window(title="Start Jobs", control_type="Button")

# 3. 切到 Sessions 先打印一下所有行，确认定位
tab.child_window(title="Sessions", control_type="TabItem").select()
time.sleep(0.5)
sessions_pane = dlg.child_window(auto_id="sessionsPage", control_type="Pane")
rows = sessions_pane.children(control_type="Custom")  # 每个 Custom 就是一行

print(f"找到 {len(rows)} 行可处理")

# 4. 循环处理每一行
for idx, row in enumerate(rows):
    print(f"=== 处理第 {idx} 行: {row.element_info.name!r} ===")

    # 4.1 选中当前行
    row.click_input()
    time.sleep(0.2)

    # 4.2 切到 Files，勾选 PCAP
    tab.child_window(title="Files", control_type="TabItem").select()
    time.sleep(0.5)
    dlg.child_window(title="PCAP", control_type="CheckBox").click_input()
    time.sleep(0.2)

    # 4.3 点击 Start Jobs 并等待它跑完
    start_btn.click_input()
    # 等待按钮在跑的时候变为不可用，再等它重新可用
    start_btn.wait_not("enabled", timeout=5)    # 确认跑起来了
    start_btn.wait("enabled", timeout=600)      # 最长等 10 分钟，跑完后它会重新启用

    print(f"第 {idx} 行处理完成。")

    # 4.4 切回 Sessions，继续下一个
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

print("所有行已全部运行完毕。")
