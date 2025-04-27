from pywinauto import Application, Desktop, timings

# 1. Attach 到主窗口
main_win = Desktop(backend="uia").window(title="ViGEM CCA-Converter (3.1.4.0)", control_type="Window")
app = Application(backend="uia").connect(handle=main_win.handle)
dlg = app.window(handle=main_win.handle)
dlg.wait("visible ready", timeout=15)

# 2. 调试：看看到底搜到几个 TabControl
tabs = dlg.descendants(control_type="Tab")
print(f"找到 {len(tabs)} 个 TabControl：")
for idx, t in enumerate(tabs):
    print(f"  #{idx}", "name=", t.element_info.name, "auto_id=", t.element_info.auto_id)

# 3. 假设第一个就是我们要的（如果不是，根据上面输出换成正确的索引）
tab = tabs[0].wrapper_object()
tab.select("Sessions")  # 切换到 Sessions

# 4. 定位 Sessions 面板下的表格
sess_pane = dlg.child_window(auto_id="timeLinePage", control_type="Pane")
sess_pane.wait("exists ready visible", timeout=10)
table = sess_pane.child_window(auto_id="timeLineView", control_type="Table").wrapper_object()

# 5. 遍历每一行，打钩 PCAP，点 Start，等“完成”
for row in table.children(control_type="Custom"):
    # 5.1 选中这一行
    try:
        row.select()
    except:
        row.click_input()
    # 5.2 确保 PCAP 被勾上
    pcap = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
    if pcap.get_toggle_state() == 0:
        pcap.toggle()
    # 5.3 点击 Start
    start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
    start_btn.click_input()
    # 5.4 等待状态变“完成”
    def is_done():
        txt = dlg.child_window(auto_id="statusLabel", control_type="Text").window_text()
        return "完成" in txt or "Completed" in txt

    try:
        timings.wait_until(timeout=60, retry_interval=1, func=is_done)
        print("✔ 这一行处理完成")
    except timings.TimeoutError:
        print("✘ 这一行超时未完成")

print("所有行都跑完了！")
