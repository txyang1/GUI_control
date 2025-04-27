from pywinauto import Application, Desktop, timings

# 1. 先用 Desktop 找到主窗口
wins = Desktop(backend="uia").windows(title_re=".*ViGEM CCA-Converter.*")
if not wins:
    raise RuntimeError("找不到 ViGEM CCA-Converter 窗口，请确认程序已启动且标题正确")
main = wins[0]

# 2. 用 handle 去 connect
app = Application(backend="uia").connect(handle=main.handle)
dlg = app.window(handle=main.handle)
dlg.wait("visible ready", timeout=15)

# 3. 定位到那个深度 Pane（它包含了 tabControl1）
tab_parent = dlg.child_window(auto_id="1575180", control_type="Pane")
tab_parent.wait("exists ready visible", timeout=10)

# 4. 拿到 TabControl，并切换到 “Sessions”
tab = tab_parent.child_window(auto_id="tabControl1", control_type="Tab").wrapper_object()
tab.select("Sessions")

# 5. 定位到 Sessions 页面的 DataGrid（timeLineView）
sess_page = tab_parent.child_window(auto_id="timeLinePage", control_type="Pane")
sess_page.wait("exists ready visible", timeout=10)

table = sess_page.child_window(auto_id="timeLineView", control_type="Table").wrapper_object()

# 6. 遍历每一行，选中→勾 PCAP →点击 Start →等待完成
for idx, row in enumerate(table.children(control_type="Custom"), start=1):
    print(f"处理第 {idx} 行…")
    # 6.1 选中
    try:
        row.select()
    except:
        row.click_input()
    # 6.2 勾 PCAP（如果没勾就勾上）
    pcap_cb = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
    if pcap_cb.get_toggle_state() == 0:
        pcap_cb.toggle()
    # 6.3 点击 Start
    start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
    start_btn.click_input()
    # 6.4 等待状态变成“完成”或“Completed”
    def done():
        txt = dlg.child_window(auto_id="statusLabel", control_type="Text").window_text()
        return "完成" in txt or "Completed" in txt

    try:
        timings.wait_until(timeout=120, retry_interval=1, func=done)
        print(f"✔ 第 {idx} 行 已完成")
    except timings.TimeoutError:
        print(f"✘ 第 {idx} 行 超时")

print("全都处理完了！")
