from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

# 1. 连接到已经打开的应用
app = Application(backend="uia").connect(title_re="ViGEM CCA-Converter.*")
dlg = app.window(title_re="ViGEM CCA-Converter.*")

# 2. 切换到 Sessions 选项卡
tab = dlg.child_window(auto_id="tabControl1", control_type="Tab").wrapper_object()
tab.select("Sessions")

# 3. 找到 Sessions 表格
sessions_table = dlg.child_window(auto_id="timeLineView", control_type="Table").wrapper_object()

# 4. 遍历每一行
for row in sessions_table.children(control_type="Custom"):
    # 4.1 选中这一行
    try:
        row.select()  # 有些 DataGridViewRow 可能支持 select()
    except:
        # 如果 select() 不可用，退而使用 click_input()
        row.click_input()
    # 4.2 勾选 PCAP
    pcap_cb = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
    if pcap_cb.get_toggle_state() == 0:  # 0 表示未勾选
        pcap_cb.toggle()
    # 4.3 点击 Start
    start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
    start_btn.click_input()
    # 4.4 等待任务完成——这里假设 statusLabel 显示 “Completed” 表示完成
    def is_done():
        status = dlg.child_window(auto_id="statusLabel", control_type="Text").window_text()
        return "Completed" in status or "完成" in status

    # 最多等 120 秒，每隔 1 秒检查一次
    try:
        timings.wait_until(timeout=120, retry_interval=1, func=is_done)
        print(f"任务完成：{row.window_text()}")
    except timings.TimeoutError:
        print(f"等待超时：{row.window_text()} 未在 120s 内完成")
