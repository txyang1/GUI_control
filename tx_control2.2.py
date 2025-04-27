from pywinauto import Application, timings
import time

app = Application(backend="uia").connect(path="ViGEMCCA-Converter.exe")
dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
dlg.wait("visible ready")

# 定位到 Input Location 里的 TabControl
tab_ctrl = dlg.child_window(auto_id="tabControl1", control_type="Tab")
# 切到 Sessions
tab_ctrl.child_window(title="Sessions", control_type="TabItem").select()
time.sleep(0.5)  # 等 UI 刷新

# **关键**：取第二个 Pane（found_index=1），它就是 Sessions 页的内容区
sessions_page = tab_ctrl.child_window(control_type="Pane", found_index=1).wrapper_object()
sessions_page.wait("visible ready")

# 假设列表是一个 Table
session_table = sessions_page.child_window(control_type="Table").wrapper_object()
rows = session_table.children(control_type="Custom")

# 复用之前拿到的 PCAP 复选框 & Run 按钮
pcap_chk = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
run_btn  = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()

for row in rows:
    row.click_input()
    pcap_chk.check()
    run_btn.click_input()

    # 等待这一行里出现“✔”
    timings.wait_until(
        timeout=300, retry_interval=1,
        func=lambda: any(c.window_text() == '✔'
                         for c in row.descendants(control_type="DataItem"))
    )
    pcap_chk.uncheck()
