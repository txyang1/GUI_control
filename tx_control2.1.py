from pywinauto import Application, timings
import time

app = Application(backend="uia").connect(path="ViGEMCCA-Converter.exe")
dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
dlg.wait("visible enabled ready", timeout=15)

# 切换到 Sessions
dlg.child_window(title="Sessions", control_type="TabItem").select()
time.sleep(1)

# 方法 A：descendants() 已经是 wrapper 了
all_tables     = dlg.descendants(control_type="Table")
session_tables = [tbl for tbl in all_tables if tbl.automation_id != "timeLineView"]
session_table  = session_tables[0]
session_rows   = session_table.children(control_type="DataItem")

# 定位 PCAP 区域
pcap_group    = dlg.child_window(auto_id="PCAP", control_type="Pane")
pcap_checkbox = pcap_group.child_window(auto_id="activateCheckBox", control_type="CheckBox")
start_button  = pcap_group.child_window(auto_id="startButton",    control_type="Button")

for row in session_rows:
    row.click_input()
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.toggle()
    start_button.click_input()
    timings.wait_until_passes(10,  0.5, lambda: not start_button.is_enabled())
    timings.wait_until_passes(600, 1.0, start_button.is_enabled)
    if pcap_checkbox.get_toggle_state() == 1:
        pcap_checkbox.toggle()
    time.sleep(0.5)

print("全部完成。")

