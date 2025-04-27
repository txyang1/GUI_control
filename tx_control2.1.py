from pywinauto import Application, timings
import time

# 1. 连接到 CCA-Converter 应用（确保它已启动）
app = Application(backend="uia").connect(path="ViGEMCCA-Converter.exe")
dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
dlg.wait("visible enabled ready", timeout=15)

# 2. 切换到 “Sessions” 页签
sessions_tab = dlg.child_window(title="Sessions", control_type="TabItem").wrapper_object()
sessions_tab.select()
time.sleep(1)  # 等待页面切换完成

# 3. 定位 Sessions 下的会话“表格”，并获取所有行
#    跳过 Timeline 页面的 table（auto_id="timeLineView"），取下一个 Table
all_tables = dlg.descendants(control_type="Table")
session_tables = [tbl for tbl in all_tables if tbl.automation_id != "timeLineView"]
session_table = session_tables[0].wrapper_object()
session_rows  = session_table.children(control_type="DataItem")

# 4. 定位 PCAP 分组里的复选框 和 “Start Jobs” 按钮
pcap_group     = dlg.child_window(auto_id="PCAP", control_type="Pane").wrapper_object()
pcap_checkbox  = pcap_group.child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
start_button   = pcap_group.child_window(auto_id="startButton",    control_type="Button").wrapper_object()

# 5. 依次处理每个会话文件夹
for row in session_rows:
    # 5.1 选中这一行（文件夹）
    row.click_input()

    # 5.2 勾选 PCAP（如果尚未勾选）
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.toggle()

    # 5.3 点击 “Start Jobs”（绿色三角）
    start_button.click_input()

    # 5.4 等待运行开始（按钮不可用），再等运行结束（按钮可用）
    timings.wait_until_passes(10,  0.5, lambda: not start_button.is_enabled())
    timings.wait_until_passes(600, 1.0, start_button.is_enabled)

    # 5.5 取消勾选 PCAP（清除蓝色对号）
    if pcap_checkbox.get_toggle_state() == 1:
        pcap_checkbox.toggle()

    time.sleep(0.5)  # 稍作缓冲，再处理下一个

print("所有 Sessions 文件夹已依次用 PCAP 运行并清除标记。")
