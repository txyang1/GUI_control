from pywinauto import Application, timings
import time

# 1. 连接到程序
app = Application(backend="uia").connect(path="ViGEMCCA-Converter.exe")
dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
dlg.wait("visible enabled ready", timeout=15)

# 2. 切换到 Sessions 选项卡
sessions_tab = dlg.child_window(title="Sessions", control_type="TabItem")
sessions_tab.select()

# 3. 定位 Sessions 下的表格（Table）并获取所有行（DataItem）
#    如果有多个 Table，请用 auto_id 或 found_index 定位
session_table = dlg.child_window(control_type="Table", found_index=0).wrapper_object()
rows = session_table.children(control_type="DataItem")

# 4. 定位 PCAP 复选框 和 Run 按钮（通常在 Output Folder 区域）
pcap_checkbox = dlg.child_window(title="PCAP", control_type="CheckBox")
run_button  = dlg.child_window(auto_id="startButton", control_type="Button")

# 5. 依次遍历每一行，选中、勾选 PCAP、运行并等待完成
for row in rows:
    # 5.1 选中当前行
    row.click_input()
    
    # 5.2 勾选 PCAP（如果尚未勾选）
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.toggle()
    
    # 5.3 点击 Run
    run_button.click_input()
    
    # 5.4 等待 Run 按钮变不可用（已开始），再等它变回可用（已结束）
    timings.wait_until_passes(
        5, 0.5,
        lambda: not run_button.is_enabled()
    )
    timings.wait_until_passes(
        300, 1,
        run_button.is_enabled
    )
    
    # 5.5 稍作缓冲
    time.sleep(1)

print("Sessions 里的所有文件已用 PCAP 跑完。")
