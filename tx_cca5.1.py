import time
import re
from pywinauto import Application

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # 改成你自己的路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 2. 定位左侧的 Sessions Tab 和右侧 PCAP Pane
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton",     control_type="Button"
    ).wrapper_object()

    # 2.1 确保 PCAP 已打勾
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 3. 切到 Sessions，定位表格
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 4. 全部 DataItem，然后用 re 过滤 “Folder Name Row X”
    all_dataitems = sessions_table.descendants(control_type="DataItem")
    pattern       = re.compile(r"^Folder Name Row \d+$")
    folder_cells  = [
        item for item in all_dataitems
        if item.element_info.name and pattern.match(item.element_info.name)
    ]

    print(f"共找到 {len(folder_cells)} 条记录：")
    for cell in folder_cells:
        print("  ", cell.element_info.name)

    # 5. 依次处理每一行
    for idx, cell in enumerate(folder_cells, start=1):
        name = cell.element_info.name
        print(f"[{idx}/{len(folder_cells)}] 选中 → {name}")

        # 5.1 点击 “Folder Name” 单元格，选中整行
        cell.click_input()
        time.sleep(0.2)

        # 5.2 点击右侧的 ▶ 按钮启动 PCAP
        pcap_start_btn.click_input()

        # 5.3 等待按钮先变灰（启动），再恢复可用（完成）
        pcap_start_btn.wait_not("enabled", timeout=5)
        pcap_start_btn.wait("enabled", timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 都已跑完！")

if __name__ == "__main__":
    main()
