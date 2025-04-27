import time
from pywinauto import Application

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # 改成你自己的路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()  # 把窗口带到前台

    # 2. 定位左侧 Sessions Tab 和右侧 PCAP 控件
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    ).wrapper_object()

    # 2.1 确保 PCAP 已经打勾
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 3. 切到 Sessions 页签，拿到表格
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 4. 在表格里找所有 “Folder Name Row X” 单元格
    folder_cells = sessions_table.descendants(
        control_type="DataItem",
        title_re="Folder Name Row \\d+"
    )
    print(f"共找到 {len(folder_cells)} 条记录：",
          [c.element_info.name for c in folder_cells])

    # 5. 依次处理
    for idx, cell in enumerate(folder_cells, start=1):
        name = cell.element_info.name
        print(f"[{idx}/{len(folder_cells)}] 选中 → {name}")

        # 5.1 点击这一行的 “Folder Name” 单元格
        cell.click_input()
        time.sleep(0.2)

        # 5.2 点击右侧 ▶ 按钮启动 PCAP
        pcap_start_btn.click_input()

        # 5.3 等待按钮先变灰（启动），再重新可用（完成）
        pcap_start_btn.wait_not("enabled", timeout=5)
        pcap_start_btn.wait("enabled",     timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 都已跑完！")

if __name__ == "__main__":
    main()
