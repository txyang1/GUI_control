import time
from pywinauto import Application

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # 改成你自己的路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)

    # 2. 定位左侧的 Sessions 页签用来切页
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP 区域里的复选框和 ▶ 按钮
    pcap_pane      = dlg.child_window(auto_id="PCAP",    control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton",     control_type="Button"
    ).wrapper_object()

    # 3.1 确保 PCAP 已勾选
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 4. 切到 Sessions
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    # 5. 定位那个 Table（SessionView），并找出它第一列的所有 DataItem
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 找名字里含 "Row" 的 DataItem，就是每行最开始那一格
    row_cells = sessions_table.descendants(
        control_type="DataItem", title_re=r"\bRow \d+\b"
    )
    print(f"共找到 {len(row_cells)} 个 Session：",
          [c.element_info.name.strip() for c in row_cells])

    # 6. 依次点击每个 Session 的第一格，然后 ▶ 运行 PCAP
    for idx, cell in enumerate(row_cells, start=1):
        name = cell.element_info.name.strip()
        print(f"[{idx}/{len(row_cells)}] 处理 {name}")

        # 6.1 选中这一行：点它在 Table 里的第一个 DataItem
        cell.scroll_into_view()
        cell.click_input()
        time.sleep(0.2)

        # 6.2 点击右侧的 ▶ 按钮启动 PCAP
        pcap_start_btn.click_input()

        # 6.3 等待按钮先变灰（启动），再恢复可用（完成）
        pcap_start_btn.wait_not("enabled", timeout=5)
        pcap_start_btn.wait("enabled",     timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 所有 Session 均已处理完毕。")

if __name__ == "__main__":
    main()
