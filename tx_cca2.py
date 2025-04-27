import time
from pywinauto import Application

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # 改成你自己的路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)

    # 2. 定位左侧的 TabControl（只用它来切 Sessions）
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP Pane 里的控件
    pcap_pane      = dlg.child_window(auto_id="PCAP",    control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton",     control_type="Button"
    ).wrapper_object()

    # 3.1 确保 PCAP 被选中
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 4. 切到 Sessions，枚举所有行
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    rows = sessions_table.children(control_type="Custom")
    print(f"共找到 {len(rows)} 条 Session：",
          [r.element_info.name for r in rows])

    # 5. 依次处理每一行
    for idx, row in enumerate(rows, start=1):
        name = row.element_info.name
        print(f"[{idx}/{len(rows)}] 选中并运行 PCAP → {name}")

        # 5.1 选中这一行
        row.click_input()
        time.sleep(0.2)

        # 5.2 点击右侧的 ▶ 按钮启动 PCAP
        pcap_start_btn.click_input()

        # 5.3 等待按钮先变灰（启动），再恢复可用（完成）
        pcap_start_btn.wait_not("enabled", timeout=5)
        pcap_start_btn.wait("enabled",     timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 都已跑完！")

if __name__ == "__main__":
    main()
