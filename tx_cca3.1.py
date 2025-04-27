import time
from pywinauto import Application

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # 改成你的可执行文件路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)

    # 2. 定位左侧 TabControl（只用它来切 Sessions）
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP Pane 里的控件（保持为 WindowSpecification）
    pcap_pane      = dlg.child_window(auto_id="PCAP",    control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    )
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton",     control_type="Button"
    )

    # 3.1 确保 PCAP 已勾选
    if pcap_checkbox.wrapper_object().get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 4. 切到 Sessions，拿到所有行（过滤掉表头“Top Row”）
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    )
    all_rows = sessions_table.children(control_type="Custom")
    # 只留下名字是 “Row X” 的那些
    rows = [r for r in all_rows if r.element_info.name.startswith("Row ")]
    print(f"共找到 {len(rows)} 条 Session：", [r.element_info.name for r in rows])

    # 5. 依次处理每一行
    for idx, row_spec in enumerate(rows, start=1):
        name = row_spec.element_info.name
        print(f"[{idx}/{len(rows)}] 选中并运行 PCAP → {name}")

        # 5.1 选中这一行
        row_spec.click_input()
        time.sleep(0.2)

        # 5.2 点击 ▶ 启动 PCAP
        pcap_start_btn.click_input()

        # 5.3 等待按钮先变灰（作业启动），再恢复可用（作业完成）
        pcap_start_btn.wait_not("enabled", timeout=5)
        pcap_start_btn.wait("enabled",     timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 都已运行完毕！")

if __name__ == "__main__":
    main()
