import time, re
from pywinauto import Application, timings

def main():
    app = Application(backend="uia").start(r"C:\Path\To\ViGEM CCA-Converter.exe")
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 定位 Sessions 页面和 PCAP 启动按钮
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    pcap_start_btn = pcap_pane.child_window(auto_id="startButton", control_type="Button")

    # 确保 PCAP 已勾
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 切到 Sessions，找左侧“Row X”
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()

    all_dataitems = sessions_table.descendants(control_type="DataItem")
    select_pattern = re.compile(r"^\s*Row \d+$")
    select_cells = [it for it in all_dataitems if it.element_info.name and select_pattern.match(it.element_info.name)]
    print(f"共找到 {len(select_cells)} 条 Session：", [c.element_info.name.strip() for c in select_cells])

    # 循环处理
    for idx, cell in enumerate(select_cells, start=1):
        name = cell.element_info.name.strip()
        print(f"[{idx}/{len(select_cells)}] 选中 → {name}")

        # 1. 点击行前面的 DataItem
        cell.click_input()
        time.sleep(0.2)

        # 2. 点击 ▶
        pcap_start_btn.click_input()

        # 3. 尝试等禁用
        try:
            pcap_start_btn.wait_not("enabled", timeout=5)
            print("    → 按钮已禁用，确认启动")
        except timings.TimeoutError:
            print("    → 按钮未禁用，跳过禁用等待")

        # 4. 等待重新可用
        pcap_start_btn.wait("enabled", timeout=600)
        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 都已跑完！")

if __name__ == "__main__":
    main()
