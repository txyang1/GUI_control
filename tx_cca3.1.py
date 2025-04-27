import re
import time
from pywinauto import Application
from pywinauto.timings import wait_until

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # ← 改成你的实际路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)

    # 2. 定位左侧 Sessions 标签页用的 TabControl
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP Pane 及其控件
    pcap_pane      = dlg.child_window(auto_id="PCAP",    control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton",     control_type="Button"
    ).wrapper_object()

    # 3.1 确保 PCAP 是勾选状态
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 4. 切到 Sessions 页签，拿到 Table 里的所有 Custom 控件
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 4.1 过滤出真正的行（名称形如 "Row 0", "Row 1", …）
    all_rows = sessions_table.children(control_type="Custom")
    rows = [r for r in all_rows if re.match(r"^Row \d+$", r.element_info.name)]
    print(f"共找到 {len(rows)} 条 Session： {[r.element_info.name for r in rows]}")

    # 5. 对每一行依次：选中，运行 PCAP，等待完成
    for idx, row in enumerate(rows, start=1):
        name = row.element_info.name
        print(f"[{idx}/{len(rows)}] 选中并运行 PCAP → {name}")

        # 5.1 选中这一行
        row.click_input()
        time.sleep(0.2)

        # 5.2 点击 ▶ 按钮启动
        pcap_start_btn.click_input()

        # 5.3 等待按钮先变灰（作业启动），再恢复可用（作业完成）
        wait_until(5,  1, lambda: not pcap_start_btn.is_enabled(),
                   retry_exception=RuntimeError)
        wait_until(600,1, lambda:     pcap_start_btn.is_enabled(),
                   retry_exception=RuntimeError)

        print(f"  ✓ {name} 完成")

    print("🎉 所有 Session 都已跑完！")

if __name__ == "__main__":
    main()
