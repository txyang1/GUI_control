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

    # 2. 定位左侧 Sessions 用的 TabControl
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP Pane 里的控件
    pcap_pane      = dlg.child_window(auto_id="PCAP",    control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton",     control_type="Button"
    ).wrapper_object()

    # 3.1 保证 PCAP 复选框已勾选
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 4. 切到 Sessions、拿到表格中的所有 Custom（行）
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    all_rows = sessions_table.children(control_type="Custom")
    # 只保留名称像 "Row 0"、"Row 1"……的行
    rows = [r for r in all_rows if re.match(r"^Row \d+$", r.element_info.name)]
    print(f"共找到 {len(rows)} 条 Session： {[r.element_info.name for r in rows]}")

    # 5. 循环处理每一行
    for idx, row_elem in enumerate(rows, start=1):
        name = row_elem.element_info.name
        print(f"[{idx}/{len(rows)}] 选中并运行 PCAP → {name}")

        # 5.1 切回 Sessions，确保当前行可见
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.5)

        # wrap row，方便调用 wrapper 方法
        row = row_elem.wrapper_object()

        # 5.2 点击它下面的第一个 DataItem（第一列单元格）
        data_cells = row.children(control_type="DataItem")
        if data_cells:
            data_cells[0].click_input()
        else:
            # 回退方案：直接点击行中心
            row.click_input()
        time.sleep(0.2)

        # 5.3 点击右侧 ▶ 按钮启动 PCAP
        pcap_start_btn.click_input()

        # 5.4 等待按钮先变灰（作业启动），再恢复可用（作业完成）
        wait_until(5,  1, lambda: not pcap_start_btn.is_enabled(),
                   retry_exception=RuntimeError)
        wait_until(600,1, lambda:     pcap_start_btn.is_enabled(),
                   retry_exception=RuntimeError)

        print(f"  ✓ {name} 完成")

    print("🎉 所有 Session 都已跑完！")

if __name__ == "__main__":
    main()
