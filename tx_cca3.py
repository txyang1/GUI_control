import time
import re
from pywinauto import Application

def main():
    # 1. 启动并连接到主窗口
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)

    # 2. 定位左侧的 Sessions 页签控制
    tab_control = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP 区域下的控件（spec 而非 wrapper）
    pcap_pane          = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox_spec = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    )
    pcap_start_spec    = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    )

    # 3.1 确保 PCAP 复选框已勾选
    pcap_checkbox = pcap_checkbox_spec.wrapper_object()
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 4. 切到 Sessions，取出所有 Row N 行
    tab_control.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    # 找到表格并拿到所有 Custom 行（包括 Top Row，但我们要过滤）
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    )
    table_wrap = sessions_table.wrapper_object()
    all_rows = table_wrap.children(control_type="Custom")

    # 只保留 “Row 0, Row 1, …” 这样的行
    rows = [r for r in all_rows if re.match(r"Row \d+", r.element_info.name)]
    print(f"共找到 {len(rows)} 条要处理的 Session：",
          [r.element_info.name for r in rows])

    # 5. 依次处理每一行
    for idx, row_wrap in enumerate(rows, start=1):
        name = row_wrap.element_info.name
        print(f"[{idx}/{len(rows)}] 选中并运行 PCAP → {name}")

        # 5.1 选中当前行
        row_wrap.click_input()
        time.sleep(0.2)

        # 5.2 点击 PCAP 的 ▶ 按钮启动
        pcap_start_spec.click_input()

        # 5.3 等待按钮先禁用（作业开始），再重新启用（作业结束）
        pcap_start_spec.wait_not("enabled", timeout=5)
        pcap_start_spec.wait("enabled",     timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 所有 Session 均已处理完毕！")

if __name__ == "__main__":
    main()
