import time
import re
from pywinauto import Application, timings

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # ← 改成你的实际路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 2. 定位左侧 TabControl（用于切到 Sessions）
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP 区域的复选框和启动按钮
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    )

    # 4. 确保全局 PCAP 已勾选
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 5. 切到 Sessions 标签页，获取所有行的名字
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    all_items = sessions_table.descendants(control_type="DataItem")
    pattern   = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip()
                 for itm in all_items
                 if itm.element_info.name and pattern.match(itm.element_info.name)]

    print(f"共找到 {len(row_names)} 条 Session：{row_names}")

    # 6. 依次处理每一行：选中→运行→等待5秒→取消选中
    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 处理 → {name}")

        # 6.1 切回 Sessions 并刷新表格元素
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(
            auto_id="SessionView", control_type="Table"
        ).wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")

        # 6.2 找到当前行对应的 DataItem 并点击选中
        cell = next(it for it in all_items if it.element_info.name.strip() == name)
        cell.click_input()
        time.sleep(0.2)

        # 6.3 点击 ▶ 启动 PCAP
        pcap_start_btn.click_input()

        # 6.4 固定等待 5 秒
        time.sleep(5)

        # 6.5 再次点击同一个 DataItem，取消选中
        cell.click_input()
        time.sleep(0.2)

        print(f"  ✓ {name} 已运行并取消选中")

    print("🎉 所有 Session 已依次运行完毕！")

if __name__ == "__main__":
    main()
