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

    # 2. 定位 TabControl（用于切换 Sessions）
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP Pane，并拿到全局勾选框 + ▶ 按钮
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    )

    # 4. 确保 PCAP 已经勾上
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 5. 切到 Sessions，先抓一次所有行名
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    all_items = sessions_table.descendants(control_type="DataItem")
    pattern   = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip()
                 for itm in all_items if itm.element_info.name and pattern.match(itm.element_info.name)]

    print(f"共找到 {len(row_names)} 条 Session: {row_names}")

    # 6. 逐行处理
    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 处理 → {name}")

        # 6.1 每次都重新切回 Sessions 并刷新表格和 DataItem
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(
            auto_id="SessionView", control_type="Table"
        ).wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")

        # 6.2 找到对应行的 DataItem 并点击，完成选中
        cell = next(it for it in all_items if it.element_info.name.strip() == name)
        cell.click_input()
        time.sleep(0.2)

        # 6.3 点击 PCAP ▶ 运行
        pcap_start_btn.click_input()

        # 6.4 尝试等待按钮变灰（启动），有的版本不会禁用则跳过
        try:
            pcap_start_btn.wait_not("enabled", timeout=5)
        except timings.TimeoutError:
            pass

        # 6.5 等待按钮重新可用（完成）
        pcap_start_btn.wait("enabled", timeout=600)
        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 已处理完毕！")

if __name__ == "__main__":
    main()
