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

    # 2. 定位 TabControl（用于切回 Sessions）
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP Pane，并拿到全局复选框 + ▶ 启动按钮
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    )

    # 4. 确保 PCAP 复选框已勾上
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 5. 切到 Sessions，抓取所有行的名称
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 用 DataItem 名称来识别行
    all_items = sessions_table.descendants(control_type="DataItem")
    pattern   = re.compile(r"^\s*Row \d+$")
    row_names = [w.element_info.name.strip() for w in all_items
                 if w.element_info.name and pattern.match(w.element_info.name)]

    print(f"共找到 {len(row_names)} 条 Session: {row_names}")

    # 6. 依次处理每一行
    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 → {name}")

        # 6.1 每次都切回 Sessions，重新定位表格和行 wrapper
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(
            auto_id="SessionView", control_type="Table"
        ).wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")

        # 6.2 找到对应行的 wrapper 并用 select()（单选模式）
        row_wrapper = next(
            w for w in all_items
            if w.element_info.name and w.element_info.name.strip() == name
        )
        row_wrapper.select()   # 这一行就被单独选中了
        time.sleep(0.2)

        # 6.3 点击 ▶ 按钮启动 PCAP
        pcap_start_btn.click_input()

        # 6.4 尝试等待按钮禁用（启动），失败就跳过
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
