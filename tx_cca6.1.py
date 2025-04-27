import time, re
from pywinauto import Application, timings

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # ← 改成你的实际路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 2. 定位 Sessions Tab
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    # 3. 定位右侧 PCAP Pane
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    )

    # 确保 PCAP 全局复选框已勾上
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 4. 切到 Sessions，拿到表格
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 5. 找到所有行复选框 (第一个是全选，去掉它)
    all_boxes = sessions_table.descendants(control_type="CheckBox")
    row_boxes = [cb.wrapper_object() for cb in all_boxes[1:]]  # 每行对应一个 CheckBox

    # 6. （可选）打印一下行数
    print(f"共找到 {len(row_boxes)} 条 Session 行，将依次运行它们的 PCAP。")

    # 7. 依次处理
    for idx, box in enumerate(row_boxes, start=1):
        print(f"[{idx}/{len(row_boxes)}] 运行第 {idx} 行的 PCAP …")

        # 7.1 先取消所有行的勾选
        for b in row_boxes:
            if b.get_toggle_state() == 1:
                b.click_input()
                time.sleep(0.1)

        # 7.2 勾选当前行
        box.click_input()
        time.sleep(0.1)

        # 7.3 点击 ▶ 启动 PCAP
        pcap_start_btn.click_input()

        # 7.4 等待按钮禁用（如果会禁用）
        try:
            pcap_start_btn.wait_not("enabled", timeout=5)
        except timings.TimeoutError:
            pass  # 有些版本可能不禁用，直接继续

        # 7.5 等待按钮重新可用（表示作业完成）
        pcap_start_btn.wait("enabled", timeout=600)
        print(f"  ✓ 第 {idx} 行完成")

    print("🎉 已全部运行完毕！")

if __name__ == "__main__":
    main()
