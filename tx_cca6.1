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

    # 2. 定位左侧 Sessions Tab 和右侧 PCAP Pane
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    pcap_pane      = dlg.child_window(auto_id="PCAP", control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    )

    # 2.1 确保 PCAP 复选框已勾
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 3. 切到 Sessions，拿到表格
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 4. 找到每行最左侧的复选框（跳过第一个“全选”）
    all_boxes   = sessions_table.descendants(control_type="CheckBox")
    row_boxes   = [box.wrapper_object() for box in all_boxes[1:]]  # 1: 跳过全选
    # 只保留和 DataItem 行数一样的那几项
    # 先找到行名用来打印
    all_dataitems   = sessions_table.descendants(control_type="DataItem")
    name_pattern    = re.compile(r"^\s*Row \d+$")
    row_items       = [it for it in all_dataitems if it.element_info.name and name_pattern.match(it.element_info.name)]
    # 保证复选框和行数对齐
    row_boxes = row_boxes[:len(row_items)]

    print(f"共找到 {len(row_items)} 条 Session：",
          [it.element_info.name.strip() for it in row_items])

    # 5. 逐行选中、跑 PCAP、取消选中
    for idx, (item, box) in enumerate(zip(row_items, row_boxes), start=1):
        name = item.element_info.name.strip()
        print(f"[{idx}/{len(row_items)}] 处理 → {name}")

        # 5.1 点击复选框选中
        box.click_input()
        time.sleep(0.2)

        # 5.2 点击 ▶ 启动 PCAP
        pcap_start_btn.click_input()

        # 5.3 等待按钮禁用（如果会禁用）
        try:
            pcap_start_btn.wait_not("enabled", timeout=5)
            print("    → 作业已启动（按钮禁用）")
        except timings.TimeoutError:
            print("    → 按钮未禁用，作业可能在后台启动")

        # 5.4 等待按钮重新可用（作业完成）
        pcap_start_btn.wait("enabled", timeout=600)
        print(f"  ✓ {name} 完成")

        # 5.5 点击复选框取消该行选中
        box.click_input()
        time.sleep(0.2)

    print("🎉 所有 Session 都已跑完并取消选中！")

if __name__ == "__main__":
    main()
