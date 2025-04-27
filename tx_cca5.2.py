import time, re
from pywinauto import Application

def main():
    # 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # ← 改成你的实际路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 定位左侧的 Sessions Tab 和右侧 PCAP Pane
    tab            = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    pcap_pane      = dlg.child_window(auto_id="PCAP",     control_type="Pane")
    pcap_checkbox  = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    ).wrapper_object()

    # 确保 PCAP 复选框已勾
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 切到 Sessions，拿到表格
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 1) 用 DataItem 拿每行的名字（" Row 0", " Row 1", ...）
    all_dataitems = sessions_table.descendants(control_type="DataItem")
    name_pat      = re.compile(r"^\s*Row \d+$")
    row_items     = [it for it in all_dataitems
                     if it.element_info.name and name_pat.match(it.element_info.name)]
    row_names     = [it.element_info.name.strip() for it in row_items]

    # 2) 拿所有 CheckBox，去掉第一个“全选”框，剩下的就是每行的复选框
    all_boxes = [cb.wrapper_object() 
                 for cb in sessions_table.descendants(control_type="CheckBox")]
    row_boxes = all_boxes[1:1+len(row_names)]

    print(f"共找到 {len(row_names)} 条 Session：", row_names)

    # 依次点击复选框 + 启动 PCAP
    for idx, (name, box) in enumerate(zip(row_names, row_boxes), start=1):
        print(f"[{idx}/{len(row_names)}] 选中 → {name}")

        # ① 点击这一行的复选框
        box.click_input()
        time.sleep(0.2)

        # ② 点击右侧的 ▶ 按钮启动 PCAP
        pcap_start_btn.click_input()

        # ③ 等待按钮先变灰（表示启动），再重新可用（表示完成）
        pcap_start_btn.wait_not("enabled", timeout=5)
        pcap_start_btn.wait("enabled", timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 都已跑完！")

if __name__ == "__main__":
    main()
