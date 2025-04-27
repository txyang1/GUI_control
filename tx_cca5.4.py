import time
import re
from pywinauto import Application

def main():
    # 1. 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # ← 改成你的实际路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 2. 定位 Sessions Tab 和 PCAP 区域
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")

    pcap_pane = dlg.child_window(auto_id="PCAP", control_type="Pane")
    # 复选框要拿 wrapper 才能 get_toggle_state()
    pcap_checkbox = pcap_pane.child_window(
        auto_id="activateCheckBox", control_type="CheckBox"
    ).wrapper_object()
    # 启动按钮保留为 Spec，这样我们才能用 wait_not / wait
    pcap_start_btn = pcap_pane.child_window(
        auto_id="startButton", control_type="Button"
    )

    # 2.1 确保 PCAP 复选框被勾上
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 3. 切到 Sessions，拿到表格
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(
        auto_id="SessionView", control_type="Table"
    ).wrapper_object()

    # 4. 找到左侧行的 DataItem（名字形如 " Row 0", " Row 1"...）
    all_dataitems = sessions_table.descendants(control_type="DataItem")
    select_pattern = re.compile(r"^\s*Row \d+$")
    select_cells = [
        it for it in all_dataitems
        if it.element_info.name and select_pattern.match(it.element_info.name)
    ]

    print(f"共找到 {len(select_cells)} 条 Session：",
          [c.element_info.name.strip() for c in select_cells])

    # 5. 依次选中每一行，运行 PCAP，等待完成
    for idx, cell in enumerate(select_cells, start=1):
        name = cell.element_info.name.strip()
        print(f"[{idx}/{len(select_cells)}] 选中 → {name}")

        # 5.1 点击行前面的 DataItem
        cell.click_input()
        time.sleep(0.2)

        # 5.2 点击启动按钮
        pcap_start_btn.click_input()

        # 5.3 等待按钮先从 enabled→disabled（启动），再从 disabled→enabled（完成）
        pcap_start_btn.wait_not("enabled", timeout=5)
        pcap_start_btn.wait("enabled",     timeout=600)

        print(f"  ✓ {name} 完成")

    print("🎉 全部 Session 都已跑完！")

if __name__ == "__main__":
    main()
