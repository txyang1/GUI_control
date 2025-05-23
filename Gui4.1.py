import time
import re
import argparse
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

# （略）PCAP_INPUT_FOLDER、PCAP_OUTPUT_FOLDER、MTRE_INPUT_FOLDER、MTRE_OUTPUT_FOLDER

def wait_for_finish(btn, timeout=60, retry_interval=0.5):
    timings.WaitUntil(timeout, retry_interval, lambda: btn.is_enabled())

def run_for_items(dlg, pane_auto_id: str, tab_name: str, batch_size: int = 1):
    """
    通用：在指定 tab_name（"Sessions" 或 "Files"）下，
    按 batch_size 大小循环选中 DataItem 并点击 start。
    """
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab").wrapper_object()
    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    start_btn      = func_pane.child_window(auto_id="startButton", control_type="Button").wrapper_object()

    # 先收集所有 item 名称
    tab.child_window(title=tab_name, control_type="TabItem").select()
    time.sleep(0.5)
    table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    items = table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    names = [it.element_info.name.strip() for it in items
             if it.element_info.name and name_pat.match(it.element_info.name)]

    for batch_start in range(0, len(names), batch_size):
        batch = names[batch_start:batch_start + batch_size]
        print(f"[{tab_name} {batch_start+1}-{batch_start+len(batch)}/{len(names)}] 运行 → {batch}")

        # 重新切到 tab，重新定位 table
        tab.child_window(title=tab_name, control_type="TabItem").select()
        time.sleep(0.2)
        table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()

        # 选中这一批
        for idx, name in enumerate(batch):
            cell = next(it.wrapper_object() for it in table.descendants(control_type="DataItem")
                        if it.element_info.name.strip() == name)
            # 保证可见
            try:
                cell.scroll_into_view()
            except Exception:
                table.set_focus()
                while cell.element_info.element.CurrentIsOffscreen:
                    send_keys("{PGDN}")
                    time.sleep(0.1)
            time.sleep(0.1)

            # 单选或多选
            if batch_size == 1 or idx == 0:
                cell.click_input()
            else:
                cell.click_input(ctrl=True)
            time.sleep(0.1)

        # 点击 Start，等待完成
        start_btn.click_input()
        wait_for_finish(start_btn)
        print(f"  ✓ 已完成 batch {batch_start//batch_size + 1}")

        # 取消选中
        for name in batch:
            cell = next(it.wrapper_object() for it in table.descendants(control_type="DataItem")
                        if it.element_info.name.strip() == name)
            cell.click_input(ctrl=True)
            time.sleep(0.05)

def set_folder(dlg, combo_auto_id: str, path: str):
    combo = dlg.child_window(auto_id=combo_auto_id, control_type="ComboBox")
    combo.child_window(control_type="Edit").set_edit_text(path)
    time.sleep(0.2)
    for title in ("Open", "Browse", "openExplorerButton", "browseButton"):
        try:
            combo.parent().child_window(title=title, control_type="Button").click_input()
            time.sleep(0.5)
            break
        except Exception:
            continue

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["sessions", "files"], default="sessions",
                        help="选择操作 Sessions 还是 Files 标签页")
    parser.add_argument("--batch-size", type=int, default=1,
                        help="仅当 mode=files 时生效，每 N 个文件一起运行")
    args = parser.parse_args()

    app = Application(backend="uia").start(r"你的.exe 路径")
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # —— PCAP 设置 & 运行 —— 
    set_folder(dlg, "inputLocationComboBox", PCAP_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", PCAP_OUTPUT_FOLDER)
    pcap_cb = dlg.child_window(auto_id="PCAP", control_type="Pane") \
                .child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if pcap_cb.get_toggle_state() == 0:
        pcap_cb.click_input(); time.sleep(0.2)

    tab_name = "Sessions" if args.mode == "sessions" else "Files"
    print(f"==== PCAP: 在 {tab_name} 下运行 ====")
    run_for_items(dlg, pane_auto_id="PCAP", tab_name=tab_name, batch_size=args.batch_size)

    pcap_cb.click_input(); time.sleep(0.5)

    # —— MTRE 设置 & 运行 ——
    set_folder(dlg, "inputLocationComboBox", MTRE_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", MTRE_OUTPUT_FOLDER)
    mtre_cb = dlg.child_window(auto_id="MTRE", control_type="Pane") \
                .child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if mtre_cb.get_toggle_state() == 0:
        mtre_cb.click_input(); time.sleep(0.2)

    print(f"==== MTRE: 在 {tab_name} 下运行 ====")
    run_for_items(dlg, pane_auto_id="MTRE", tab_name=tab_name, batch_size=args.batch_size)

    mtre_cb.click_input(); time.sleep(0.5)
    print("🎉 全部完成！")

if __name__ == "__main__":
    main()
