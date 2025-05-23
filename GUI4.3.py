import time
import re
import argparse
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

# 根据实际环境修改这些路径
PCAP_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\trace_Example_No_Copy_Outside"
PCAP_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output"
MTRE_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\MT_RE"
MTRE_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output2"

# 模式对应的文件夹映射
FOLDERS = {
    "PCAP": {"input": PCAP_INPUT_FOLDER, "output": PCAP_OUTPUT_FOLDER},
    "MTRE": {"input": MTRE_INPUT_FOLDER, "output": MTRE_OUTPUT_FOLDER}
}

def set_folder(dlg, combo_auto_id: str, path: str):
    combo = dlg.child_window(auto_id=combo_auto_id, control_type="ComboBox")
    combo.child_window(control_type="Edit").set_edit_text(path)
    time.sleep(0.2)
    parent = combo.parent()
    for title in ("Open", "Browse", "openExplorerButton", "browseButton"):
        try:
            parent.child_window(title=title, control_type="Button").click_input()
            time.sleep(0.5)
            break
        except Exception:
            continue

def ensure_visible(item, table):
    """
    把单个 DataItem 滚动到可视区：
      1. 优先尝试 item.scroll_into_view()
      2. 否则将焦点给 table，根据 item 相对于 table 的位置
         循环使用 PageUp/PageDown，直到它完全可见。
    """
    try:
        item.scroll_into_view()
        return
    except Exception:
        pass

    table.set_focus()
    time.sleep(0.05)

    while True:
        itm_rect = item.rectangle()
        tbl_rect = table.rectangle()
        # 已经完全可见
        if itm_rect.top >= tbl_rect.top and itm_rect.bottom <= tbl_rect.bottom:
            break
        # 如果在下方，向下翻页
        if itm_rect.bottom > tbl_rect.bottom:
            send_keys("{PGDN}")
        # 如果在上方，向上翻页
        elif itm_rect.top < tbl_rect.top:
            send_keys("{PGUP}")
        time.sleep(0.1)

def run_for_rows(dlg, pane_auto_id: str, batch_size: int = 1):
    """
    在 Files 页签中，按 batch_size 分批：
      1. 依次选中本批次每一行
      2. 批次内所有行选中后，仅点击一次 Run
      3. 等待完成
      4. 依次取消选中每一行（每次都重新定位并 ensure_visible）
    """
    # 切换到 Files 标签页
    dlg.child_window(auto_id="tabControl1", control_type="Tab") \
       .child_window(title="Files", control_type="TabItem").select()
    time.sleep(0.5)

    # 收集所有行名
    table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()
    items = table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [it.element_info.name.strip()
                 for it in items
                 if it.element_info.name and name_pat.match(it.element_info.name)]

    # Run 按钮引用
    run_btn = dlg.child_window(auto_id=pane_auto_id, control_type="Pane") \
                 .child_window(auto_id="startButton", control_type="Button")

    total = len(row_names)
    for batch_start in range(0, total, batch_size):
        batch = row_names[batch_start:batch_start + batch_size]
        print(f"\n===== 批次 {batch_start//batch_size + 1} "
              f"({batch_start+1}-{min(batch_start+batch_size, total)}/{total}) → {batch}")

        # 1. 依次选中每一行
        for name in batch:
            dlg.child_window(auto_id="tabControl1", control_type="Tab") \
               .child_window(title="Files", control_type="TabItem").select()
            time.sleep(0.2)
            table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()

            itm = next(it for it in table.descendants(control_type="DataItem")
                       if it.element_info.name.strip() == name)
            ensure_visible(itm, table)
            time.sleep(0.1)
            itm.click_input()
            time.sleep(0.1)
            print(f"    已选中 {name}")

        # 2. 批次内所有行选中后，仅点击一次 Run
        print("    ▶ 点击 Run")
        run_btn.click_input()
        timings.wait_until(timeout=600, retry_interval=1, func=lambda: run_btn.is_enabled())
        time.sleep(0.2)
        print(f"  ✓ 批次 {batch_start//batch_size + 1} 完成")

        # 3. 依次取消选中每一行
        for name in batch:
            dlg.child_window(auto_id="tabControl1", control_type="Tab") \
               .child_window(title="Files", control_type="TabItem").select()
            time.sleep(0.2)
            table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()

            itm = next(it for it in table.descendants(control_type="DataItem")
                       if it.element_info.name.strip() == name)
            ensure_visible(itm, table)
            time.sleep(0.05)
            itm.click_input()
            time.sleep(0.1)
            print(f"    已取消选中 {name}")

        print(f"===== 结束批次 {batch_start//batch_size + 1} =====")

def run_mode(dlg, mode: str, batch_size: int):
    """
    对 PCAP 或 MTRE 模式，设置输入/输出、勾选复选框、在 Files 页签分批运行、再重置复选框。
    """
    folders = FOLDERS[mode]
    print(f"==== {mode}: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", folders["input"])
    set_folder(dlg, "outputFolderComboBox", folders["output"])

    pane = dlg.child_window(auto_id=mode, control_type="Pane")
    checkbox = pane.child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if checkbox.get_toggle_state() == 0:
        checkbox.click_input()
        time.sleep(0.2)

    print(f"==== {mode}: 依次运行每个 Files（batch_size={batch_size}） ====")
    run_for_rows(dlg, pane_auto_id=mode, batch_size=batch_size)

    print(f"==== {mode} 完成后：重置 {mode} ====")
    checkbox.click_input()
    time.sleep(0.5)

def main():
    parser = argparse.ArgumentParser(
        description="运行 ViGEM CCA-Converter GUI，可指定 PCAP/MTRE 模式并设置 Files 批量大小。"
    )
    parser.add_argument(
        '--mode', choices=['PCAP','MTRE','ALL'], default='ALL',
        help="选择要运行的模式：PCAP、MTRE 或 ALL（同时运行两者）"
    )
    parser.add_argument(
        '--batch-size', type=int, default=1,
        help="Files 标签页下每隔多少个文件分为一批运行（默认为 1，逐个执行）"
    )
    args = parser.parse_args()

    app = Application(backend="uia").start(
        r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\15_ViGEM_CCA\CCA Tool\bin\ViGEM.CCA.Converter.Gui.exe"
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    modes = [args.mode] if args.mode != 'ALL' else ['PCAP','MTRE']
    for m in modes:
        run_mode(dlg, m, args.batch_size)

    print("🎉 所有选定模式的 Files 批量运行完毕！")

if __name__ == "__main__":
    main()
