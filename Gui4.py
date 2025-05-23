import time
import re
import argparse
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

# 路径按需修改
PCAP_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\trace_Example_No_Copy_Outside"
PCAP_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output"
MTRE_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\MT_RE"
MTRE_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output2"

def wait_for_finish(btn, timeout=60, retry_interval=0.5):
    timings.WaitUntil(timeout, retry_interval, lambda: btn.is_enabled())

def run_for_rows(dlg, pane_auto_id: str):
    """逐行运行 Sessions"""
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab").wrapper_object()
    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button").wrapper_object()

    # 收集所有 Row 名称
    table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    items = table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [it.element_info.name.strip() for it in items if it.element_info.name and name_pat.match(it.element_info.name)]

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} Sessions → {name}")
        # 重新定位 Sessions tab & table
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
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
        time.sleep(0.2)

        # 点击、运行、取消
        cell.click_input(); time.sleep(0.2)
        func_start_btn.click_input()
        wait_for_finish(func_start_btn)
        cell.click_input(); time.sleep(0.2)
        print(f"  ✓ 完成 {name}")

def run_for_files(dlg, pane_auto_id: str, batch_size: int = 1):
    """按文件处理 Files 页面，batch_size>1 时每 batch_size 个一起启动一次"""
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab").wrapper_object()
    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button").wrapper_object()

    # 收集所有 File 名称
    tab.child_window(title="Files", control_type="TabItem").select()
    time.sleep(0.5)
    table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    items = table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    file_names = [it.element_info.name.strip() for it in items if it.element_info.name and name_pat.match(it.element_info.name)]

    # 按 batch_size 分片
    for batch_start in range(0, len(file_names), batch_size):
        batch = file_names[batch_start:batch_start + batch_size]
        print(f"[Files {batch_start+1}-{batch_start+len(batch)}/{len(file_names)}] 运行 {pane_auto_id} Files → {batch}")

        # 重新切回 Files tab & 定位 table
        tab.child_window(title="Files", control_type="TabItem").select()
        time.sleep(0.2)
        table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()

        # 多选：第一个普通点击，后续 Ctrl+点击
        for i, name in enumerate(batch):
            cell = next(it.wrapper_object() for it in table.descendants(control_type="DataItem")
                        if it.element_info.name.strip() == name)
            try:
                cell.scroll_into_view()
            except Exception:
                table.set_focus()
                while cell.element_info.element.CurrentIsOffscreen:
                    send_keys("{PGDN}")
                    time.sleep(0.1)
            time.sleep(0.1)

            if i == 0:
                cell.click_input()
            else:
                cell.click_input(ctrl=True)
            time.sleep(0.1)

        # 启动这一批
        func_start_btn.click_input()
        wait_for_finish(func_start_btn)
        print(f"  ✓ 完成 batch {batch_start//batch_size + 1}")

        # 取消选中：Ctrl+点击每个已选文件
        for name in batch:
            cell = next(it.wrapper_object() for it in table.descendants(control_type="DataItem")
                        if it.element_info.name.strip() == name)
            cell.click_input(ctrl=True)
            time.sleep(0.05)

def set_folder(dlg, combo_auto_id: str, path: str):
    combo = dlg.child_window(auto_id=combo_auto_id, control_type="ComboBox")
    combo.child_window(control_type="Edit").set_edit_text(path)
    time.sleep(0.2)
    parent = combo.parent()
    for title in ("Open", "Browse", "openExplorerButton", "browseButton"):
        try:
            btn = parent.child_window(title=title, control_type="Button")
            btn.click_input()
            time.sleep(0.5)
            break
        except Exception:
            continue

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["sessions", "files"], default="sessions",
                        help="运行模式：sessions（逐Session） 或 files（逐File）")
    parser.add_argument("--batch-size", type=int, default=1,
                        help="仅当 --mode=files 时生效，每隔多少个文件一起运行")
    args = parser.parse_args()

    app = Application(backend="uia").start(
        r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\15_ViGEM_CCA\CCA Tool\bin\ViGEM.CCA.Converter.Gui.exe"
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # PCAP 设置 & 执行
    print("==== PCAP 设置 ====")
    set_folder(dlg, "inputLocationComboBox", PCAP_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", PCAP_OUTPUT_FOLDER)
    pcap_cb = dlg.child_window(auto_id="PCAP", control_type="Pane")\
                .child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if pcap_cb.get_toggle_state() == 0:
        pcap_cb.click_input(); time.sleep(0.2)

    print(f"==== PCAP: 依次运行每个 {'Session' if args.mode=='sessions' else 'File'} ====")
    if args.mode == "sessions":
        run_for_rows(dlg, pane_auto_id="PCAP")
    else:
        run_for_files(dlg, pane_auto_id="PCAP", batch_size=args.batch_size)

    pcap_cb.click_input(); time.sleep(0.5)

    # MTRE 设置 & 执行
    print("==== MTRE 设置 ====")
    set_folder(dlg, "inputLocationComboBox", MTRE_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", MTRE_OUTPUT_FOLDER)
    mtre_cb = dlg.child_window(auto_id="MTRE", control_type="Pane")\
                .child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if mtre_cb.get_toggle_state() == 0:
        mtre_cb.click_input(); time.sleep(0.2)

    print(f"==== MTRE: 依次运行每个 {'Session' if args.mode=='sessions' else 'File'} ====")
    if args.mode == "sessions":
        run_for_rows(dlg, pane_auto_id="MTRE")
    else:
        run_for_files(dlg, pane_auto_id="MTRE", batch_size=args.batch_size)

    mtre_cb.click_input(); time.sleep(0.5)

    print("🎉 所有任务已执行完毕！")

if __name__ == "__main__":
    main()
