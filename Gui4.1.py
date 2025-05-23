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

def run_for_rows(dlg, pane_auto_id: str, batch_size: int = 1):
    """
    在 Files 页签中，按 batch_size 批量选中行并点击 ▶ 运行，支持 batch_size=1（逐条）或更大。
    """
    # 切换到 Files 标签页
    dlg.child_window(auto_id="tabControl1", control_type="Tab") \
       .child_window(title="Files", control_type="TabItem").select()
    time.sleep(0.5)

    # 收集所有行名
    files_table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()
    all_items = files_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [
        itm.element_info.name.strip()
        for itm in all_items
        if itm.element_info.name and name_pat.match(itm.element_info.name)
    ]

    # 启动按钮
    func_start_btn = dlg.child_window(auto_id=pane_auto_id, control_type="Pane") \
                        .child_window(auto_id="startButton", control_type="Button")

    total = len(row_names)
    for start in range(0, total, batch_size):
        batch = row_names[start:start + batch_size]
        print(f"[Files {start+1}-{min(start+batch_size, total)}/{total}] 运行 {pane_auto_id} → {batch}")

        # 切回 Files tab，重新定位 table
        dlg.child_window(auto_id="tabControl1", control_type="Tab") \
           .child_window(title="Files", control_type="TabItem").select()
        time.sleep(0.2)
        files_table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()

        # 批量选中
        for i, name in enumerate(batch):
            itm = next(
                it for it in files_table.descendants(control_type="DataItem")
                if it.element_info.name.strip() == name
            )
            # 保证可见
            try:
                itm.scroll_into_view()
                time.sleep(0.1)
            except Exception:
                files_table.set_focus()
                while itm.element_info.element.CurrentIsOffscreen:
                    send_keys("{PGDN}")
                    time.sleep(0.1)
            time.sleep(0.1)

            # 第一个直接点击，后续 Ctrl 多选
            if i == 0:
                itm.click_input()
            else:
                itm.click_input(ctrl=True)
            time.sleep(0.1)

        # 点击 ▶ 启动
        func_start_btn.click_input()
        timings.wait_until(
            timeout=600,
            retry_interval=1,
            func=lambda: func_start_btn.is_enabled()
        )
        time.sleep(0.2)
        print(f"  ✓ 完成 batch {start//batch_size + 1}")

        # 取消选中
        for name in batch:
            itm = next(
                it for it in files_table.descendants(control_type="DataItem")
                if it.element_info.name.strip() == name
            )
            itm.click_input(ctrl=True)
            time.sleep(0.05)

def run_mode(dlg, mode: str, batch_size: int):
    """
    对单个模式（PCAP 或 MTRE）进行输入/输出设置、激活、按 Files 批量/逐条运行和重置。
    """
    folders = FOLDERS[mode]
    print(f"==== {mode}: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", folders["input"])
    set_folder(dlg, "outputFolderComboBox", folders["output"])

    # 勾选模式对应的复选框
    pane = dlg.child_window(auto_id=mode, control_type="Pane")
    checkbox = pane.child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if checkbox.get_toggle_state() == 0:
        checkbox.click_input()
        time.sleep(0.2)

    print(f"==== {mode}: 依次运行每个 Files（batch_size={batch_size}） ====")
    run_for_rows(dlg, pane_auto_id=mode, batch_size=batch_size)

    # 重置复选框
    print(f"==== {mode} 完成后：重置 {mode} ====")
    checkbox.click_input()
    time.sleep(0.5)

def main():
    parser = argparse.ArgumentParser(
        description="运行 ViGEM CCA 转换工具，可指定 PCAP、MTRE 或同时运行，并可设置 Files 的批量大小。"
    )
    parser.add_argument(
        '--mode', choices=['PCAP', 'MTRE', 'ALL'], default='ALL',
        help="选择要运行的模式：PCAP、MTRE 或 ALL（同时运行两者）"
    )
    parser.add_argument(
        '--batch-size', type=int, default=1,
        help="Files 标签下每隔多少个文件一起运行一次（默认为 1，即逐个运行）"
    )
    args = parser.parse_args()

    # 启动并连接主窗口
    app = Application(backend="uia").start(
        r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\15_ViGEM_CCA\CCA Tool\bin\ViGEM.CCA.Converter.Gui.exe"
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 根据参数选择运行模式
    modes = [args.mode] if args.mode != 'ALL' else ['PCAP', 'MTRE']
    for m in modes:
        run_mode(dlg, m, args.batch_size)

    print("🎉 所有选定模式的 Files 批量运行完毕！")

if __name__ == "__main__":
    main()
