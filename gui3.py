import time
import re
import argparse
from pywinauto import Application, timings

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
            btn = parent.child_window(title=title, control_type="Button")
            btn.click_input()
            time.sleep(0.5)
            break
        except Exception:
            continue


def run_for_rows(dlg, pane_auto_id: str):
    """
    在 Sessions 页签中遍历所有行，点击 ▶ 并等待运行完成。
    """
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    all_items = sessions_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip() for itm in all_items if itm.element_info.name and name_pat.match(itm.element_info.name)]

    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button")

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")

        # 重新选中行
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")
        cell = next(it for it in all_items if it.element_info.name.strip() == name)
        cell.click_input()
        time.sleep(0.2)

        # 点击 ▶ 启动
        func_start_btn.click_input()

        # 等待启动按钮可用（运行结束）
        timings.wait_until(
            timeout=600,          # 最多等 600 秒
            retry_interval=1,     # 每秒检查一次
            func=lambda: func_start_btn.is_enabled()
        )
        time.sleep(0.2)

        # 取消选中
        cell.click_input()
        time.sleep(0.2)
        print(f"  ✓ {pane_auto_id} 已完成 {name}")


def run_mode(dlg, mode: str):
    """
    对单个模式（PCAP 或 MTRE）进行输入/输出设置、激活、运行和重置。
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

    print(f"==== {mode}: 依次运行每个 Session ====")
    run_for_rows(dlg, pane_auto_id=mode)

    # 重置复选框
    print(f"==== {mode} 完成后：重置 {mode} ====")
    checkbox.click_input()
    time.sleep(0.5)


def main():
    parser = argparse.ArgumentParser(description="运行 ViGEM CCA 转换工具，可指定 PCAP、MTRE 或同时运行。")
    parser.add_argument(
        '--mode', choices=['PCAP', 'MTRE', 'ALL'], default='ALL',
        help="选择要运行的模式：PCAP、MTRE 或 ALL（默认同时运行两者）"
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
    for mode in modes:
        run_mode(dlg, mode)

    print("🎉 所有选定模式的会话均已运行完毕！")

if __name__ == "__main__":
    main()
