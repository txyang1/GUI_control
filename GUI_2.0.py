import time
import re
import os
import argparse
from pywinauto import Application, timings

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
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    all_items = sessions_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip() for itm in all_items if itm.element_info.name and name_pat.match(itm.element_info.name)]

    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button").wrapper_object()

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")

        # 选中行并启动
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")
        cell = next(it for it in all_items if it.element_info.name.strip() == name)
        cell.click_input()
        time.sleep(0.2)

        # 点击 ▶ 启动并动态等待完成
        func_start_btn.click_input()
        # 等待按钮重新可用 (或图标切换回▶)，表示处理完成
        timings.wait_until_passes(
            timeout=300,       # 最长等待5分钟
            retry_interval=1,  # 每秒检查一次
            func=lambda: func_start_btn.is_enabled()
        )

        # 取消选中
        cell.click_input()
        time.sleep(0.2)

        print(f"  ✓ {pane_auto_id} 已完成 {name}")

def main():
    parser = argparse.ArgumentParser(description="自动运行 ViGEM CCA-Converter 的 PCAP 或 MTRE 流程")
    parser.add_argument("-i", "--input-folder", required=True, help="输入文件夹路径 (PCAP 或 MT_RE)")
    parser.add_argument("-o", "--output-folder", required=True, help="输出文件夹路径")
    args = parser.parse_args()

    basename = os.path.basename(os.path.normpath(args.input_folder)).lower()
    if basename in ("mt_re", "mtre"):
        mode = "MT_RE"
    else:
        mode = "PCAP"

    # 启动应用
    app = Application(backend="uia").start(r"C:\\Users\\qxz5y3m\\Desktop\\Transfer_Tool_for_trace\\15_ViGEM_CCA\\CCA Tool\\bin\\ViGEM.CCA.Converter.Gui.exe")
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 设置输入/输出文件夹
    print(f"==== {mode}: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", args.input_folder)
    set_folder(dlg, "outputFolderComboBox", args.output_folder)

    # 确认并勾选对应模块
    print(f"==== 确认 {mode} 方块已勾选 ====")
    checkbox = dlg.child_window(auto_id=mode, control_type="Pane") \
                   .child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if checkbox.get_toggle_state() == 0:
        checkbox.click_input()
        time.sleep(0.2)

    # 循环运行
    print(f"==== {mode}: 依次运行每个 Session ====")
    run_for_rows(dlg, pane_auto_id=mode)

    # 重置勾选
    print(f"==== {mode} 完成后：取消勾选 {mode} ====")
    checkbox.click_input()
    time.sleep(0.5)

    print(f"🎉 所有 {mode} 已依次运行完毕！")

if __name__ == "__main__":
    main()
