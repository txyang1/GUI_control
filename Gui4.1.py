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
    在 Files 页签中，按 batch_size 分批，每批内部依次单独选中行、点击 ▶、等待完成、再清除选中。
    """
    # 切换到 Files 标签页
    dlg.child_window(auto_id="tabControl1", control_type="Tab") \
       .child_window(title="Files", control_type="TabItem").select()
    time.sleep(0.5)

    # 收集所有行名
    table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()
    all_items = table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [
        itm.element_info.name.strip()
        for itm in all_items
        if itm.element_info.name and name_pat.match(itm.element_info.name)
    ]

    # 启动按钮引用
    func_start_btn = dlg.child_window(auto_id=pane_auto_id, control_type="Pane") \
                        .child_window(auto_id="startButton", control_type="Button")

    total = len(row_names)
    for batch_start in range(0, total, batch_size):
        batch = row_names[batch_start:batch_start + batch_size]
        print(f"[Files {batch_start+1}-{min(batch_start+batch_size, total)}/{total}] 运行 {pane_auto_id} → {batch}")

        # 对批次内每个文件依次单独运行
        for name in batch:
            # 重新切回 Files tab & 定位 table
            dlg.child_window(auto_id="tabControl1", control_type="Tab") \
               .child_window(title="Files", control_type="TabItem").select()
            time.sleep(0.2)
            table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()

            # 找到这一行
            itm = next(
                it for it in table.descendants(control_type="DataItem")
                if it.element_info.name.strip() == name
            )

            # 滚动至可见
            try:
                itm.scroll_into_view()
                time.sleep(0.1)
            except Exception:
                table.set_focus()
                while itm.element_info.element.CurrentIsOffscreen:
                    send_keys("{PGDN}")
                    time.sleep(0.1)
            time.sleep(0.1)

            # 选中并运行
            itm.click_input()
            time.sleep(0.2)
            func_start_btn.click_input()
            timings.wait_until(
                timeout=600,
                retry_interval=1,
                func=lambda: func_start_btn.is_enabled()
            )
            time.sleep(0.2)

            # 取消选中
            itm.click_input()
            time.sleep(0.1)

            print(f"  ✓ 已完成 {name}")

        print(f"  —— 完成 batch {batch_start//batch_size + 1} ——\n")

def run_mode(dlg, mode: str, batch_size: int):
    """
    对 PCAP 或 MTRE 模式，设置文件夹、激活、在 Files 页签上分批运行，再重置复选框。
    """
    folders = FOLDERS[mode]
    print(f"==== {mode}: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", folders["input"])
    set_folder(dlg, "outputFolderComboBox", folders["output"])

    # 勾选模式复选框
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
        help="Files 标签下每隔多少个文件（依次）运行一次（默认 1，即逐个运行）"
    )
    args = parser.parse_args()

    # 启动并连接主窗口
    app = Application(backend="uia").start(
        r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\15_ViGEM_CCA\CCA Tool\bin\ViGEM.CCA.Converter.Gui.exe"
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 根据参数选择模式
    modes = [args.mode] if args.mode != 'ALL' else ['PCAP', 'MTRE']
    for m in modes:
        run_mode(dlg, m, args.batch_size)

    print("🎉 所有选定模式的 Files 批量运行完毕！")

if __name__ == "__main__":
    main()


def run_for_rows(dlg, pane_auto_id: str, batch_size: int = 1):
    """
    在 Files 页签中，将所有行按 batch_size 分批。
    对每批里的每一行，都做：
      - 选中
      - 点击 Run
      - 等待完成
      - 取消选中
    打印出“批次开始”和“批次完成”信息。
    """
    # 切到 Files 标签页
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
    # 分批处理
    for batch_start in range(0, total, batch_size):
        batch = row_names[batch_start:batch_start + batch_size]
        print(f"\n===== 开始批次 {batch_start//batch_size + 1}（{batch_start+1}-{min(batch_start+batch_size, total)}/{total}）: {batch} =====")

        for name in batch:
            # 切回 Files tab 并定位表格
            dlg.child_window(auto_id="tabControl1", control_type="Tab") \
               .child_window(title="Files", control_type="TabItem").select()
            time.sleep(0.2)
            table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()

            # 找到该行
            itm = next(it for it in table.descendants(control_type="DataItem")
                       if it.element_info.name.strip() == name)

            # 滚动到可见
            try:
                itm.scroll_into_view()
            except Exception:
                table.set_focus()
                while itm.element_info.element.CurrentIsOffscreen:
                    send_keys("{PGDN}")
            time.sleep(0.1)

            # 点击选中
            itm.click_input()
            time.sleep(0.1)

            # 点击 Run 并等待完成
            run_btn.click_input()
            timings.wait_until(
                timeout=600,
                retry_interval=1,
                func=lambda: run_btn.is_enabled()
            )
            time.sleep(0.1)

            # 取消选中
            itm.click_input()
            time.sleep(0.1)
            print(f"    ✓ 已完成 {name}")

        print(f"===== 完成批次 {batch_start//batch_size + 1} =====\n")
