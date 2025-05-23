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
    2. 否则获取 table 和 item 的屏幕坐标，
       如果 item 在 table 上方用 PageUp， 在下方用 PageDown，
       循环直到完全可见为止。
    """
    try:
        item.scroll_into_view()
        return
    except Exception:
        pass

    # 拿到可视表格和目标行的屏幕坐标
    tbl_rect = table.rectangle()
    itm_rect = item.rectangle()

    # 如果已经完全在可视区内，直接返回
    if itm_rect.top >= tbl_rect.top and itm_rect.bottom <= tbl_rect.bottom:
        return

    # 循环翻页，直到可见
    while not (itm_rect.top >= tbl_rect.top and itm_rect.bottom <= tbl_rect.bottom):
        if itm_rect.top < tbl_rect.top:
            send_keys("{PGUP}")
        else:
            send_keys("{PGDN}")
        time.sleep(0.1)
        itm_rect = item.rectangle()
        tbl_rect = table.rectangle()

def run_for_rows(dlg, pane_auto_id: str, batch_size: int = 1):
    """
    在 Files 页签中，按 batch_size 分批：
      1. 依次点击批次内每一行以选中它
      2. 批次内所有行点击完毕后，仅点击一次 Run
      3. 等待完成
      4. 依次取消选中批次内每一行
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
        print(f"\n===== 开始批次 {batch_start//batch_size + 1} "
              f"({batch_start+1}-{min(batch_start+batch_size, total)}/{total}): {batch} =====")

        # 1. 依次点击选中本批次每一行
        for name in batch:
            # 切回 Files tab 并重新定位表格
            dlg.child_window(auto_id="tabControl1", control_type="Tab") \
               .child_window(title="Files", control_type="TabItem").select()
            time.sleep(0.2)
            table = dlg.child_window(auto_id="FileView", control_type="Table").wrapper_object()

            # 找到对应行
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
            print(f"    已选中 {name}")

        # 2. 批次内所有行选中后，仅此一次点击 Run
        print("    ▶ 点击 Run")
        run_btn.click_input()
        timings.wait_until(
            timeout=600,
            retry_interval=1,
            func=lambda: run_btn.is_enabled()
        )
        time.sleep(0.2)
        print(f"  ✓ 批次 {batch_start//batch_size + 1} 完成")

        # 3. 依次取消选中本批次每一行
        for name in batch:
            

            # 找到对应行
            itm = next(it for it in table.descendants(control_type="DataItem")
                       if it.element_info.name.strip() == name)
           

            itm.click_input()
            time.sleep(0.05)
        print(f"===== 结束批次 {batch_start//batch_size + 1} =====\n")


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
