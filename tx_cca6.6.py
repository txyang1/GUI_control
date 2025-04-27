import time
import re
from pywinauto import Application, timings

# 您需要在这里修改为实际的路径
PCAP_INPUT_FOLDER  = r"C:\Path\To\Pcap\Input"
PCAP_OUTPUT_FOLDER = r"C:\Path\To\Pcap\Output"
MTRE_INPUT_FOLDER  = r"C:\Path\To\Mtre\Input"
MTRE_OUTPUT_FOLDER = r"C:\Path\To\Mtre\Output"

def set_folder(dlg, combo_auto_id: str, path: str):
    """
    通过下拉框的 auto_id 设置文件夹路径，并点击对应的 Open/Browse 按钮（如果存在）。
    """
    combo = dlg.child_window(auto_id=combo_auto_id, control_type="ComboBox")
    # 直接写入编辑框
    combo.child_window(control_type="Edit").set_edit_text(path)
    time.sleep(0.2)
    # 尝试点击同层级的 Open 或 Browse 按钮
    for title in ("Open", "Browse", "openExplorerButton", "browseButton"):
        try:
            btn = combo.parent().child_window(title=title, control_type="Button")
            btn.click_input()
            time.sleep(0.5)
            break
        except Exception:
            continue

def run_for_rows(dlg, pane_auto_id: str):
    """
    在 Sessions 下，依次选中每行：选→启动 pane_auto_id 功能 → 等待 5s → 取消选中
    pane_auto_id 例 "PCAP" 或 "MTRE"
    """
    # 切到 Sessions
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    # 定位表格
    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()

    # 先抓所有行名称
    all_items = sessions_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip()
                 for itm in all_items if itm.element_info.name and name_pat.match(itm.element_info.name)]

    print(f"→ 在 Sessions 中找到 {len(row_names)} 行：{row_names}")

    # 定位功能 Pane 和它的启动按钮
    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button")

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")

        # 刷新并选中对应行
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")
        cell = next(it for it in all_items if it.element_info.name.strip() == name)
        cell.click_input()
        time.sleep(0.2)

        # 点击启动
        func_start_btn.click_input()

        # 固定等待 5 秒
        time.sleep(5)

        # 再次点击同一个 DataItem 取消选中
        cell.click_input()
        time.sleep(0.2)

        print(f"  ✓ {pane_auto_id} 已完成 {name}")

def main():
    # 启动并连接
    app = Application(backend="uia").start(
        r"C:\Path\To\ViGEM CCA-Converter.exe"  # ← 改为实际可执行文件路径
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 1. PCAP：先设置输入/输出文件夹
    print("==== PCAP: 设置文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", PCAP_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", PCAP_OUTPUT_FOLDER)
    # 再运行 PCAP
    run_for_rows(dlg, pane_auto_id="PCAP")

    # 2. MTRE：同样地，设置它的输入/输出文件夹
    print("==== MTRE: 设置文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", MTRE_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", MTRE_OUTPUT_FOLDER)
    # 再运行 MTRE
    run_for_rows(dlg, pane_auto_id="MTRE")

    print("🎉 所有 PCAP 与 MTRE 均已依次运行完毕！")

if __name__ == "__main__":
    main()
