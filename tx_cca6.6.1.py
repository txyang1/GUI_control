import time
import re
from pywinauto import Application, timings

# 请根据你的实际环境修改这些路径
PCAP_INPUT_FOLDER  = r"C:\Path\To\Pcap\Input"
PCAP_OUTPUT_FOLDER = r"C:\Path\To\Pcap\Output"
MTRE_INPUT_FOLDER  = r"C:\Path\To\Mtre\Input"
MTRE_OUTPUT_FOLDER = r"C:\Path\To\Mtre\Output"

def set_folder(dlg, combo_auto_id: str, path: str):
    combo = dlg.child_window(auto_id=combo_auto_id, control_type="ComboBox")
    combo.child_window(control_type="Edit").set_edit_text(path)
    time.sleep(0.2)
    # 尝试点击同层级的 Open / Browse 按钮
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

    print(f"→ 在 Sessions 中找到 {len(row_names)} 行：{row_names}")

    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button")

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")

        # 刷新并选中行
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")
        cell = next(it for it in all_items if it.element_info.name.strip() == name)
        cell.click_input()
        time.sleep(0.2)

        # 点击 ▶ 启动
        func_start_btn.click_input()
        # 固定等待 5 秒
        time.sleep(5)
        # 再次点击以取消选中
        cell.click_input()
        time.sleep(0.2)

        print(f"  ✓ {pane_auto_id} 已完成 {name}")

def main():
    # 启动并连接
    app = Application(backend="uia").start(r"C:\Path\To\ViGEM CCA-Converter.exe")
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # PCAP 设置并运行
    print("==== PCAP: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", PCAP_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", PCAP_OUTPUT_FOLDER)
    run_for_rows(dlg, pane_auto_id="PCAP")

    # 在运行 MTRE 之前，先点击一次 PCAP 的 ▶ 按钮
    print("==== 切换到 MTRE 前：额外点击一次 PCAP ▶ 按钮 ====")
    dlg.child_window(auto_id="PCAP", control_type="Pane") \
       .child_window(auto_id="startButton", control_type="Button") \
       .click_input()
    time.sleep(0.5)

    # MTRE 设置并运行
    print("==== MTRE: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", MTRE_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", MTRE_OUTPUT_FOLDER)
    run_for_rows(dlg, pane_auto_id="MTRE")

    print("🎉 所有 PCAP 与 MTRE 均已依次运行完毕！")

if __name__ == "__main__":
    main()
