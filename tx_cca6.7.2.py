import time
import re
from pywinauto import Application, timings

# 请根据你的实际环境修改这些路径
PCAP_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\trace_Example_No_Copy_Outside"
PCAP_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output"
MTRE_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\MT_RE"
MTRE_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output2"


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


def click_when_ready(pane, timeout=60, retry_interval=0.5):
    """
    在 pane 下等待 startButton 可见、可用并就绪，然后点击。
    """
    btn_spec = pane.child_window(auto_id="startButton", control_type="Button")
    btn_spec.wait("visible enabled ready", timeout=timeout)
    btn_spec.wrapper_object().click_input()


def wait_for_completion(pane, timeout=600, retry_interval=1):
    """
    在 pane 下等待 startButton 先被禁用（任务开始），再重新启用（任务完成）。
    """
    btn_spec = pane.child_window(auto_id="startButton", control_type="Button")
    timings.wait_until(
        timeout=timeout,
        retry_interval=retry_interval,
        func=lambda: not btn_spec.wrapper_object().is_enabled()
    )
    timings.wait_until(
        timeout=timeout,
        retry_interval=retry_interval,
        func=lambda: btn_spec.wrapper_object().is_enabled()
    )


def run_for_rows(dlg, pane_auto_id: str):
    # 切换到 Sessions 标签页，收集所有行名称
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    all_items = sessions_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip() for itm in all_items if itm.element_info.name and name_pat.match(itm.element_info.name)]

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")

        # 刷新并选中当前行
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")
        cell = next(it for it in all_items if it.element_info.name.strip() == name)
        try:
            cell.scroll_into_view()
        except Exception:
            pass
        cell.click_input()
        time.sleep(0.2)

        # 每次获取最新的 func_pane
        func_pane = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
        # 点击并等待启动
        click_when_ready(func_pane, timeout=60, retry_interval=0.5)
        # 等待任务启动并完成
        try:
            wait_for_completion(func_pane, timeout=600, retry_interval=1)
        except Exception as e:
            print(f"  ⚠️ {pane_auto_id} 在运行 {name} 时发生超时或错误: {e}")
            continue

        # 任务完成后，取消选中当前行
        try:
            cell.click_input()
        except Exception:
            pass
        time.sleep(0.5)

        print(f"  ✓ {pane_auto_id} 已完成 {name}")


def main():
    # 启动并连接主窗口
    app = Application(backend="uia").start(
        r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\15_ViGEM_CCA\CCA Tool\bin\ViGEM.CCA.Converter.Gui.exe"
    )
    dlg = app.window(title_re=r".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # 1. PCAP 输入/输出文件夹设置
    print("==== PCAP: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", PCAP_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", PCAP_OUTPUT_FOLDER)

    # 2. 确认 PCAP 方块已勾选
    print("==== 确认 PCAP 方块已勾选 ====")
    pcap_checkbox = (
        dlg.child_window(auto_id="PCAP", control_type="Pane")
           .child_window(auto_id="activateCheckBox", control_type="CheckBox")
           .wrapper_object()
    )
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input()
        time.sleep(0.2)

    # 3. PCAP 循环执行
    print("==== PCAP: 依次运行每个 Session ====")
    run_for_rows(dlg, pane_auto_id="PCAP")

    # 4. PCAP 完成后，重置勾选状态
    print("==== PCAP 完成后：重置 PCAP 勾选 ====")
    pcap_checkbox.click_input()
    time.sleep(0.5)

    # 5. MTRE 输入/输出文件夹设置
    print("==== MTRE: 设置输入/输出文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", MTRE_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", MTRE_OUTPUT_FOLDER)

    # 6. 确认 MTRE 方块已勾选
    print("==== 确认 MTRE 方块已勾选 ====")
    mtre_checkbox = (
        dlg.child_window(auto_id="MTRE", control_type="Pane")
           .child_window(auto_id="activateCheckBox", control_type="CheckBox")
           .wrapper_object()
    )
    if mtre_checkbox.get_toggle_state() == 0:
        mtre_checkbox.click_input()
        time.sleep(0.2)

    # 7. MTRE 循环执行
    print("==== MTRE: 依次运行每个 Session ====")
    run_for_rows(dlg, pane_auto_id="MTRE")

    # 8. MTRE 完成后，重置勾选状态
    print("==== MTRE 完成后：重置 MTRE 勾选 ====")
    mtre_checkbox.click_input()
    time.sleep(0.5)

    print("🎉 所有 PCAP 与 MTRE 已依次运行完毕！")


if __name__ == "__main__":
    main()
