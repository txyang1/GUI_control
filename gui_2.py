import time
import re
from pywinauto import Application, timings
from pywinauto.findwindows import ElementNotFoundError  # 捕获控件未找到异常

# 请根据你的实际环境修改这些路径
PCAP_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\trace_Example_No_Copy_Outside"
PCAP_OUTPUT_FOLDER = r"C:\Users\qxz5m\Desktop\Transfer_Tool_for_trace\output"
MTRE_INPUT_FOLDER  = r"C:\Users\qxz5m\Desktop\Transfer_Tool_for_trace\MT_RE"
MTRE_OUTPUT_FOLDER = r"C:\Users\qxz5m\Desktop\Transfer_Tool_for_trace\output2"

# 方法1: 动态等待按钮再次可用

def wait_for_finish(btn, timeout=60, retry_interval=0.5):
    """
    等待按钮再次可用，表示当前任务已完成
    """
    timings.WaitUntil(
        timeout,
        retry_interval,
        lambda: btn.is_enabled()
    )

# 方法2: 监听 Session 行状态（示例中假设完成后名称中会出现“完成”关键词）

def wait_for_session_complete(dlg, session_name, timeout=60, retry_interval=0.5):
    """
    等待指定 Session 行的状态变化，完成后名称包含“完成"
    """
    end_time = time.time() + timeout
    while time.time() < end_time:
        tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)

        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")
        for itm in all_items:
            name = itm.element_info.name.strip() if itm.element_info.name else ''
            if name == session_name and '完成' in name:
                return
        time.sleep(retry_interval)
    raise RuntimeError(f"Session '{session_name}' 未在 {timeout}s 内完成")


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
    # 缓存控件引用，减少循环内查找开销
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button").wrapper_object()

    # 收集所有 Session 行名称
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    all_items = sessions_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^Row \d+$")
    row_names = [itm.element_info.name.strip()
                 for itm in all_items if itm.element_info.name and name_pat.match(itm.element_info.name)]

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")

        # 切换到 Sessions 选项卡并刷新列表
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        all_items = sessions_table.descendants(control_type="DataItem")

        # 安全获取当前行，防止找不到时抛错
        try:
            cell = next(it for it in all_items if it.element_info.name.strip() == name)
        except StopIteration:
            print(f"  ⚠ 未找到 Session 行 '{name}'，跳过")
            continue

        cell.click_input()
        time.sleep(0.2)

        # 点击 ▶ 启动
        func_start_btn.click_input()
        # 方法1: 等待按钮重新可用
        wait_for_finish(func_start_btn)
        # 方法2: 或者监听行状态
        # wait_for_session_complete(dlg, name)

        # 取消选中当前行
        cell.click_input()
        time.sleep(0.2)

        print(f"  ✓ {pane_auto_id} 已完成 {name}")


def main():
    app = Application(backend="uia").start(
        r"C:\Users\qxz5m\Desktop\Transfer_Tool_for_trace\15_ViGEM_CCA\CCA Tool\bin\ViGEM.CCA.Converter.Gui.exe"
    )
    dlg = app.window(title_re=".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # PCAP 部分
    set_folder(dlg, "inputLocationComboBox", PCAP_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", PCAP_OUTPUT_FOLDER)
    pcap_checkbox = dlg.child_window(auto_id="PCAP", control_type="Pane") \
                     .child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if pcap_checkbox.get_toggle_state() == 0:
        pcap_checkbox.click_input(); time.sleep(0.2)
    print("==== PCAP: 依次运行每个 Session ====")
    run_for_rows(dlg, pane_auto_id="PCAP")
    pcap_checkbox.click_input(); time.sleep(0.5)

    # MTRE 部分
    set_folder(dlg, "inputLocationComboBox", MTRE_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", MTRE_OUTPUT_FOLDER)
    mtre_checkbox = dlg.child_window(auto_id="MTRE", control_type="Pane") \
                     .child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if mtre_checkbox.get_toggle_state() == 0:
        mtre_checkbox.click_input(); time.sleep(0.2)
    print("==== MTRE: 依次运行每个 Session ====")
    run_for_rows(dlg, pane_auto_id="MTRE")
    mtre_checkbox.click_input(); time.sleep(0.5)

    print("🎉 所有 PCAP 与 MTRE 已依次运行完毕！")

if __name__ == "__main__":
    main()
