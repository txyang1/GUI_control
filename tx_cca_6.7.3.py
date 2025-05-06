import time
import re
import os
from pywinauto import Application, timings, findwindows

# 全局应用对象
APP = None

# 请根据你的实际环境修改这些路径
PCAP_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\trace_Example_No_Copy_Outside"
PCAP_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output"
MTRE_INPUT_FOLDER  = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\MT_RE"
MTRE_OUTPUT_FOLDER = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\output2"
ERROR_LOG_FILE     = r"C:\Users\qxz5y3m\Desktop\Transfer_Tool_for_trace\mtre_errors.log"


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
    等待并点击 startButton
    """
    btn_spec = pane.child_window(auto_id="startButton", control_type="Button")
    btn_spec.wait("visible enabled ready", timeout=timeout)
    btn_spec.wrapper_object().click_input()


def wait_for_completion(pane, timeout=600, retry_interval=1):
    """
    等待 startButton 先 disabled 再 enabled
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


def validate_mtre_inputs(folder):
    """
    简单校验 MTRE 输入目录是否非空
    """
    files = os.listdir(folder)
    if not files:
        raise RuntimeError(f"MTRE 输入目录 {folder} 没有文件！")
    return True


def check_and_close_error():
    """
    查找 MTRE 弹出的错误对话框，关闭并返回错误消息
    """
    try:
        # 匹配标题含 Error 或 Warning 的窗口
        dlg = APP.window(found_index=0, title_re=r".*(Error|Warning).*", control_type="Window")
        if dlg.exists(timeout=1):
            msg = dlg.window(control_type="Text").window_text()
            dlg.close()
            return msg
    except Exception:
        pass
    return None


def run_for_rows(dlg, pane_auto_id: str, is_mtre=False):
    if is_mtre:
        try:
            validate_mtre_inputs(MTRE_INPUT_FOLDER)
        except Exception as e:
            print(f"‼️ MTRE 输入校验失败: {e}")
            return

    # 收集 Sessions
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)
    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    items = sessions_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip() for itm in items if itm.element_info.name and name_pat.match(itm.element_info.name)]

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")
        # 选中行
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        items = sessions_table.descendants(control_type="DataItem")
        cell = next(it for it in items if it.element_info.name.strip() == name)
        try: cell.scroll_into_view()
        except: pass
        cell.click_input(); time.sleep(0.2)

        func_pane = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
        # 点击启动
        click_when_ready(func_pane)
        # 先检查错误弹窗
        err = check_and_close_error()
        if err:
            print(f"  ‼️ {pane_auto_id} 在 {name} 上弹错: {err}, 已跳过")
            with open(ERROR_LOG_FILE, "a") as f:
                f.write(f"{time.ctime()}: {pane_auto_id} {name} 错误: {err}\n")
            try: cell.click_input()
            except: pass
            continue
        # 等待任务
        try:
            wait_for_completion(func_pane)
        except Exception as e:
            print(f"  ⚠️ {pane_auto_id} 在 {name} 超时: {e}, 已跳过")
            with open(ERROR_LOG_FILE, "a") as f:
                f.write(f"{time.ctime()}: {pane_auto_id} {name} 超时: {e}\n")
            try: cell.click_input()
            except: pass
            continue

        # 完成后取消选中
        try: cell.click_input()
        except: pass
        time.sleep(0.5)
        print(f"  ✓ {pane_auto_id} 已完成 {name}")


def main():
    global APP
    APP = Application(backend="uia").start(
        r"C:\Users\qxz5m\Desktop\Transfer_Tool_for_trace\15_ViGEM_CCA\CCA Tool\bin\ViGEM.CCA.Converter.Gui.exe"
    )
    dlg = APP.window(title_re=r".*ViGEM CCA-Converter.*")
    dlg.wait("visible enabled ready", timeout=30)
    dlg.set_focus()

    # PCAP
    print("==== PCAP: 设置文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", PCAP_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", PCAP_OUTPUT_FOLDER)
    pcap_cb = dlg.child_window(auto_id="PCAP", control_type="Pane").child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if pcap_cb.get_toggle_state()==0: pcap_cb.click_input(); time.sleep(0.2)
    print("==== PCAP: 执行 Sessions ====")
    run_for_rows(dlg, "PCAP")
    print("==== PCAP: 重置勾选 ====")
    pcap_cb.click_input(); time.sleep(0.5)

    # MTRE
    print("==== MTRE: 设置文件夹 ====")
    set_folder(dlg, "inputLocationComboBox", MTRE_INPUT_FOLDER)
    set_folder(dlg, "outputFolderComboBox", MTRE_OUTPUT_FOLDER)
    mtre_cb = dlg.child_window(auto_id="MTRE", control_type="Pane").child_window(auto_id="activateCheckBox", control_type="CheckBox").wrapper_object()
    if mtre_cb.get_toggle_state()==0: mtre_cb.click_input(); time.sleep(0.2)
    print("==== MTRE: 执行 Sessions ====")
    run_for_rows(dlg, "MTRE", is_mtre=True)
    print("==== MTRE: 重置勾选 ====")
    mtre_cb.click_input(); time.sleep(0.5)

    print("🎉 所有任务已完成！")

if __name__ == "__main__":
    main()
