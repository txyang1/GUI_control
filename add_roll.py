import time
import re
from pywinauto import Application, timings
from pywinauto.uia_defines import ScrollAmount

def wait_for_finish(btn, timeout=60, retry_interval=0.5):
    timings.WaitUntil(timeout, retry_interval, lambda: btn.is_enabled())

def run_for_rows(dlg, pane_auto_id: str):
    # 切换到 Sessions 选项卡
    tab = dlg.child_window(auto_id="tabControl1", control_type="Tab").wrapper_object()
    tab.child_window(title="Sessions", control_type="TabItem").select()
    time.sleep(0.5)

    # 先拿一次不会变的引用
    func_pane      = dlg.child_window(auto_id=pane_auto_id, control_type="Pane")
    func_start_btn = func_pane.child_window(auto_id="startButton", control_type="Button").wrapper_object()

    # 获取所有行名称
    sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
    all_items = sessions_table.descendants(control_type="DataItem")
    name_pat = re.compile(r"^\s*Row \d+$")
    row_names = [itm.element_info.name.strip() for itm in all_items
                 if itm.element_info.name and name_pat.match(itm.element_info.name)]

    for idx, name in enumerate(row_names, start=1):
        print(f"[{idx}/{len(row_names)}] 运行 {pane_auto_id} → {name}")

        # 重新定位表格，并找到这一行
        tab.child_window(title="Sessions", control_type="TabItem").select()
        time.sleep(0.2)
        sessions_table = dlg.child_window(auto_id="SessionView", control_type="Table").wrapper_object()
        cell = next(it for it in sessions_table.descendants(control_type="DataItem")
                    if it.element_info.name.strip() == name)

        # —— 关键：用 ScrollPattern 翻页，直到目标行不再离屏 —— 
        scroll_pattern = sessions_table.iface_scroll
        # 当控件“离屏”时不停向下滚一页
        while cell.element_info.element.CurrentIsOffscreen:
            scroll_pattern.Scroll(ScrollAmount.NoAmount, ScrollAmount.LargeIncrement)
            time.sleep(0.1)
        # 稍微等一下，确保滚动完成
        time.sleep(0.2)

        # 点击启动
        cell.click_input()
        time.sleep(0.2)
        func_start_btn.click_input()
        wait_for_finish(func_start_btn, timeout=60, retry_interval=0.5)
        # 取消选中
        cell.click_input()
        time.sleep(0.2)

        print(f"  ✓ {pane_auto_id} 已完成 {name}")
