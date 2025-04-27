from pywinauto import Application, timings
import time

def automate_converter(converter_exe, input_folder, output_folder):
    # 启动或连接应用
    app = Application(backend="uia").start(converter_exe)
    dlg = app.window(title_re="ViGEM CCA-Converter.*")
    dlg.wait('visible', timeout=30)

    # === 第一步：设置输入路径 ===
    input_edit = (
        dlg.child_window(auto_id="inputLocationComboBox", control_type="ComboBox")
           .child_window(control_type="Edit").wrapper_object()
    )
    input_edit.set_edit_text(input_folder)
    # 点击第一个 "Open" 按钮以加载输入目录
    dlg.child_window(title="Open", control_type="Button", found_index=0).click_input()

    # 点击工具栏的 Refresh 按钮，确保会话列表更新
    toolbar = dlg.child_window(auto_id="toolStrip1", control_type="ToolBar").wrapper_object()
    toolbar.child_window(title="Refresh", control_type="Button").click_input()
    timings.wait_until_passes(5, 0.5,
        lambda: dlg.child_window(title="Sessions", control_type="TabItem").exists()
    )

    # === 第二步：设置输出路径 ===
    output_edit = (
        dlg.child_window(auto_id="outputFolderComboBox", control_type="ComboBox")
           .child_window(control_type="Edit").wrapper_object()
    )
    output_edit.set_edit_text(output_folder)
    # 点击输出区域内的 Open 按钮以应用路径
    output_panel = dlg.child_window(auto_id="panel4", control_type="Pane")
    output_panel.child_window(title="Open", control_type="Button").click_input()

    # === 第三步：切换到 Sessions，遍历文件并运行 PCAP ===
    dlg.child_window(title="Sessions", control_type="TabItem").click_input()
    timings.wait_until_passes(10, 0.5,
        lambda: dlg.child_window(control_type="Table").exists()
    )
    table = dlg.child_window(control_type="Table").wrapper_object()

    # 遍历所有会话行
    rows = table.children(control_type="DataItem")
    for row in rows:
        row.click_input()
        time.sleep(0.5)

        # 勾选 PCAP 格式
        pcap_chk = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
        if pcap_chk.get_toggle_state() == 0:
            pcap_chk.toggle()

        # 点击 Start 按钮
        start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
        start_btn.click_input()

        # 等待进度完成
        progress = dlg.child_window(auto_id="progressBar", control_type="ProgressBar")
        timings.wait_until_passes(300, 1,
            lambda: not progress.is_visible()
        )

    print("All sessions processed.")

if __name__ == '__main__':
    converter_exe = r"C:\Path\To\ViGEM_CCA-Converter.exe"
    input_folder = r"C:\Data\Input"
    output_folder = r"C:\Data\Output"
    automate_converter(converter_exe, input_folder, output_folder)
