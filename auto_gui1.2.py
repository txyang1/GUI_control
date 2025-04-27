from pywinauto import Application, timings
import time

def automate_converter(converter_exe, input_folder, output_folder):
    # Launch the application
    app = Application(backend="uia").start(converter_exe)
    dlg = app.window(title_re="ViGEM CCA-Converter.*")
    dlg.wait('visible', timeout=30)

    # === Step 1: Set Input Path ===
    input_cb = dlg.child_window(auto_id="inputLocationComboBox", control_type="ComboBox")
    input_edit = input_cb.child_window(control_type="Edit").wrapper_object()
    input_edit.set_edit_text(input_folder)
    # Click the first "Open" button next to input
    open_buttons = dlg.descendants(control_type="Button", title="Open")
    if open_buttons:
        open_buttons[0].click_input()
    else:
        raise RuntimeError("Input Open button not found")

    # === Step 2: Set Output Path ===
    output_cb = dlg.child_window(auto_id="outputFolderComboBox", control_type="ComboBox")
    output_edit = output_cb.child_window(control_type="Edit").wrapper_object()
    output_edit.set_edit_text(output_folder)
    # Click "Browse" button for output
    browse_btn = dlg.child_window(auto_id="browseButton", control_type="Button").wrapper_object()
    browse_btn.click_input()

    # === Step 3: Select Sessions Tab ===
    sessions_tab = dlg.child_window(title="Sessions", control_type="TabItem").wrapper_object()
    sessions_tab.click_input()
    # Wait until the Sessions pane loads
    timings.wait_until_passes(10, 0.5, lambda: dlg.child_window(auto_id="1771800", control_type="Pane").exists())
    sessions_page = dlg.child_window(auto_id="1771800", control_type="Pane")

    # Wait for the sessions table to appear
    timings.wait_until_passes(10, 0.5, lambda: sessions_page.child_window(control_type="Table").exists())
    table = sessions_page.child_window(control_type="Table").wrapper_object()

    # === Step 4: Iterate over session rows ===
    rows = table.children(control_type="Custom")  # each row is a Custom control
    print(f"Found {len(rows)} session rows")
    for row in rows:
        row.click_input()
        time.sleep(0.5)

        # Ensure PCAP is checked
        pcap_chk = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
        if pcap_chk.get_toggle_state() == 0:
            pcap_chk.toggle()

        # Click Start
        start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
        start_btn.click_input()

        # Wait for process to finish
        progress = dlg.child_window(auto_id="progressBar", control_type="ProgressBar")
        timings.wait_until_passes(300, 1, lambda: not progress.is_visible())

    print("All sessions processed.")


if __name__ == '__main__':
    converter_exe = r"C:\Path\To\ViGEM_CCA-Converter.exe"
    input_folder = r"C:\Data\Input"
    output_folder = r"C:\Data\Output"
    automate_converter(converter_exe, input_folder, output_folder)
