from pywinauto import Application, timings
import time


def automate_converter(converter_exe, input_folder, output_folder):
    # Launch or connect to the application
    app = Application(backend="uia").start(converter_exe)
    dlg = app.window(title_re="ViGEM CCA-Converter.*")
    dlg.wait('visible', timeout=30)

    # === Step 1: Set Input Path ===
    input_cb = dlg.child_window(auto_id="inputLocationComboBox", control_type="ComboBox").wrapper_object()
    input_cb.set_text(input_folder)
    dlg.child_window(title="Open", control_type="Button", found_index=0).click()

    # === Step 2: Set Output Path ===
    output_cb = dlg.child_window(auto_id="outputFolderComboBox", control_type="ComboBox").wrapper_object()
    output_cb.set_text(output_folder)
    # There are multiple "Open" buttons; this one is the 5th in the tree under the output pane
    dlg.child_window(title="Open", control_type="Button", found_index=4).click()

    # === Step 3: Iterate Sessions Files & Run PCAP ===
    # Select "Sessions" tab
    dlg.child_window(title="Sessions", control_type="TabItem").select()
    # Give time for table to populate
    timings.wait_until_passes(5, 0.5, lambda: dlg.child_window(control_type="Table").exists())
    table = dlg.child_window(control_type="Table").wrapper_object()

    # Iterate each row in the sessions table
    rows = table.children(control_type="DataItem")
    for row in rows:
        # Click the session row to select it
        row.click_input()
        time.sleep(0.5)

        # Ensure PCAP format is checked
        pcap_chk = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
        if not pcap_chk.get_toggle_state():
            pcap_chk.toggle()

        # Click the "Start" button
        start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
        start_btn.click_input()

        # Wait for progress to complete (progress bar disappears or button re-enables)
        progress = dlg.child_window(auto_id="progressBar", control_type="ProgressBar")
        # Wait until ProgressBar is no longer visible
        timings.wait_until_passes(300, 1, lambda: not progress.is_visible())

        # Optional: uncheck PCAP to reset state
        # pcap_chk.toggle()

    print("All sessions processed.")


if __name__ == '__main__':
    converter_exe = r"C:\Path\To\ViGEM_CCA-Converter.exe"
    input_folder = r"C:\Data\Input"
    output_folder = r"C:\Data\Output"
    automate_converter(converter_exe, input_folder, output_folder)
