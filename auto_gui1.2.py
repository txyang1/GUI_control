from pywinauto import Application, timings
import time

def automate_converter(converter_exe, input_folder, output_folder):
    # Launch or connect to the application
    app = Application(backend="uia").start(converter_exe)
    dlg = app.window(title_re="ViGEM CCA-Converter.*")
    dlg.wait('visible', timeout=30)

    # === Step 1: Set Input Path and Refresh ===
    # Type the folder path directly into the Input Location edit
    input_edit = dlg.child_window(auto_id="inputLocationComboBox", control_type="ComboBox")\
                    .child_window(control_type="Edit").wrapper_object()
    input_edit.set_edit_text(input_folder)
    # Click the Refresh button in the toolbar to load files
    toolbar = dlg.child_window(auto_id="toolStrip1", control_type="ToolBar")
    toolbar.child_window(title="Refresh", control_type="Button").click_input()
    timings.wait_until_passes(5, 0.5, lambda: dlg.child_window(title="Sessions", control_type="TabItem").exists())

    # === Step 2: Set Output Path ===
    output_edit = dlg.child_window(auto_id="outputFolderComboBox", control_type="ComboBox")\
                     .child_window(control_type="Edit").wrapper_object()
    output_edit.set_edit_text(output_folder)
    # Click the Browse button next to output to apply
    dlg.child_window(auto_id="browseButton", control_type="Button").click_input()

    # === Step 3: Iterate Sessions Files & Run PCAP ===
    # Switch to Sessions tab
    dlg.child_window(title="Sessions", control_type="TabItem").click_input()
    timings.wait_until_passes(10, 0.5, lambda: dlg.child_window(control_type="Table").exists())
    table = dlg.child_window(control_type="Table").wrapper_object()

    # Iterate all session rows
    rows = [r for r in table.children(control_type="DataItem")]
    for row in rows:
        row.click_input()
        time.sleep(0.5)

        # Ensure PCAP is selected
        pcap_chk = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
        if pcap_chk.get_toggle_state() == 0:
            pcap_chk.toggle()

        # Start conversion
        start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
        start_btn.click_input()

        # Wait until progress finishes
        progress = dlg.child_window(auto_id="progressBar", control_type="ProgressBar")
        timings.wait_until_passes(300, 1, lambda: not progress.is_visible())

    print("All sessions processed.")

if __name__ == '__main__':
    converter_exe = r"C:\Path\To\ViGEM_CCA-Converter.exe"
    input_folder = r"C:\Data\Input"
    output_folder = r"C:\Data\Output"
    automate_converter(converter_exe, input_folder, output_folder)
