from pywinauto import Application, timings
import time

def automate_converter(converter_exe, input_folder, output_folder):
    # Launch or connect to the application
    app = Application(backend="uia").start(converter_exe)
    dlg = app.window(title_re="ViGEM CCA-Converter.*")
    dlg.wait('visible', timeout=30)

    # === Step 1: Set Input Path ===
    input_cb = dlg.child_window(auto_id="inputLocationComboBox", control_type="ComboBox")
    # Get the Edit control inside the ComboBox and set text
    input_edit = input_cb.child_window(control_type="Edit").wrapper_object()
    input_edit.set_edit_text(input_folder)
    dlg.child_window(title="Open", control_type="Button", found_index=0).click_input()

    # === Step 2: Set Output Path ===
    output_cb = dlg.child_window(auto_id="outputFolderComboBox", control_type="ComboBox")
    output_edit = output_cb.child_window(control_type="Edit").wrapper_object()
    output_edit.set_edit_text(output_folder)
    dlg.child_window(title="Open", control_type="Button", found_index=4).click_input()

    # === Step 3: Iterate Sessions Files & Run PCAP ===
    dlg.child_window(title="Sessions", control_type="TabItem").select()
    timings.wait_until_passes(5, 0.5, lambda: dlg.child_window(control_type="Table").exists())
    table = dlg.child_window(control_type="Table").wrapper_object()

    rows = [r for r in table.children(control_type="DataItem") if r.element_info.control_type == "DataItem"]
    for row in rows:
        row.click_input()
        time.sleep(0.5)

        # Ensure PCAP format is checked
        pcap_chk = dlg.child_window(title="PCAP", control_type="CheckBox").wrapper_object()
        if not pcap_chk.get_toggle_state():
            pcap_chk.toggle()

        # Click the "Start" button
        start_btn = dlg.child_window(auto_id="startButton", control_type="Button").wrapper_object()
        start_btn.click_input()

        # Wait for progress to complete
        progress = dlg.child_window(auto_id="progressBar", control_type="ProgressBar")
        timings.wait_until_passes(300, 1, lambda: not progress.is_visible())

    print("All sessions processed.")


if __name__ == '__main__':
    converter_exe = r"C:\Path\To\ViGEM_CCA-Converter.exe"
    input_folder = r"C:\Data\Input"
    output_folder = r"C:\Data\Output"
    automate_converter(converter_exe, input_folder, output_folder)
