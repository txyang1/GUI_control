from pywinauto import Application, timings
import time


def automate_vigem_sessions(converter_exe, input_folder, output_folder):
    # Start or connect to the application
    app = Application(backend='uia').start(converter_exe)
    dlg = app.window(title_re='ViGEM CCA-Converter.*')
    dlg.wait('visible', timeout=30)

    # === Step 1: Set Input Path and Refresh ===
    # Enter input folder
    input_combo = dlg.child_window(auto_id='inputLocationComboBox', control_type='ComboBox')
    input_edit = input_combo.child_window(control_type='Edit').wrapper_object()
    input_edit.set_edit_text(input_folder)
    # Click the first Open button next to input edit
    dlg.child_window(title='Open', control_type='Button', found_index=0).click_input()
    # Refresh file list
    dlg.child_window(title='Refresh', control_type='Button').click_input()
    timings.wait_until_passes(5, 0.5,
                                lambda: dlg.child_window(title='Sessions', control_type='TabItem').exists())

    # === Step 2: Set Output Path ===
    output_combo = dlg.child_window(auto_id='outputFolderComboBox', control_type='ComboBox')
    output_edit = output_combo.child_window(control_type='Edit').wrapper_object()
    output_edit.set_edit_text(output_folder)
    # Open system folder dialog
    dlg.child_window(auto_id='browseButton', control_type='Button').click_input()
    # Wait for system folder dialog (#32770)
    folder_dlg = app.window(class_name='#32770')
    folder_dlg.wait('visible', timeout=10)
    # Enter path and confirm
    fd_edit = folder_dlg.child_window(control_type='Edit').wrapper_object()
    fd_edit.set_edit_text(output_folder)
    try:
        folder_dlg.child_window(title='确定', control_type='Button').click_input()
    except:
        folder_dlg.child_window(title='Open', control_type='Button').click_input()

    # === Step 3: Process Sessions ===
    # Switch to Sessions tab
    dlg.child_window(title='Sessions', control_type='TabItem').click_input()
    # Wait until the sessions table appears
    timings.wait_until_passes(10, 0.5,
                                lambda: dlg.child_window(control_type='Table').exists())
    table = dlg.child_window(control_type='Table').wrapper_object()

    # Each session row is a Custom element under the table
    session_rows = table.children(control_type='Custom')
    for row in session_rows:
        # Select the session row
        row.click_input()
        time.sleep(0.5)

        # Ensure PCAP checkbox is selected
        pcap_chk = dlg.child_window(title='PCAP', control_type='CheckBox').wrapper_object()
        if pcap_chk.get_toggle_state() == 0:
            pcap_chk.toggle()

        # Click Start button
        start_btn = dlg.child_window(auto_id='startButton', control_type='Button').wrapper_object()
        start_btn.click_input()

        # Wait for progress bar to disappear
        progress = dlg.child_window(auto_id='progressBar', control_type='ProgressBar')
        timings.wait_until_passes(300, 1,
                                    lambda: not progress.exists() or not progress.is_visible())
        # Small pause before next iteration
        time.sleep(1)

    print('All sessions processed.')


if __name__ == '__main__':
    converter_path = r'C:\Path\To\ViGEM_CCA-Converter.exe'
    input_path = r'C:\Data\Input'
    output_path = r'C:\Data\Output'
    automate_vigem_sessions(converter_path, input_path, output_path)
