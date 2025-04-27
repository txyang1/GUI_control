from pywinauto import Application, timings

# 1. 连接到已经打开的 Converter 窗口
app = Application(backend="uia").connect(title_re="ViGEM CCA-Converter.*")
dlg = app.window(title_re="ViGEM CCA-Converter.*")

# 2. 切换到 Sessions 页签
tab_ctrl = dlg.child_window(auto_id="tabControl1", control_type="Tab")
sessions_tab = tab_ctrl.child_window(title="Sessions", control_type="TabItem")
sessions_tab.select()

# 等待 Sessions 页面上的表格渲染出来
timings.wait_until_passes(10, 0.5, 
    lambda: dlg.child_window(control_type="Table").exists()
)

# 3. 找到那个 Table（DataGridView）
sessions_table = dlg.child_window(control_type="Table").wrapper_object()

# 4. 遍历每一行（每个 Custom），点击第一列的 DataItem（checkbox 单元格）
for row in sessions_table.children(control_type="Custom"):
    # row.children(control_type="DataItem")[0] 就是该行最左边那一格
    checkbox_cell = row.children(control_type="DataItem")[0]
    checkbox_cell.click_input()
    # （可选）打印行名确认
    folder_name = row.children(control_type="DataItem")[2].window_text()
    print(f"已选中文件夹：{folder_name}")
