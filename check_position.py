from pywinauto import Application, timings

# 1. 连接到已经打开的 Converter 窗口
app = Application(backend="uia").connect(title_re="ViGEM CCA-Converter.*")
dlg = app.window(title_re="ViGEM CCA-Converter.*")

# 2. 切换到 Sessions 页签
tab = dlg.child_window(auto_id="tabControl1", control_type="Tab")
tab.child_window(title="Sessions", control_type="TabItem").select()

# 等待表格出来
timings.wait_until_passes(10, 0.5,
    lambda: dlg.child_window(control_type="Table").exists()
)

# 3. 拿到 Table
table = dlg.child_window(control_type="Table").wrapper_object()

# 4. 遍历“行容器”（每个 Custom），然后用 descendants() 找到真正的 DataItem 列表
for row in table.children(control_type="Custom"):
    # 把这一行所有深层次的 DataItem 都搜出来
    cells = row.descendants(control_type="DataItem")
    if not cells:
        # 如果实在没找到，跳过
        continue

    # cells[0] 就是这一行最左边的单元格 (checkbox Cell)
    checkbox_cell = cells[0]
    checkbox_cell.click_input()
    # （可选）打印一下，确认到底点了哪一行
    print("选中了：", cells[2].window_text())  # 第 2 列通常是文件夹名
