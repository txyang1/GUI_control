 # **1. 优先尝试 DataItem 自带的滚动方法**
        try:
            cell.scroll_into_view()
            time.sleep(0.2)
        except Exception:
            # **2. 回退：对表格发送 PageDown，直到该行真正可见为止**
            sessions_table.set_focus()
            while cell.element_info.element.CurrentIsOffscreen:
                send_keys("{PGDN}")
                time.sleep(0.1)
            time.sleep(0.2)
