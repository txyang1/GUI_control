# GUI_control
# ViGEM CCA-Converter 自动化 Runner

这是一个基于 Python 和 pywinauto 的自动化脚本，用于批量执行 ViGEM CCA-Converter GUI 工具中的 **Files** 标签页下的多个文件处理任务。脚本支持按单个或自定义批量大小依次选中文件、一次性运行、再依次取消选中。

## 功能特点

* **自动启动并连接** ViGEM CCA-Converter GUI 应用。
* **批量或逐条** 选中文件并执行 Run 操作。
* **自定义批量大小**，通过命令行参数 `--batch-size` 控制一次运行的文件数。
* **支持 PCAP/MTRE** 两种模式，可分别或同时运行。
* **智能滚动**：确保目标文件行在可视区，避免因列表过长而选不中或取消选不中。

## 环境依赖

* Python 3.7+
* [pywinauto](https://pypi.org/project/pywinauto/) 库

```bash
pip install pywinauto
```

## 目录结构

```plaintext
├── run_converter.py      # 主脚本
├── README.md             # 本说明文档
└── trace_Example_No_Copy_Outside/  # 测试用输入文件夹示例
```

## 配置

在脚本顶部，修改以下常量为你的实际文件夹路径：

```python
PCAP_INPUT_FOLDER  = r"C:\path\to\pcap_input"
PCAP_OUTPUT_FOLDER = r"C:\path\to\pcap_output"
MTRE_INPUT_FOLDER  = r"C:\path\to\mtre_input"
MTRE_OUTPUT_FOLDER = r"C:\path\to\mtre_output"
```

## 使用方法

```bash
python run_converter.py [--mode MODE] [--batch-size N]
```

* `--mode`: 指定运行模式，可选值 `PCAP`, `MTRE`, `ALL`（默认 `ALL`）。
* `--batch-size`: 每批次选中并运行的文件数量，默认 `1`（逐个运行）。

### 示例

* **逐个运行 PCAP 模式**：

  ```bash
  python gui3.1.py --mode PCAP
  ```

* **每 5 个文件为一批运行 MTRE**：

  ```bash
  python Gui4.1.py --mode MTRE --batch-size 5
  ```

* **同时运行 PCAP 和 MTRE，每批 3 个文件**：

  ```bash
  python Gui4.1.py --mode ALL --batch-size 3
  ```

## 注意事项

1. 确保 ViGEM CCA-Converter GUI 路径在脚本中 `Application.start()` 调用处正确。
2. 界面加载和滚动动作依赖窗口响应速度，可根据实际情况适当调整 `time.sleep()` 时长。
3. 如果遇到滚动失效，可检查 UIA 模式和表格控件是否支持 ScrollPattern。

---

版权所有 © 2025
