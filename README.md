# 文档错别字/病句自动扫描脚本（Windows 10）

这个脚本可以直接扫描文本或文档，检测：
- 错别字（基于 `pycorrector`）
- 常见病句模式（重复词、重复标点、超长无停顿等）

并给出错误内容与其在全文中的位置（start/end + 行列号）。

## 环境
- Windows 10 台式机（Python 3.10+）

## 安装
```bash
pip install -r requirements.txt
```

## 命令行使用
```bash
python text_proofreader.py 你的文档.txt -o report.tsv
```

支持类型：`txt/md/docx/csv/log`（其中 `.docx` 使用内置解析，不依赖 `python-docx`）

输出示例：
- 终端预览前 30 条问题
- `report.tsv` 输出所有问题

## 图形界面使用
```bash
python text_proofreader.py --gui
```
点击“选择文件并扫描”后，会弹窗显示扫描结果，主窗口表格显示每条问题和位置。

## 性能建议（25 万字 / 30 分钟目标）
1. 优先在本地 SSD 上扫描。  
2. 关闭后台高占用软件。  
3. 首次运行会加载模型，后续扫描更快。  
4. 对于超大文档，建议按章节分批处理并汇总报告。  

> 实际性能受 CPU、内存和文本复杂度影响。

> 若未安装 `pycorrector`，脚本仍可运行，但只执行病句规则检测。
