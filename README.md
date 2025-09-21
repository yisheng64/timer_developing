# 任务计时器 - 增强版

这是一个改进的任务计时器，专门解决了对文档文件（docx, xlsx, pdf等）监控的问题。

## 主要改进

### 1. 多策略文件监控
- **Windows文件锁定检测**: 使用Windows API检测文件是否被其他进程锁定
- **进程访问检测**: 通过psutil检查哪些进程正在访问目标文件
- **应用程序运行检测**: 检测相关应用程序（如Word、Excel、PDF阅读器）是否在运行

### 2. 增强的应用程序识别
- 支持更多办公软件：WPS、LibreOffice、OpenOffice
- 更准确的进程名匹配
- 支持多种PDF阅读器：Adobe Reader、Foxit、Chrome、Edge等

### 3. 改进的检测逻辑
- 综合使用多种检测方法，提高准确性
- 特别针对文档文件优化
- 更好的错误处理和异常捕获

## 安装依赖

运行以下命令安装所需依赖：

```bash
python install_dependencies.py
```

或者手动安装：

```bash
pip install psutil pywin32
```

## 使用方法

1. 运行程序：
   ```bash
   python test1.py
   ```

2. 点击"添加任务文件"选择要监控的文件

3. 打开对应的应用程序（如Word、Excel、PDF阅读器）打开该文件

4. 程序会自动开始计时，关闭文件时记录耗时

## 支持的文件类型

- **文档**: .doc, .docx
- **表格**: .xls, .xlsx  
- **演示文稿**: .ppt, .pptx
- **PDF**: .pdf
- **文本文件**: .txt, .py, .js, .html, .css
- **可执行文件**: .exe

## 支持的应用程序

- Microsoft Office (Word, Excel, PowerPoint)
- WPS Office
- LibreOffice
- Adobe Acrobat Reader
- Foxit Reader
- Chrome/Edge (PDF查看)
- 各种文本编辑器

## 日志记录

程序会自动将计时记录保存到 `task_time_log.csv` 文件中，包含：
- 任务文件路径
- 开始时间
- 结束时间
- 耗时（秒）

## 注意事项

1. 确保安装了pywin32库以获得最佳监控效果
2. 某些应用程序可能需要管理员权限才能被检测到
3. 如果监控不准确，可以尝试以管理员身份运行程序


#版本开发规划
##1 实现对exe文件的监控，进行简单的界面布局设置
##2 实现对word、pdf、excel、ppt的实时监控
##3 实现多线程监控
##4 实现历史记录存储，并在检测到启动相同软件时询问是否接着历史记录继续计时
##5 界面优化，包括开始计时按钮，历史计时展示，历史记录删除等功能