"""
并行任务监控器（修复 docx/pdf/xlsx 无法启动监控的问题）
依赖：psutil, pywin32
安装：pip install psutil pywin32
运行：python task_timer_parallel.py
"""

import os
import time
import psutil
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import csv
from datetime import datetime
import subprocess

# Windows-specific imports
try:
    import win32file
    import win32con
    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False
    print("警告: pywin32 未安装，某些文件监控功能可能不可用")

# 扩展名到常见打开程序关键字的兜底映射（仅作最后的辅助判断）
FALLBACK_APP_KEYWORDS = {
    '.doc': ['winword', 'word', 'wps', 'libreoffice', 'openoffice'],
    '.docx': ['winword', 'word', 'wps', 'libreoffice', 'openoffice'],
    '.xls': ['excel', 'wps', 'libreoffice', 'openoffice'],
    '.xlsx': ['excel', 'wps', 'libreoffice', 'openoffice'],
    '.ppt': ['powerpnt', 'powerpoint', 'wps', 'libreoffice', 'openoffice'],
    '.pptx': ['powerpnt', 'powerpoint', 'wps', 'libreoffice', 'openoffice'],
    '.pdf': ['acrord', 'acrobat', 'foxit', 'sumatra', 'evince', 'preview', 'chrome', 'msedge', 'edge', 'firefox', 'wps', 'libreoffice'],
    '.txt': ['notepad', 'notepad++', 'gedit', 'sublime', 'code', 'vscode', 'atom'],
    '.py': ['python', 'pycharm', 'code', 'vscode', 'spyder', 'jupyter'],
    '.exe': [],  # 可执行文件通常由自身进程打开，后面通过 cmdline/open_files 能检测到
}

# 更详细的进程名匹配规则
PROCESS_PATTERNS = {
    'word': ['winword.exe', 'wps.exe', 'soffice.exe', 'libreoffice.exe'],
    'excel': ['excel.exe', 'wps.exe', 'soffice.exe', 'libreoffice.exe'],
    'powerpoint': ['powerpnt.exe', 'wps.exe', 'soffice.exe', 'libreoffice.exe'],
    'pdf': ['acrord32.exe', 'acrobat.exe', 'foxitreader.exe', 'sumatra.exe', 'chrome.exe', 'msedge.exe', 'firefox.exe'],
    'text': ['notepad.exe', 'notepad++.exe', 'code.exe', 'sublime_text.exe', 'atom.exe'],
}

def normalize_path(p):
    """标准化路径以便比较（大小写和规范化）"""
    try:
        return os.path.normcase(os.path.abspath(p))
    except Exception:
        return p.lower()

def is_file_locked_windows(file_path):
    """
    使用Windows API检查文件是否被锁定（被其他进程打开）
    返回 True 如果文件被锁定，False 如果文件未被锁定
    """
    if not WINDOWS_AVAILABLE:
        return False
    
    try:
        # 尝试以独占模式打开文件
        handle = win32file.CreateFile(
            file_path,
            win32con.GENERIC_READ,
            0,  # 不共享
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None
        )
        win32file.CloseHandle(handle)
        return False  # 文件未被锁定
    except Exception:
        return True  # 文件被锁定或无法访问

def get_file_handles_windows(file_path):
    """
    获取正在使用指定文件的所有进程句柄（Windows）
    返回进程ID列表
    """
    if not WINDOWS_AVAILABLE:
        return []
    
    try:
        # 使用handle.exe工具获取文件句柄信息
        result = subprocess.run(
            ['handle.exe', file_path],
            capture_output=True,
            text=True,
            timeout=5,
            check=False
        )
        
        if result.returncode == 0:
            pids = []
            for line in result.stdout.split('\n'):
                if 'pid:' in line.lower():
                    try:
                        pid = int(line.split('pid:')[1].split()[0])
                        pids.append(pid)
                    except (ValueError, IndexError):
                        continue
            return pids
    except Exception:
        pass
    
    return []

def check_file_access_via_psutil(file_path):
    """
    通过psutil检查文件是否被访问
    返回正在访问该文件的进程列表
    """
    accessing_processes = []
    
    try:
        for proc in psutil.process_iter():
            try:
                # 检查命令行参数
                cmdline = proc.cmdline()
                if cmdline and file_path.lower() in ' '.join(cmdline).lower():
                    accessing_processes.append(proc.pid)
                    continue
                
                # 检查打开的文件
                try:
                    for f in proc.open_files():
                        if normalize_path(f.path) == normalize_path(file_path):
                            accessing_processes.append(proc.pid)
                            break
                except (psutil.AccessDenied, psutil.ZombieProcess, psutil.NoSuchProcess, NotImplementedError):
                    pass
                    
            except (psutil.AccessDenied, psutil.ZombieProcess, psutil.NoSuchProcess):
                continue
                
    except Exception:
        pass
    
    return accessing_processes

class TaskMonitor(threading.Thread):
    """
    单个任务的监控线程（守护线程）
    - 通过多种策略判断某个进程是否在使用目标文件
    - 一旦检测到文件被打开就开始计时，文件关闭则记录一次日志
    """
    def __init__(self, task_path, update_callback, scan_interval=2.0):
        super().__init__(daemon=True)
        self.task_path = normalize_path(task_path)    # 任务文件的绝对标准化路径
        self.task_basename = os.path.basename(task_path).lower()
        self.ext = os.path.splitext(task_path)[1].lower()
        self.update_callback = update_callback       # 线程安全的 UI 更新回调 (主线程执行)
        self.scan_interval = scan_interval
        self._stop_event = threading.Event()

        # 计时变量
        self.start_time = None
        self.end_time = None

    def stop(self):
        """停止线程运行"""
        self._stop_event.set()

    def stopped(self):
        return self._stop_event.is_set()

    def _is_file_being_accessed(self):
        """
        使用多种方法检查文件是否正在被访问
        返回 True 如果文件正在被访问，False 否则
        """
        # 方法1: 检查文件是否被锁定（Windows）
        if WINDOWS_AVAILABLE and is_file_locked_windows(self.task_path):
            return True
        
        # 方法2: 通过psutil检查进程访问
        accessing_pids = check_file_access_via_psutil(self.task_path)
        if accessing_pids:
            return True
        
        # 方法3: 检查相关应用程序是否在运行
        return self._check_related_apps_running()
    
    def _check_related_apps_running(self):
        """
        检查与文件类型相关的应用程序是否正在运行
        这对于文档文件特别有效，因为应用程序可能不会在命令行中显示文件路径
        """
        try:
            # 获取文件扩展名对应的进程模式
            app_type = None
            if self.ext in ['.doc', '.docx']:
                app_type = 'word'
            elif self.ext in ['.xls', '.xlsx']:
                app_type = 'excel'
            elif self.ext in ['.ppt', '.pptx']:
                app_type = 'powerpoint'
            elif self.ext == '.pdf':
                app_type = 'pdf'
            elif self.ext in ['.txt', '.py', '.js', '.html', '.css']:
                app_type = 'text'
            
            if not app_type:
                return False
            
            # 检查是否有相关进程在运行
            for proc in psutil.process_iter():
                try:
                    proc_name = proc.name().lower()
                    if proc_name in PROCESS_PATTERNS.get(app_type, []):
                        # 进一步检查这个进程是否可能在使用我们的文件
                        if self._proc_may_open_file(proc):
                            return True
                except (psutil.AccessDenied, psutil.ZombieProcess, psutil.NoSuchProcess):
                    continue
                    
        except Exception:
            pass
        
        return False

    def _proc_may_open_file(self, proc):
        """
        判断进程 proc 是否可能正在打开/运行 self.task_path
        1) 检查命令行参数（cmdline）是否包含文件路径或 basename
        2) 检查 proc.open_files() 返回是否包含该文件（当可用且有权限时）
        3) 兜底：根据扩展名匹配常见应用进程名关键字
        返回 True/False
        """
        try:
            # 有些进程访问属性会抛异常，全部包裹
            # 1) 检查 cmdline（可在大多数系统上获取）
            try:
                cmd = proc.cmdline()
                if cmd:
                    joined = " ".join(cmd).lower()
                    # 使用标准化路径比较，及 basename 比较（考虑某些app只用 basename）
                    if normalize_path(self.task_path) in joined or self.task_basename in joined:
                        return True
            except (psutil.AccessDenied, psutil.ZombieProcess, psutil.NoSuchProcess):
                pass

            # 2) 检查 open_files（更可靠，但可能受权限限制或平台限制）
            try:
                for f in proc.open_files():
                    # f.path 可能是绝对路径，做标准化比较
                    if normalize_path(f.path) == self.task_path:
                        return True
            except (psutil.AccessDenied, psutil.ZombieProcess, psutil.NoSuchProcess, NotImplementedError):
                # NotImplementedError: Windows 在某些情况下不支持或进程拒绝
                pass

            # 3) 兜底：根据扩展名匹配进程名中常见关键词
            try:
                pname = proc.name().lower()
                candidates = FALLBACK_APP_KEYWORDS.get(self.ext, [])
                for kw in candidates:
                    if kw in pname:
                        return True
            except (psutil.AccessDenied, psutil.ZombieProcess, psutil.NoSuchProcess):
                pass

        except Exception:
            # 任何意外都认为此进程不匹配，继续检查其它进程
            return False

        return False

    def run(self):
        """
        循环扫描系统进程，判断任务何时开始与结束
        - 使用多种方法检测文件是否正在被访问
        - 从未运行到运行：记录 start_time
        - 从运行到未运行：记录 end_time、保存日志并通知 UI
        """
        while not self.stopped():
            # 使用新的综合检测方法
            found = self._is_file_being_accessed()
            
            if found:
                # 若尚未开始计时，则开始
                if not self.start_time:
                    self.start_time = time.time()
                    # 通知 UI（通过主线程 safe 回调）
                    self.update_callback(self.task_path, "运行中")
                    # 打印调试信息
                    print(f"[监控] 任务开始：{self.task_path}")

            # 如果曾经在运行，现在找不到任何进程在使用 —— 认为任务结束
            if self.start_time and not found:
                self.end_time = time.time()
                duration = self.end_time - self.start_time
                # 保存记录
                try:
                    self.save_record(duration)
                except Exception as e:
                    print(f"[监控] 保存记录失败：{e}")
                # 通知 UI（通过主线程 safe 回调）
                self.update_callback(self.task_path, f"完成: {duration:.2f} 秒")
                print(f"[监控] 任务结束：{self.task_path}，耗时 {duration:.2f} 秒")
                # 重置计时准备下一次打开
                self.start_time = None
                self.end_time = None

            time.sleep(self.scan_interval)

    def save_record(self, duration):
        """
        将计时记录追加到 CSV 日志
        字段：任务文件（完整路径），开始时间，结束时间，耗时（秒）
        """
        record_file = "task_time_log.csv"
        file_exists = os.path.isfile(record_file)
        with open(record_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["任务文件(绝对路径)", "开始时间", "结束时间", "耗时(秒)"])
            st = datetime.fromtimestamp(self.start_time).strftime("%Y-%m-%d %H:%M:%S") if self.start_time else ""
            et = datetime.fromtimestamp(self.end_time).strftime("%Y-%m-%d %H:%M:%S") if self.end_time else ""
            writer.writerow([self.task_path, st, et, round(duration, 2)])


class TaskTimerApp:
    """
    主界面及任务管理
    - 支持添加多个任务（使用完整路径作为唯一键）
    - 每个任务对应一个 TaskMonitor 线程
    - UI 列表显示每个任务的当前状态
    """
    def __init__(self, root):
        self.root = root
        self.root.title("任务计时器 - 并行监控（增强版）")
        self.root.geometry("700x380")

        # 存放 monitor 的 dict： key = 标准化绝对路径, value = TaskMonitor 对象
        self.monitors = {}
        # 存放 listbox 的索引映射： key = 标准化绝对路径, value = listbox index
        self.listbox_index = {}

        # 顶部说明
        lbl = tk.Label(root, text="选择任意文件（文档/可执行/表格等），程序会在对应应用打开该文件时开始计时，关闭时记录耗时。", wraplength=650, justify="left")
        lbl.pack(padx=10, pady=8)

        # 按钮行
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=6)
        self.add_btn = tk.Button(btn_frame, text="添加任务文件", command=self.choose_file)
        self.add_btn.pack(side=tk.LEFT, padx=6)
        self.stop_all_btn = tk.Button(btn_frame, text="停止所有监控", command=self.stop_all, state=tk.DISABLED)
        self.stop_all_btn.pack(side=tk.LEFT, padx=6)
        self.open_log_btn = tk.Button(btn_frame, text="打开日志文件夹", command=self.open_log_folder)
        self.open_log_btn.pack(side=tk.LEFT, padx=6)

        # 任务状态列表
        self.task_listbox = tk.Listbox(root, width=100, height=15)
        self.task_listbox.pack(padx=10, pady=8)

        # 绑定窗口关闭，做清理
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def choose_file(self):
        """弹出文件选择对话框并添加任务"""
        file_path = filedialog.askopenfilename(title="选择任务文件")
        if not file_path:
            return
        task_key = normalize_path(file_path)
        if task_key in self.monitors:
            messagebox.showinfo("提示", "该任务（同一路径）已在监控中。")
            return

        # 新建 monitor，传入一个线程安全的 update_callback（在主线程执行）
        def safe_update(task_path, status):
            # 通过 root.after 在主线程更新 UI
            self.root.after(0, lambda: self.update_status(task_path, status))

        monitor = TaskMonitor(file_path, update_callback=safe_update)
        monitor.start()

        # 在列表中插入项并记录索引
        display_name = os.path.basename(file_path) + "    (" + file_path + ")"
        idx = self.task_listbox.size()
        self.task_listbox.insert(tk.END, f"{display_name}  → 等待运行")
        self.listbox_index[task_key] = idx
        self.monitors[task_key] = monitor

        # 启用停止按钮
        self.stop_all_btn.config(state=tk.NORMAL)

    def update_status(self, task_path, status):
        """在 listbox 中更新任务状态（在主线程被调用）"""
        task_key = normalize_path(task_path)
        idx = self.listbox_index.get(task_key)
        if idx is None:
            # 可能任务已被移除，忽略
            return
        # 取原显示名（包含完整路径）
        orig_text = self.task_listbox.get(idx)
        # 原文本形如: "<basename> (fullpath)  → 状态"，我们用 split 保持最左侧不变
        if "→" in orig_text:
            left = orig_text.split("→")[0].rstrip()
        else:
            left = orig_text
        new_text = f"{left}  → {status}"
        self.task_listbox.delete(idx)
        self.task_listbox.insert(idx, new_text)

    def stop_all(self):
        """停止所有监控线程并清理"""
        for monitor in list(self.monitors.values()):
            try:
                monitor.stop()
            except Exception:
                pass
        self.monitors.clear()
        self.listbox_index.clear()
        self.task_listbox.delete(0, tk.END)
        self.task_listbox.insert(tk.END, ">>> 所有监控已停止")
        self.stop_all_btn.config(state=tk.DISABLED)

    def open_log_folder(self):
        """在文件管理器中打开当前工作目录（日志保存在当前目录）"""
        try:
            path = os.path.abspath(".")
            if os.name == "nt":
                os.startfile(path)
            elif os.name == "posix":
                # macOS 或 Linux
                if 'darwin' in os.sys.platform:
                    os.system(f'open "{path}"')
                else:
                    os.system(f'xdg-open "{path}"')
        except Exception as e:
            messagebox.showwarning("错误", f"无法打开文件夹：{e}")

    def on_close(self):
        """程序退出前停止所有线程并退出"""
        if messagebox.askokcancel("退出", "确认退出并停止所有监控？"):
            self.stop_all()
            # 小延迟确保线程退出（通常很快）
            time.sleep(0.2)
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = TaskTimerApp(root)
    root.mainloop()
