#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件批量整理 / 重命名工具 - File Organizer
适用平台: Windows
技术栈: Python 3 + tkinter

安全设计:
  - 默认操作 test_files 测试文件夹
  - 禁止操作桌面/系统/下载等关键目录
  - 所有变更必须先预览，确认后才执行
  - 文件名冲突自动加 _1 _2 后缀
  - 生成 log.txt 记录完整操作
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import datetime
import sys

# ── 安全限制 ───────────────────────────────────
FORBIDDEN_DIRS = [
    os.path.expanduser("~\\Desktop"),
    os.path.expanduser("~\\Downloads"),
    os.path.expanduser("~\\Documents"),
    os.environ.get("WINDIR", "C:\\Windows"),
    "C:\\Windows",
    "C:\\Program Files",
    "C:\\Program Files (x86)",
]

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_TEST_DIR = os.path.join(BASE_DIR, "test_files")


class FileOrganizer:
    """文件批量整理/重命名 GUI 工具"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🔮 批量文件整理工具 v1.0")
        self.root.geometry("900x650")
        self.root.minsize(700, 500)

        # 当前工作目录
        self.current_dir = tk.StringVar(value=DEFAULT_TEST_DIR)
        # 文件列表: [(原名, 大小, 类型, 修改时间), ...]
        self.files = []
        # 预览列表: [(原名, 新名, 操作类型), ...]
        self.preview_list = []

        self._setup_ui()
        self._ensure_test_folder()

    # ════════════════════════════════════════════
    #  UI 搭建
    # ════════════════════════════════════════════

    def _setup_ui(self):
        """创建界面元素"""
        root = self.root

        # ── 顶部: 文件夹选择 ──
        top_frame = ttk.Frame(root, padding=5)
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="📁 目标文件夹:").pack(side=tk.LEFT)
        ttk.Entry(top_frame, textvariable=self.current_dir, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="选择...", command=self.select_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(top_frame, text="刷新", command=self.load_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(top_frame, text="🔄 重置测试", command=self.reset_test_folder).pack(side=tk.LEFT, padx=2)

        # ── 文件列表 ──
        list_frame = ttk.LabelFrame(root, text="📄 文件列表", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

        columns = ("name", "size", "ext", "mtime")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=12)
        self.tree.heading("name", text="文件名")
        self.tree.heading("size", text="大小")
        self.tree.heading("ext", text="类型")
        self.tree.heading("mtime", text="修改时间")
        self.tree.column("name", width=300)
        self.tree.column("size", width=100, anchor=tk.E)
        self.tree.column("ext", width=80, anchor=tk.CENTER)
        self.tree.column("mtime", width=150, anchor=tk.CENTER)

        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # ── 操作按钮区 ──
        op_frame = ttk.LabelFrame(root, text="⚙️ 操作（点击后自动预览）", padding=5)
        op_frame.pack(fill=tk.X, padx=5, pady=2)

        btn_frame = ttk.Frame(op_frame)
        btn_frame.pack()

        ttk.Button(btn_frame, text="📂 按类型分类整理", command=self.preview_classify).pack(side=tk.LEFT, padx=3, pady=2)
        ttk.Button(btn_frame, text="➕ 添加前缀", command=lambda: self._prompt_text("输入前缀:", self.preview_add_prefix)).pack(side=tk.LEFT, padx=3, pady=2)
        ttk.Button(btn_frame, text="➕ 添加后缀", command=lambda: self._prompt_text("输入后缀:", self.preview_add_suffix)).pack(side=tk.LEFT, padx=3, pady=2)
        ttk.Button(btn_frame, text="🕒 按修改时间重命名", command=self.preview_rename_by_time).pack(side=tk.LEFT, padx=3, pady=2)

        # ── 预览区 ──
        preview_frame = ttk.LabelFrame(root, text="👁️ 预览（原名 → 新名）", padding=5)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

        p_columns = ("old", "new", "action")
        self.preview_tree = ttk.Treeview(preview_frame, columns=p_columns, show="headings", height=5)
        self.preview_tree.heading("old", text="原名")
        self.preview_tree.heading("new", text="新名")
        self.preview_tree.heading("action", text="操作")
        self.preview_tree.column("old", width=280)
        self.preview_tree.column("new", width=280)
        self.preview_tree.column("action", width=120, anchor=tk.CENTER)

        p_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=p_scroll.set)
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        p_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # ── 底部: 执行按钮 ──
        bottom_frame = ttk.Frame(root, padding=5)
        bottom_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(bottom_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT)

        ttk.Button(bottom_frame, text="🚫 清空预览", command=self.clear_preview).pack(side=tk.RIGHT, padx=3)
        ttk.Button(bottom_frame, text="✅ 确认执行", command=self.execute_changes).pack(side=tk.RIGHT, padx=3)

    # ════════════════════════════════════════════
    #  初始化和安全检测
    # ════════════════════════════════════════════

    def _ensure_test_folder(self):
        """确保测试文件夹存在并生成测试文件"""
        test_dir = DEFAULT_TEST_DIR
        if not os.path.exists(test_dir):
            os.makedirs(test_dir)
            self._create_test_files(test_dir)
            self.current_dir.set(test_dir)
            self.set_status(f"✅ 已创建测试文件夹: {test_dir}")
        self.load_files()

    def reset_test_folder(self):
        """重置测试文件夹（仅限 test_files）"""
        dir_path = self.current_dir.get()
        if os.path.realpath(dir_path) != os.path.realpath(DEFAULT_TEST_DIR):
            messagebox.showwarning("限制", "重置功能仅适用于 test_files 文件夹")
            return

        if not messagebox.askyesno("确认重置", f"将清空 {dir_path} 下所有文件并重新生成，是否继续？"):
            return

        for entry in os.scandir(dir_path):
            if entry.is_file():
                os.remove(entry.path)
            elif entry.is_dir():
                shutil.rmtree(entry.path)

        self._create_test_files(dir_path)
        self.load_files()
        self.set_status("✅ test_files 已重置，15 个测试文件重新生成")

    def _create_test_files(self, target_dir):
        """生成测试文件供练手"""
        test_files = {
            "项目报告.docx": "这是一份项目报告的内容。\n" * 20,
            "会议记录.docx": "会议记录内容。\n" * 15,
            "年度总结.docx": "年度总结内容。\n" * 25,
            "财务报表.xlsx": "模拟Excel数据",
            "销售数据.xlsx": "模拟销售数据",
            "客户名单.xlsx": "客户清单",
            "产品照片.jpg": "模拟图片数据",
            "团队合照.jpg": "模拟图片数据",
            "LOGO图标.jpg": "模拟图片数据",
            "需求文档.txt": "需求文档内容。\n" * 10,
            "安装说明.txt": "安装说明。\n" * 8,
            "README.txt": "项目说明文件",
            "演示文稿.pptx": "模拟PPT数据",
            "培训材料.pptx": "模拟培训PPT",
            "数据备份.zip": "压缩包数据",
        }

        for name, content in test_files.items():
            path = os.path.join(target_dir, name)
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
            except Exception:
                pass  # 二进制文件写文本不影响测试

        # 设置不同修改时间（便于按时间排序测试）
        base = datetime.datetime.now()
        for i, name in enumerate(test_files):
            path = os.path.join(target_dir, name)
            if os.path.exists(path):
                t = base - datetime.timedelta(days=i, hours=i)
                os.utime(path, (t.timestamp(), t.timestamp()))

        self.set_status(f"✅ 已生成 {len(test_files)} 个测试文件")

    def _is_safe_dir(self, dir_path):
        """检查目录是否安全（禁止操作关键系统目录）"""
        if not dir_path:
            return False
        real = os.path.realpath(dir_path)
        for forbidden in FORBIDDEN_DIRS:
            try:
                if os.path.realpath(forbidden) == real:
                    return False
                # 检查是否在禁止目录的子孙目录中
                if real.startswith(os.path.realpath(forbidden) + os.sep):
                    return False
            except Exception:
                continue
        return True

    # ════════════════════════════════════════════
    #  文件加载
    # ════════════════════════════════════════════

    def select_folder(self):
        """选择文件夹（带安全检测）"""
        path = filedialog.askdirectory(title="选择要整理的文件夹")
        if not path:
            return
        if not self._is_safe_dir(path):
            messagebox.showerror("安全限制", f"禁止操作系统关键目录！\n请使用 test_files 或其他安全目录。")
            return
        self.current_dir.set(path)
        self.load_files()

    def load_files(self):
        """加载当前文件夹的文件列表"""
        self.files.clear()
        self.clear_preview()
        dir_path = self.current_dir.get()

        for item in self.tree.get_children():
            self.tree.delete(item)

        if not os.path.isdir(dir_path):
            self.set_status("⚠️ 目录不存在")
            return

        if not self._is_safe_dir(dir_path):
            self.set_status("⛔ 禁止操作系统目录")
            return

        for entry in os.scandir(dir_path):
            if entry.is_file() and not self._is_skip_file(entry.name):
                stat = entry.stat()
                size = self._format_size(stat.st_size)
                ext = os.path.splitext(entry.name)[1].upper() or "(无)"
                mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                self.files.append((entry.name, size, ext, mtime))
                self.tree.insert("", tk.END, values=(entry.name, size, ext, mtime))

        self.set_status(f"📄 共 {len(self.files)} 个文件  |  目录: {dir_path}")

    # ════════════════════════════════════════════
    #  文件过滤
    # ════════════════════════════════════════════

    SKIP_FILES = {"log.txt", "file_organizer.py", "file_organizer.log"}

    def _is_skip_file(self, name):
        """判断是否应该跳过此文件（日志文件、脚本自身等）"""
        if name in self.SKIP_FILES:
            return True
        if name.startswith("."):
            return True
        if name.startswith("~$"):  # Office 临时文件
            return True
        if name.endswith(".tmp"):
            return True
        return False

    def _format_size(self, bytes_val):
        """格式化文件大小"""
        if bytes_val < 1024:
            return f"{bytes_val} B"
        elif bytes_val < 1024 * 1024:
            return f"{bytes_val / 1024:.1f} KB"
        else:
            return f"{bytes_val / 1024 / 1024:.1f} MB"

    # ════════════════════════════════════════════
    #  通用工具
    # ════════════════════════════════════════════

    def set_status(self, text):
        self.status_label.config(text=text)

    def clear_preview(self):
        """清空预览列表"""
        self.preview_list.clear()
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.set_status("预览已清空")

    def _safe_new_name(self, dir_path, base_name, new_name):
        """生成安全的文件名（冲突自动加 _1 _2）"""
        name_no_ext, ext = os.path.splitext(new_name)
        candidate = new_name
        counter = 1
        while os.path.exists(os.path.join(dir_path, candidate)):
            candidate = f"{name_no_ext}_{counter}{ext}"
            counter += 1
        return candidate

    def _prompt_text(self, label, callback):
        """弹出输入框获取用户输入"""
        dialog = tk.Toplevel(self.root)
        dialog.title("输入")
        dialog.geometry("350x120")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text=label, padding=10).pack()
        entry_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=entry_var, width=40)
        entry.pack(padx=10, pady=5)
        entry.focus_set()

        def confirm():
            text = entry_var.get().strip()
            if text:
                dialog.destroy()
                callback(text)
            else:
                messagebox.showwarning("提示", "输入不能为空")

        ttk.Button(dialog, text="确定", command=confirm).pack(pady=5)
        dialog.bind("<Return>", lambda e: confirm())

    def _show_preview(self, preview_data, action_name):
        """展示预览结果"""
        self.clear_preview()
        if not preview_data:
            messagebox.showinfo("提示", "没有需要操作的文件")
            return

        self.preview_list = preview_data
        for old_name, new_name, action in preview_data:
            self.preview_tree.insert("", tk.END, values=(old_name, new_name, action))
        self.set_status(f"👁️ 预览: {len(preview_data)} 个文件将被 {action_name}，请确认后执行")

    # ════════════════════════════════════════════
    #  操作功能（预览模式）
    # ════════════════════════════════════════════

    def preview_classify(self):
        """按类型分类 - 预览"""
        dir_path = self.current_dir.get()
        preview = []
        for entry in os.scandir(dir_path):
            if not entry.is_file() or self._is_skip_file(entry.name):
                continue
            ext = os.path.splitext(entry.name)[1].lower()
            if not ext:
                folder_name = "其他"
            else:
                # 映射常见类型
                type_map = {
                    ".jpg": "图片", ".jpeg": "图片", ".png": "图片", ".gif": "图片",
                    ".bmp": "图片", ".webp": "图片",
                    ".doc": "文档", ".docx": "文档", ".txt": "文档", ".pdf": "文档",
                    ".xls": "表格", ".xlsx": "表格", ".csv": "表格",
                    ".ppt": "演示", ".pptx": "演示",
                    ".zip": "压缩包", ".rar": "压缩包", ".7z": "压缩包",
                    ".py": "代码", ".js": "代码", ".html": "代码", ".css": "代码",
                    ".json": "数据", ".xml": "数据",
                }
                folder_name = type_map.get(ext, ext[1:].upper())
            new_path = os.path.join(folder_name, entry.name)
            preview.append((entry.name, new_path, "移动至「{}」文件夹".format(folder_name)))

        self._show_preview(preview, "按类型分类")

    def preview_add_prefix(self, prefix):
        """添加前缀 - 预览"""
        dir_path = self.current_dir.get()
        preview = []
        for entry in os.scandir(dir_path):
            if not entry.is_file() or self._is_skip_file(entry.name):
                continue
            new_name = prefix + entry.name
            safe_name = self._safe_new_name(dir_path, entry.name, new_name)
            if safe_name != entry.name:
                preview.append((entry.name, safe_name, "添加前缀"))
        self._show_preview(preview, f'添加前缀"{prefix}"')

    def preview_add_suffix(self, suffix):
        """添加后缀（文件名后、扩展名前） - 预览"""
        dir_path = self.current_dir.get()
        preview = []
        for entry in os.scandir(dir_path):
            if not entry.is_file() or self._is_skip_file(entry.name):
                continue
            name_no_ext, ext = os.path.splitext(entry.name)
            new_name = name_no_ext + suffix + ext
            safe_name = self._safe_new_name(dir_path, entry.name, new_name)
            if safe_name != entry.name:
                preview.append((entry.name, safe_name, "添加后缀"))
        self._show_preview(preview, f'添加后缀"{suffix}"')

    def preview_rename_by_time(self):
        """按修改时间重命名 - 预览"""
        dir_path = self.current_dir.get()
        # 按修改时间排序
        files_with_time = []
        for entry in os.scandir(dir_path):
            if not entry.is_file() or self._is_skip_file(entry.name):
                continue
            stat = entry.stat()
            files_with_time.append((entry.name, stat.st_mtime, os.path.splitext(entry.name)[1]))

        files_with_time.sort(key=lambda x: x[1])

        preview = []
        used_names = set()
        for idx, (name, mtime, ext) in enumerate(files_with_time, 1):
            dt = datetime.datetime.fromtimestamp(mtime)
            base = dt.strftime("%Y%m%d_%H%M%S")
            new_name = f"{base}{ext}"
            # 防冲突
            counter = 1
            while new_name in used_names or os.path.exists(os.path.join(dir_path, new_name)):
                new_name = f"{base}_{counter}{ext}"
                counter += 1
            used_names.add(new_name)
            preview.append((name, new_name, "按时间重命名"))

        self._show_preview(preview, "按修改时间重命名")

    # ════════════════════════════════════════════
    #  执行修改
    # ════════════════════════════════════════════

    def execute_changes(self):
        """确认执行预览中的修改"""
        if not self.preview_list:
            messagebox.showinfo("提示", "预览列表为空，请先进行操作预览")
            return

        if not messagebox.askyesno("确认执行", f"即将修改 {len(self.preview_list)} 个文件，是否确认？\n\n操作不可撤销，但 log.txt 会记录所有变更。"):
            return

        dir_path = self.current_dir.get()
        log_entries = []
        success_count = 0
        error_count = 0

        for old_name, new_name, action in self.preview_list:
            old_path = os.path.join(dir_path, old_name)
            new_path = os.path.join(dir_path, new_name)

            if not os.path.exists(old_path):
                error_count += 1
                log_entries.append(f"[错误] 源文件不存在: {old_path}")
                continue

            try:
                # 如果是按类型分类，需要创建子文件夹并移动
                if "文件夹" in action:
                    target_dir = os.path.join(dir_path, os.path.dirname(new_name))
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir)
                    final_path = os.path.join(dir_path, new_name)
                else:
                    final_path = new_path

                os.rename(old_path, final_path)
                success_count += 1
                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_entries.append(f"[修改] 时间={now} | 原路径={old_path} | 新路径={final_path} | 操作={action}")

            except Exception as e:
                error_count += 1
                log_entries.append(f"[错误] 时间={datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {old_name} -> {new_name} | 失败原因={e}")

        # 写入 log.txt
        log_path = os.path.join(dir_path, "log.txt")
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"\n{'=' * 60}\n")
                f.write(f"操作时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"目录: {dir_path}\n")
                f.write(f"操作类型: {self.preview_list[0][2] if self.preview_list else '批量操作'}\n")
                f.write(f"{'=' * 60}\n")
                for entry in log_entries:
                    f.write(entry + "\n")
                f.write(f"结果: 成功={success_count}, 失败={error_count}\n\n")
            self.set_status(f"✅ 执行完成: 成功 {success_count} / 失败 {error_count} | log.txt 已保存")
        except Exception as e:
            self.set_status(f"⚠️ log.txt 写入失败: {e}")

        # 刷新文件列表
        self.clear_preview()
        self.load_files()

        # 统计跳过的文件
        total_files = sum(1 for e in os.scandir(dir_path) if e.is_file() and not self._is_skip_file(e.name))
        skip_count = total_files - len(self.preview_list)
        if skip_count < 0:
            skip_count = 0

        # 提示结果
        log_path = os.path.join(dir_path, "log.txt")
        msg = f"执行完成！\n成功: {success_count}\n跳过: {skip_count}\n失败: {error_count}\n\n日志位置:\n{log_path}"
        messagebox.showinfo("执行结果", msg)

    # ════════════════════════════════════════════
    #  启动
    # ════════════════════════════════════════════

    def run(self):
        self.root.mainloop()


def main():
    # 检查是否操作系统关键目录
    test_dir = DEFAULT_TEST_DIR
    app = FileOrganizer()
    app.run()


if __name__ == "__main__":
    main()
