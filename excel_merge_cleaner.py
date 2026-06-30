#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 批量合并与数据清洗工具 - Excel Merge & Cleaner
适用平台: Windows
技术栈: Python 3 + tkinter + pandas + openpyxl

功能:
  - 选择多个 Excel 文件
  - 合并所有文件的第一张工作表
  - 清理空行 / 去除重复行
  - 导出含"合并总表"和"处理统计"的新 Excel
  - 生成 log.txt 记录操作
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import datetime
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class ExcelMergeCleaner:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🔮 Excel 批量合并与数据清洗 v1.0")
        self.root.geometry("750x500")
        self.root.minsize(600, 400)

        self.files = []  # 选择的文件路径列表
        self._setup_ui()

    def _setup_ui(self):
        root = self.root

        # ── 顶部: 文件选择 ──
        top_frame = ttk.Frame(root, padding=5)
        top_frame.pack(fill=tk.X)

        ttk.Button(top_frame, text="📂 选择 Excel 文件", command=self.select_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(top_frame, text="清空列表", command=self.clear_files).pack(side=tk.LEFT, padx=2)

        # ── 文件列表 ──
        list_frame = ttk.LabelFrame(root, text="📄 已选择的文件", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

        columns = ("file", "rows", "cols")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        self.tree.heading("file", text="文件名")
        self.tree.heading("rows", text="行数")
        self.tree.heading("cols", text="列数")
        self.tree.column("file", width=400)
        self.tree.column("rows", width=80, anchor=tk.CENTER)
        self.tree.column("cols", width=80, anchor=tk.CENTER)

        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # ── 操作按钮区 ──
        op_frame = ttk.Frame(root, padding=5)
        op_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(op_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT)

        ttk.Button(op_frame, text="✅ 合并并导出", command=self.merge_and_export).pack(side=tk.RIGHT, padx=3)

    def select_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not paths:
            return

        for path in paths:
            if path in self.files:
                continue
            self.files.append(path)
            try:
                df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
                rows_preview = len(df)
                cols_preview = len(df.columns)
            except Exception as e:
                messagebox.showerror("读取失败", f"无法读取文件:\n{path}\n\n错误: {e}")
                rows_preview = "?"
                cols_preview = "?"
            self.tree.insert("", tk.END, values=(os.path.basename(path), rows_preview, cols_preview))

        self.set_status(f"📄 已选择 {len(self.files)} 个文件")

    def clear_files(self):
        self.files.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.set_status("列表已清空")

    def set_status(self, text):
        self.status_label.config(text=text)

    def merge_and_export(self):
        if not self.files:
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return

        output_path = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialdir=BASE_DIR,
            initialfile=f"合并结果_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not output_path:
            return

        log_path = os.path.join(os.path.dirname(output_path), "log.txt")
        log_entries = []
        all_dfs = []
        stats = []  # [(文件名, 原始行数, 清理后行数, 空行)]

        self.set_status("⏳ 正在处理...")
        self.root.update()

        for filepath in self.files:
            fname = os.path.basename(filepath)
            try:
                df = pd.read_excel(filepath, sheet_name=0, engine="openpyxl")
                original_rows = len(df)

                log_entries.append(f"[读取] {fname} | 原始行数={original_rows} | 列数={len(df.columns)}")

                # 标准化列名
                df.columns = [str(c).strip() if pd.notna(c) else f"列{i}" for i, c in enumerate(df.columns)]

                # 清理空行（仅本文件内）
                empty_rows = int(df.isna().all(axis=1).sum())
                df = df.dropna(how="all")

                cleaned_rows = len(df)
                all_dfs.append(df)
                stats.append((fname, original_rows, cleaned_rows, empty_rows))

            except Exception as e:
                log_entries.append(f"[错误] {fname} | 读取失败: {e}")
                messagebox.showerror("读取错误", f"文件 {fname} 读取失败:\n{e}")
                self.set_status("❌ 处理中断")
                return

        if not all_dfs:
            messagebox.showwarning("提示", "没有有效数据可合并")
            return

        # 先合并所有数据
        merged = pd.concat(all_dfs, ignore_index=True)
        total_original = sum(s[1] for s in stats)
        total_empty = sum(s[3] for s in stats)
        before_dedup = len(merged)

        # 再对合并结果按整行去重（处理跨文件的重复行）
        merged = merged.drop_duplicates()
        dup_removed = before_dedup - len(merged)
        final_rows = len(merged)

        log_entries.append(f"[去重] 合并后总行数={before_dedup} | 删除重复行={dup_removed} | 最终行数={final_rows}")

        # 处理统计表
        stat_data = []
        for s in stats:
            success = "✅" if s[2] > 0 else "⚠️ 无数据"
            stat_data.append({
                "源文件名": s[0],
                "原始行数": s[1],
                "清理后行数": s[2],
                "移除空行": s[3],
                "状态": success,
            })
        stat_data.append({
            "源文件名": "📊 合计",
            "原始行数": total_original,
            "清理后行数": before_dedup,
            "移除空行": total_empty,
            "状态": f"去重删除了{dup_removed}行，最终{final_rows}行",
        })
        stat_df = pd.DataFrame(stat_data)

        # 导出
        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                merged.to_excel(writer, sheet_name="合并总表", index=False)
                stat_df.to_excel(writer, sheet_name="处理统计", index=False)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            self.set_status("❌ 导出失败")
            return

        # 写 log.txt
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"\n{'=' * 60}\n")
                f.write(f"操作时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"输出文件: {output_path}\n")
                f.write(f"{'=' * 60}\n")
                for entry in log_entries:
                    f.write(entry + "\n")
                f.write(f"\n最终结果: 合并={len(all_dfs)}个文件 | 合并前行数={before_dedup}")
                f.write(f" | 移除空行={total_empty} | 删除重复行={dup_removed} | 最终行数={final_rows}\n\n")
        except Exception as e:
            self.set_status(f"⚠️ log.txt 写入失败: {e}")

        self.set_status(f"✅ 合并完成: {len(all_dfs)} 个文件 -> {output_path}")

        msg = (
            f"✅ 合并完成！\n\n"
            f"文件数量: {len(all_dfs)}\n"
            f"合并前总行数: {before_dedup}\n"
            f"移除空行: {total_empty}\n"
            f"删除重复行: {dup_removed}\n"
            f"最终行数: {final_rows}\n\n"
            f"导出路径:\n{output_path}"
        )
        messagebox.showinfo("执行结果", msg)

    def run(self):
        self.root.mainloop()


def main():
    app = ExcelMergeCleaner()
    app.run()


if __name__ == "__main__":
    main()
