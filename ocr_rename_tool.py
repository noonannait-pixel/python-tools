#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证件图片 OCR 批量重命名工具
==============================
适用平台: Windows
技术栈: Python 3 + tkinter + rapidocr-onnxruntime + Pillow

功能:
  - 选择图片文件夹
  - 自动 OCR 识别图片中的"姓名"
  - 按姓名重命名并复制到 output 文件夹
  - 重名自动编号
  - 识别失败归入"识别失败"子文件夹
  - 自动生成 3 张测试图片
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import sys
import shutil
import datetime
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

# ── 全局配置 ──
SUPPORTED_EXT = {".jpg", ".jpeg", ".png"}

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

OUTPUT_DIR = BASE_DIR / "output"
FAILED_DIR = OUTPUT_DIR / "识别失败"
TEST_DIR = BASE_DIR
LOG_FILE = BASE_DIR / "log.txt"


class OCRRenameTool:
    """OCR 批量重命名主程序"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🔮 证件图片OCR批量重命名 v1.0")
        self.root.geometry("800x600")
        self.root.minsize(650, 500)

        self.input_dir = tk.StringVar()  # 输入文件夹路径
        self.ocr_engine = None           # OCR 引擎实例（延迟加载）
        self._setup_ui()
        self._ensure_test_images()

    # ════════════════════════════════════════════
    #  UI 布局
    # ════════════════════════════════════════════

    def _setup_ui(self):
        root = self.root

        # ── 顶部: 文件夹选择 ──
        top = ttk.Frame(root, padding=5)
        top.pack(fill=tk.X)

        ttk.Label(top, text="📁 图片文件夹:").pack(side=tk.LEFT)
        ttk.Entry(top, textvariable=self.input_dir, width=55).pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="选择...", command=self.select_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="🖼️ 生成测试图片", command=self.generate_test_images).pack(side=tk.LEFT, padx=2)

        # ── 操作按钮 ──
        action_frame = ttk.Frame(root, padding=5)
        action_frame.pack(fill=tk.X)

        self.start_btn = ttk.Button(action_frame, text="🚀 开始识别并重命名", command=self.start_process, width=25)
        self.start_btn.pack(side=tk.LEFT, padx=2)

        self.status_label = ttk.Label(action_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=10)

        self.progress = ttk.Progressbar(action_frame, mode='determinate', length=200)
        self.progress.pack(side=tk.RIGHT, padx=5)

        # ── 结果列表 ──
        result_frame = ttk.LabelFrame(root, text="📋 处理结果", padding=5)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

        columns = ("original", "ocr_text", "name", "new_name", "status")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=10)
        self.tree.heading("original", text="原文件名")
        self.tree.heading("ocr_text", text="OCR识别文本")
        self.tree.heading("name", text="提取姓名")
        self.tree.heading("new_name", text="新文件名")
        self.tree.heading("status", text="状态")

        self.tree.column("original", width=120)
        self.tree.column("ocr_text", width=200)
        self.tree.column("name", width=80, anchor=tk.CENTER)
        self.tree.column("new_name", width=150)
        self.tree.column("status", width=80, anchor=tk.CENTER)

        scroll_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # ── 日志区 ──
        log_frame = ttk.LabelFrame(root, text="📝 运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, padx=5, pady=2)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ════════════════════════════════════════════
    #  日志 / 状态
    # ════════════════════════════════════════════

    def log(self, msg):
        """追加日志到界面和文件"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {msg}"
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, line + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()

    def set_status(self, text):
        self.status_label.config(text=text)
        self.root.update()

    # ════════════════════════════════════════════
    #  文件夹选择
    # ════════════════════════════════════════════

    def select_folder(self):
        path = filedialog.askdirectory(title="选择包含证件图片的文件夹")
        if path:
            self.input_dir.set(path)

    # ════════════════════════════════════════════
    #  生成测试图片（用 Pillow 绘制中文）
    # ════════════════════════════════════════════

    def _find_chinese_font(self):
        """尝试查找系统可用的中文字体"""
        candidates = [
            "C:\\Windows\\Fonts\\msyh.ttc",       # 微软雅黑
            "C:\\Windows\\Fonts\\simsun.ttc",     # 宋体
            "C:\\Windows\\Fonts\\simhei.ttf",     # 黑体
            "C:\\Windows\\Fonts\\yahei.ttf",      # 雅黑变体
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
        # 如果都不存在，回退到默认字体（可能不显示中文）
        return None

    def _make_test_image(self, name, filename, width=400, height=300):
        """生成一张带姓名文字的模拟证件图片"""
        img = Image.new("RGB", (width, height), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)

        # 加载中文字体
        font_path = self._find_chinese_font()
        if font_path:
            try:
                font_large = ImageFont.truetype(font_path, 48)
                font_small = ImageFont.truetype(font_path, 24)
            except Exception:
                font_large = ImageFont.load_default()
                font_small = ImageFont.load_default()
        else:
            font_large = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # 画一个简单"证件照"背景框架
        draw.rectangle([20, 20, width - 20, height - 20], outline=(200, 200, 200), width=2)
        draw.text((width // 2 - 60, 40), "证件照片", fill=(150, 150, 150), font=font_small)

        # 中央写姓名
        text = f"姓名：{name}"
        # 粗略居中
        try:
            bbox = draw.textbbox((0, 0), text, font=font_large)
            tw = bbox[2] - bbox[0]
        except Exception:
            tw = len(text) * 24
        tx = (width - tw) // 2
        ty = (height - 60) // 2
        draw.text((tx, ty), text, fill=(0, 0, 0), font=font_large)

        # 底部写一些干扰文字（模拟真实证件有更多文字）
        draw.text((50, height - 60), "编号: 2024XXXXX", fill=(180, 180, 180), font=font_small)
        draw.text((50, height - 35), "其他信息不处理", fill=(180, 180, 180), font=font_small)

        img.save(filename, "JPEG")
        return filename

    def _ensure_test_images(self):
        """程序启动时检查并生成测试图片"""
        test_cases = [
            ("张三", TEST_DIR / "test_张三.jpg"),
            ("李四", TEST_DIR / "test_李四.jpg"),
            ("张三", TEST_DIR / "test_张三_2.jpg"),  # 同名测试
        ]
        generated = []
        for name, path in test_cases:
            if not path.exists():
                self._make_test_image(name, str(path))
                generated.append(path.name)

        if generated:
            self.log(f"✅ 已生成测试图片: {', '.join(generated)}")
        # 如果图片已存在，也记录一下
        existing = [str(p.name) for _, p in test_cases if p.exists()]
        self.set_status(f"测试图片就绪 ({len(existing)} 张)")

    def generate_test_images(self):
        """手动点击生成测试图片"""
        test_cases = [
            ("张三", TEST_DIR / "test_张三.jpg"),
            ("李四", TEST_DIR / "test_李四.jpg"),
            ("张三", TEST_DIR / "test_张三_2.jpg"),
        ]
        for name, path in test_cases:
            self._make_test_image(name, str(path))
        self.log("✅ 测试图片已重新生成 (张三 x2, 李四 x1)")
        self.set_status("测试图片已生成")

    # ════════════════════════════════════════════
    #  OCR 引擎初始化
    # ════════════════════════════════════════════

    def _init_ocr(self):
        """延迟加载 OCR 引擎"""
        if self.ocr_engine is None:
            try:
                self.log("⏳ 正在加载 OCR 引擎 (rapidocr-onnxruntime)...")
                from rapidocr_onnxruntime import RapidOCR
                self.ocr_engine = RapidOCR()
                self.log("✅ OCR 引擎加载完成")
            except ImportError:
                messagebox.showerror("依赖缺失",
                    "缺少 rapidocr-onnxruntime，请运行:\n"
                    "pip install rapidocr-onnxruntime")
                return False
            except Exception as e:
                messagebox.showerror("OCR 加载失败", str(e))
                return False
        return True

    # ════════════════════════════════════════════
    #  姓名提取规则
    # ════════════════════════════════════════════

    def _extract_name(self, ocr_text):
        """从 OCR 识别文本中提取姓名
        支持格式:
          - 姓名：张三
          - 姓名: 张三
          - 姓名 张三
        """
        if not ocr_text:
            return None
        # 匹配 "姓名" 后跟冒号（全角/半角）或空格
        patterns = [
            r"姓名[\s:：]*([\u4e00-\u9fff\w]+)",    # 姓名：xxx / 姓名: xxx / 姓名 xxx
            r"姓名[\s:：]*([\u4e00-\u9fff]{2,4})",   # 更精确：姓名后跟2-4个汉字
        ]
        for pattern in patterns:
            match = re.search(pattern, ocr_text)
            if match:
                name = match.group(1).strip()
                # 过滤明显不是姓名的结果（纯数字、太长、太短）
                if 1 <= len(name) <= 6 and not name.isdigit():
                    return name
        return None

    # ════════════════════════════════════════════
    #  文件处理
    # ════════════════════════════════════════════

    def _safe_copy(self, src_path, dest_dir, base_name):
        """复制文件到目标目录，自动处理重名编号"""
        dest_dir.mkdir(parents=True, exist_ok=True)
        ext = Path(src_path).suffix.lower()
        dest_path = dest_dir / f"{base_name}{ext}"
        counter = 1
        while dest_path.exists():
            dest_path = dest_dir / f"{base_name}_{counter}{ext}"
            counter += 1
        shutil.copy2(src_path, str(dest_path))
        return dest_path

    # ════════════════════════════════════════════
    #  主流程
    # ════════════════════════════════════════════

    def start_process(self):
        """开始 OCR 识别并重命名"""
        input_path = self.input_dir.get()
        if not input_path or not os.path.isdir(input_path):
            messagebox.showwarning("提示", "请先选择有效的图片文件夹")
            return

        if not self._init_ocr():
            return

        # 收集图片文件
        image_files = []
        for f in os.scandir(input_path):
            if f.is_file() and Path(f.name).suffix.lower() in SUPPORTED_EXT:
                image_files.append(f.path)

        if not image_files:
            messagebox.showinfo("提示", "所选文件夹中没有 jpg/png 图片")
            return

        # 清空旧结果
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 准备输出目录
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        FAILED_DIR.mkdir(parents=True, exist_ok=True)

        total = len(image_files)
        success_count = 0
        fail_count = 0
        log_entries = []

        self.start_btn.config(state=tk.DISABLED)
        self.progress["maximum"] = total
        self.progress["value"] = 0
        self.log(f"📂 开始处理: {input_path} ({total} 张图片)")

        for idx, img_path in enumerate(image_files):
            fname = Path(img_path).name
            self.set_status(f"⏳ ({idx + 1}/{total}) {fname}")
            self.progress["value"] = idx + 1
            self.root.update()

            # ── OCR 识别 ──
            ocr_raw = ""
            extracted_name = None
            try:
                result, elapse = self.ocr_engine(img_path)
                if result:
                    # 合并所有识别文本
                    text_parts = [line[1] for line in result if line and len(line) > 1]
                    ocr_raw = " ".join(text_parts)
                    extracted_name = self._extract_name(ocr_raw)
            except Exception as e:
                ocr_raw = f"[OCR错误] {e}"

            # ── 处理 ──
            if extracted_name:
                # 成功：复制到 output，按姓名重命名
                try:
                    dest = self._safe_copy(img_path, OUTPUT_DIR, extracted_name)
                    new_name = dest.name
                    status = "✅ 成功"
                    success_count += 1
                except Exception as e:
                    new_name = f"[复制失败] {e}"
                    status = "❌ 错误"
                    fail_count += 1
                    # 复制失败的也存到失败目录
                    try:
                        self._safe_copy(img_path, FAILED_DIR, Path(img_path).stem)
                    except Exception:
                        pass
            else:
                # 失败：复制到 识别失败
                status = "⚠️ 未识别"
                fail_count += 1
                new_name = f"识别失败/{Path(img_path).stem}{Path(img_path).suffix}"
                try:
                    self._safe_copy(img_path, FAILED_DIR, Path(img_path).stem)
                except Exception as e:
                    new_name = f"[复制失败] {e}"
                    status = "❌ 错误"

            # ── 界面更新 ──
            display_ocr = (ocr_raw[:50] + "...") if len(ocr_raw) > 50 else ocr_raw
            self.tree.insert("", tk.END, values=(
                fname,
                display_ocr or "(空)",
                extracted_name or "(未识别)",
                new_name,
                status,
            ))

            # ── 日志 ──
            log_entry = (
                f"[{status}] 原文件={fname} | OCR文本={ocr_raw[:80]} | "
                f"姓名={extracted_name or '(未识别)'} | 新文件={new_name}"
            )
            log_entries.append(log_entry)

        # ── 写 log.txt ──
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"\n{'=' * 60}\n")
                f.write(f"操作时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"输入目录: {input_path}\n")
                f.write(f"输出目录: {OUTPUT_DIR}\n")
                f.write(f"{'=' * 60}\n")
                for entry in log_entries:
                    f.write(entry + "\n")
                f.write(f"\n结果: 成功={success_count} | 失败={fail_count} | 总计={total}\n\n")
            self.log(f"📝 log.txt 已保存: {LOG_FILE}")
        except Exception as e:
            self.log(f"⚠️ log.txt 写入失败: {e}")

        # ── 完成 ──
        self.start_btn.config(state=tk.NORMAL)
        self.set_status(f"✅ 完成: 成功 {success_count} / 失败 {fail_count}")
        self.progress["value"] = 0
        self.log(f"🏁 处理完成: {total} 张图片, 成功 {success_count}, 失败 {fail_count}")

        messagebox.showinfo(
            "处理完成",
            f"成功: {success_count} 张\n"
            f"失败/未识别: {fail_count} 张\n\n"
            f"输出目录:\n{OUTPUT_DIR}\n\n"
            f"失败图片已在:\n{FAILED_DIR}"
        )

    # ════════════════════════════════════════════
    #  启动
    # ════════════════════════════════════════════

    def run(self):
        self.root.mainloop()


def main():
    app = OCRRenameTool()
    app.run()


if __name__ == "__main__":
    main()
