# -*- coding: utf-8 -*-
"""
EML附件提取器
功能：批量提取EML邮件文件中的附件
支持：
  - 单个EML文件选择
  - 文件夹批量处理（递归遍历子目录）
  - 自动创建以邮件主题命名的子文件夹
  - 按文件类型分类保存
  - 实时日志显示
"""

import os
import email
import re
from email import policy
from email.parser import BytesParser
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading

# ==================== 配置常量 ====================
SUPPORTED_EXTENSIONS = [".eml"]
DEFAULT_OUTPUT_FOLDER = "提取的附件"


class EmlExtractorApp:
    """EML附件提取器GUI应用"""

    def __init__(self, root):
        self.root = root
        self.root.title("EML附件提取器")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        # 状态变量
        self.processing = False
        self.eml_files = []
        self.output_dir = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        """初始化界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部工具栏
        self.create_toolbar(main_frame)

        # 文件列表区域
        self.create_file_list(main_frame)

        # 输出设置区域
        self.create_output_settings(main_frame)

        # 操作按钮区域
        self.create_action_buttons(main_frame)

        # 日志区域
        self.create_log_panel(main_frame)

        # 状态栏
        self.create_status_bar(main_frame)

    def create_toolbar(self, parent):
        """创建顶部工具栏"""
        toolbar_frame = ttk.LabelFrame(parent, text="选择EML文件", padding="5")
        toolbar_frame.pack(fill=tk.X, pady=(0, 10))

        btn_frame = ttk.Frame(toolbar_frame)
        btn_frame.pack(fill=tk.X)

        ttk.Button(
            btn_frame, text="选择文件", command=self.select_files, width=15
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame, text="选择文件夹", command=self.select_folder, width=15
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame, text="清空列表", command=self.clear_files, width=15
        ).pack(side=tk.LEFT, padx=5)

        # 文件统计标签
        self.file_count_label = ttk.Label(btn_frame, text="已选择: 0 个文件")
        self.file_count_label.pack(side=tk.RIGHT, padx=10)

    def create_file_list(self, parent):
        """创建文件列表区域"""
        list_frame = ttk.LabelFrame(parent, text="待处理文件列表", padding="5")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建Treeview
        columns = ("序号", "文件名", "路径")
        self.file_tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", height=8
        )

        # 设置列
        self.file_tree.heading("序号", text="序号")
        self.file_tree.heading("文件名", text="文件名")
        self.file_tree.heading("路径", text="文件路径")

        self.file_tree.column("序号", width=50, anchor="center")
        self.file_tree.column("文件名", width=200)
        self.file_tree.column("路径", width=450)

        # 滚动条
        scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.file_tree.yview
        )
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_output_settings(self, parent):
        """创建输出设置区域"""
        output_frame = ttk.LabelFrame(parent, text="输出设置", padding="5")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(output_frame, text="输出目录:").pack(side=tk.LEFT, padx=5)

        output_entry = ttk.Entry(
            output_frame, textvariable=self.output_dir, width=60
        )
        output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        ttk.Button(
            output_frame, text="浏览", command=self.select_output_dir, width=10
        ).pack(side=tk.LEFT, padx=5)

        # 选项
        self.create_subfolder = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            output_frame,
            text="按邮件主题创建子文件夹",
            variable=self.create_subfolder,
        ).pack(side=tk.LEFT, padx=10)

        self.classify_by_type = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            output_frame,
            text="按文件类型分类保存",
            variable=self.classify_by_type,
        ).pack(side=tk.LEFT, padx=10)

    def create_action_buttons(self, parent):
        """创建操作按钮区域"""
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        # 进度条
        self.progress = ttk.Progressbar(
            action_frame, orient=tk.HORIZONTAL, mode="determinate"
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        # 开始按钮
        self.start_btn = ttk.Button(
            action_frame,
            text="开始提取",
            command=self.start_extraction,
            width=15,
        )
        self.start_btn.pack(side=tk.RIGHT, padx=5)

    def create_log_panel(self, parent):
        """创建日志面板"""
        log_frame = ttk.LabelFrame(parent, text="处理日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=10, wrap=tk.WORD, font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 清空日志按钮
        ttk.Button(
            log_frame, text="清空日志", command=self.clear_log, width=10
        ).pack(anchor=tk.E, pady=(5, 0))

    def create_status_bar(self, parent):
        """创建状态栏"""
        self.status_label = ttk.Label(
            parent, text="就绪", relief=tk.SUNKEN, anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)

    # ==================== 功能方法 ====================

    def log(self, message, level="INFO"):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def select_files(self):
        """选择EML文件"""
        files = filedialog.askopenfilenames(
            title="选择EML文件",
            filetypes=[("EML文件", "*.eml"), ("所有文件", "*.*")],
        )
        if files:
            self.add_files(files)

    def select_folder(self):
        """选择文件夹（递归遍历所有子目录）"""
        folder = filedialog.askdirectory(title="选择包含EML文件的文件夹")
        if folder:
            folder_path = Path(folder)
            eml_files = list(folder_path.rglob("*.eml"))  # 递归遍历
            if eml_files:
                self.add_files([str(f) for f in eml_files])
                self.log(f"从文件夹及子目录扫描到 {len(eml_files)} 个EML文件")
            else:
                messagebox.showwarning("提示", "该文件夹及子目录中没有找到EML文件")

    def add_files(self, files):
        """添加文件到列表"""
        added_count = 0
        for file_path in files:
            if file_path not in self.eml_files:
                self.eml_files.append(file_path)
                added_count += 1

        self.refresh_file_list()
        if added_count > 0:
            self.log(f"添加了 {added_count} 个文件")

    def refresh_file_list(self):
        """刷新文件列表显示"""
        # 清空现有项目
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # 添加文件
        for i, file_path in enumerate(self.eml_files, 1):
            path = Path(file_path)
            self.file_tree.insert(
                "", tk.END, values=(i, path.name, str(path.parent))
            )

        # 更新计数
        self.file_count_label.config(text=f"已选择: {len(self.eml_files)} 个文件")

    def clear_files(self):
        """清空文件列表"""
        self.eml_files.clear()
        self.refresh_file_list()
        self.log("已清空文件列表")

    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)

    def select_output_dir(self):
        """选择输出目录"""
        folder = filedialog.askdirectory(title="选择附件输出目录")
        if folder:
            self.output_dir.set(folder)

    def sanitize_filename(self, filename):
        """清理文件名中的非法字符"""
        # 替换Windows不允许的字符
        illegal_chars = r'[<>:"/\\|?*]'
        sanitized = re.sub(illegal_chars, "_", filename)
        # 移除首尾空格和点
        sanitized = sanitized.strip(" .")
        # 限制长度
        if len(sanitized) > 200:
            sanitized = sanitized[:200]
        return sanitized if sanitized else "未命名"

    def decode_header(self, header_value):
        """解码邮件头"""
        if header_value is None:
            return ""
        decoded_parts = email.header.decode_header(header_value)
        decoded_str = ""
        for part, charset in decoded_parts:
            if isinstance(part, bytes):
                try:
                    decoded_str += part.decode(charset or "utf-8", errors="replace")
                except (LookupError, UnicodeDecodeError):
                    decoded_str += part.decode("utf-8", errors="replace")
            else:
                decoded_str += part
        return decoded_str

    def extract_attachments(self, eml_path, output_base_dir):
        """从单个EML文件提取附件"""
        attachments_found = []

        try:
            with open(eml_path, "rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)

            # 获取邮件主题
            subject = self.decode_header(msg.get("Subject", ""))
            if not subject:
                subject = Path(eml_path).stem

            safe_subject = self.sanitize_filename(subject)

            # 确定基础输出目录
            if self.create_subfolder.get():
                base_output_dir = Path(output_base_dir) / safe_subject
            else:
                base_output_dir = Path(output_base_dir)

            base_output_dir.mkdir(parents=True, exist_ok=True)

            # 遍历邮件部分，提取附件
            for part in msg.walk():
                content_disposition = part.get_content_disposition()

                if content_disposition == "attachment":
                    filename = part.get_filename()
                    if filename:
                        # 解码文件名
                        filename = self.decode_header(filename)
                        safe_filename = self.sanitize_filename(filename)

                        # 获取附件内容
                        payload = part.get_payload(decode=True)
                        if payload:
                            # 确定最终输出目录（是否按类型分类）
                            if self.classify_by_type.get():
                                ext = Path(safe_filename).suffix.lower().lstrip(".")
                                output_dir = base_output_dir / (ext if ext else "其他")
                            else:
                                output_dir = base_output_dir
                            output_dir.mkdir(parents=True, exist_ok=True)

                            # 处理文件名冲突
                            output_path = output_dir / safe_filename
                            counter = 1
                            while output_path.exists():
                                name, ext = os.path.splitext(safe_filename)
                                output_path = output_dir / f"{name}_{counter}{ext}"
                                counter += 1

                            # 保存附件
                            with open(output_path, "wb") as f:
                                f.write(payload)

                            attachments_found.append(str(output_path))

            return attachments_found, None

        except Exception as e:
            return [], str(e)

    def start_extraction(self):
        """开始提取"""
        if self.processing:
            messagebox.showwarning("提示", "正在处理中，请稍候...")
            return

        if not self.eml_files:
            messagebox.showwarning("提示", "请先选择要处理的EML文件")
            return

        # 检查输出目录
        output_dir = self.output_dir.get()
        if not output_dir:
            # 使用第一个文件所在目录作为默认输出
            first_file = Path(self.eml_files[0])
            output_dir = str(first_file.parent / DEFAULT_OUTPUT_FOLDER)
            self.output_dir.set(output_dir)

        # 在新线程中执行
        self.processing = True
        self.start_btn.config(state=tk.DISABLED)
        thread = threading.Thread(target=self.extraction_worker, args=(output_dir,))
        thread.daemon = True
        thread.start()

    def extraction_worker(self, output_dir):
        """提取工作线程"""
        total_files = len(self.eml_files)
        total_attachments = 0
        success_count = 0
        error_count = 0

        self.log(f"开始处理 {total_files} 个EML文件...")
        self.log(f"输出目录: {output_dir}")
        self.update_status("正在处理...")

        for i, eml_path in enumerate(self.eml_files):
            # 更新进度
            progress_value = (i + 1) / total_files * 100
            self.progress["value"] = progress_value
            self.update_status(f"正在处理: {i + 1}/{total_files}")

            file_name = Path(eml_path).name
            self.log(f"处理: {file_name}")

            # 提取附件
            attachments, error = self.extract_attachments(eml_path, output_dir)

            if error:
                self.log(f"  错误: {error}", "ERROR")
                error_count += 1
            elif attachments:
                for att in attachments:
                    self.log(f"  提取: {Path(att).name}")
                total_attachments += len(attachments)
                success_count += 1
            else:
                self.log(f"  该邮件没有附件")
                success_count += 1

        # 完成
        self.progress["value"] = 100
        self.log("=" * 50)
        self.log(f"处理完成!")
        self.log(f"成功处理: {success_count} 个文件")
        self.log(f"处理失败: {error_count} 个文件")
        self.log(f"提取附件: {total_attachments} 个")
        self.log(f"输出目录: {output_dir}")

        self.update_status(
            f"完成 - 提取了 {total_attachments} 个附件 (成功: {success_count}, 失败: {error_count})"
        )

        # 重置状态
        self.processing = False
        self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))

        # 询问是否打开输出目录
        if total_attachments > 0:
            self.root.after(
                100,
                lambda: self.ask_open_folder(output_dir),
            )

    def ask_open_folder(self, folder_path):
        """询问是否打开输出文件夹"""
        if messagebox.askyesno("完成", f"提取完成！\n是否打开输出文件夹？"):
            try:
                os.startfile(folder_path)
            except Exception as e:
                self.log(f"无法打开文件夹: {e}", "ERROR")


def main():
    """主函数"""
    root = tk.Tk()
    app = EmlExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
