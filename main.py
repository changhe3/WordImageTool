import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Listbox, END
import threading
import re
import requests
import io
from docx import Document
from docx.shared import Inches
import os


# --- Core Logic (核心处理逻辑已更新，实现精确定位) ---

def find_and_replace_images_in_doc(doc_path, log_callback):
    """
    Opens a .docx file, finds markdown-style image links, and replaces
    ONLY the link with the actual image, preserving surrounding text.

    Args:
        doc_path (str): The path to the Word document.
        log_callback (function): A function to call for logging messages.

    Returns:
        str: The path to the newly saved document, or a status string.
    """
    try:
        doc = Document(doc_path)
        # 正则表达式，捕获整个markdown标签（用于分割）和其中的URL（用于下载）
        # Group 1: The full tag, e.g., ![](url)
        # Group 2: The URL itself
        url_pattern = re.compile(r'(!\[.*?\]\((.*?)\))')
        images_found = 0
        doc_changed = False

        for para in doc.paragraphs:
            # 如果段落中没有 '!' 或 'http'，快速跳过，提升效率
            if '!' not in para.text or 'http' not in para.text:
                continue

            # 检查段落中是否存在匹配项
            if url_pattern.search(para.text):
                doc_changed = True
                # 使用 finditer 获取所有匹配项及其位置
                matches = list(url_pattern.finditer(para.text))

                # 将段落文本按图片链接分割成文本和图片两部分
                segments = []
                last_end = 0
                for match in matches:
                    # 添加链接前的文本部分
                    segments.append(('text', para.text[last_end:match.start()]))
                    # 添加图片URL部分
                    segments.append(('image', match.group(2)))  # group(2) is the URL
                    last_end = match.end()
                # 添加最后一个链接后的文本部分
                segments.append(('text', para.text[last_end:]))

                # 清空原段落的所有run，但保留段落本身的格式
                for run in para.runs:
                    run.clear()

                log_callback(f"  -> 在段落中找到链接，开始重建段落...")

                # 重建段落
                for seg_type, content in segments:
                    if seg_type == 'text':
                        if content:  # 仅在有内容时添加
                            para.add_run(content)
                            log_callback(f"    -> 添加文本: '{content}'")

                    elif seg_type == 'image':
                        images_found += 1
                        log_callback(f"    -> 准备处理图片链接: {content}")
                        try:
                            log_callback("      -> 正在下载图片...")
                            response = requests.get(content, timeout=20, headers={'User-Agent': 'Mozilla/5.0'})
                            response.raise_for_status()
                            image_stream = io.BytesIO(response.content)
                            log_callback("      -> 图片下载成功。")

                            # 在段落中添加图片
                            run = para.add_run()
                            run.add_picture(image_stream, width=Inches(2.5))
                            log_callback("      -> 图片成功插入段落。")

                        except requests.exceptions.RequestException as e:
                            log_callback(f"      -> [错误] 图片下载失败: {e}")
                            # 下载失败时，将原链接文本加回去
                            para.add_run(f"![下载失败]({content})")
                        except Exception as e:
                            log_callback(f"      -> [错误] 插入图片时出错: {e}")
                            para.add_run(f"![插入失败]({content})")

        if not doc_changed:
            log_callback("  -> 未在文档中找到符合格式的图片链接。")
            return "no_images_found"

        base, ext = os.path.splitext(doc_path)
        new_doc_path = f"{base}_processed.docx"
        log_callback(f"  -> 处理完成，正在保存到新文件: {os.path.basename(new_doc_path)}")
        doc.save(new_doc_path)
        return new_doc_path

    except Exception as e:
        log_callback(f"  -> [严重错误] 处理文档时发生意外: {e}")
        import traceback
        log_callback(traceback.format_exc())
        return None


# --- Tkinter GUI Application (界面部分无需改动) ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word 图片链接批量替换工具 (精确定位版)")
        self.root.geometry("800x650")
        self.file_paths = []

        # --- 设置样式 ---
        self.root.configure(bg="#f0f0f0")
        font_main = ("Arial", 12)
        font_log = ("Courier New", 10)

        # --- 创建主框架 ---
        main_frame = tk.Frame(root, padx=20, pady=20, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 文件选择部分 ---
        select_frame = tk.Frame(main_frame, bg="#f0f0f0")
        select_frame.pack(fill=tk.X, pady=(0, 10))

        self.select_button = tk.Button(select_frame, text="1. 选择Word文档 (可多选)", command=self.select_files,
                                       font=font_main, bg="#0078d4", fg="white", relief=tk.FLAT, padx=10, pady=5)
        self.select_button.pack(side=tk.LEFT, ipadx=10)

        self.clear_button = tk.Button(select_frame, text="清空列表", command=self.clear_list, font=font_main,
                                      bg="#dc3545", fg="white", relief=tk.FLAT, padx=10, pady=5)
        self.clear_button.pack(side=tk.LEFT, padx=(10, 0))

        # --- 文件列表显示 ---
        list_frame = tk.LabelFrame(main_frame, text="待处理文件列表", font=font_main, padx=10, pady=10, bg="#f0f0f0")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.file_listbox = Listbox(list_frame, font=("Arial", 11), relief=tk.SOLID, bd=1)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # --- 处理按钮 ---
        self.process_button = tk.Button(main_frame, text="2. 开始批量处理", command=self.start_processing_thread,
                                        font=font_main, bg="#28a745", fg="white", state=tk.DISABLED, relief=tk.FLAT,
                                        padx=10, pady=5)
        self.process_button.pack(fill=tk.X, pady=10, ipady=5)

        # --- 日志输出区域 ---
        log_frame = tk.LabelFrame(main_frame, text="处理日志", font=font_main, padx=10, pady=10, bg="#f0f0f0")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED, font=font_log,
                                                  relief=tk.SOLID, bd=1, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, message):
        """线程安全地向日志区域添加消息"""

        def _log():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        self.root.after(0, _log)

    def select_files(self):
        """打开文件对话框以选择多个.docx文件"""
        paths = filedialog.askopenfilenames(
            title="请选择一个或多个Word文档",
            filetypes=[("Word Documents", "*.docx")]
        )
        if paths:
            # 防止重复添加
            for p in paths:
                if p not in self.file_paths:
                    self.file_paths.append(p)
            self.update_file_listbox()
            self.log(f"添加了 {len(paths)} 个文件到处理列表。当前共 {len(self.file_paths)} 个。")

    def update_file_listbox(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, END)
        for path in self.file_paths:
            self.file_listbox.insert(END, os.path.basename(path))

        if self.file_paths:
            self.process_button.config(state=tk.NORMAL)
        else:
            self.process_button.config(state=tk.DISABLED)

    def clear_list(self):
        """清空已选择的文件列表"""
        self.file_paths = []
        self.update_file_listbox()
        self.log("文件列表已清空。")

    def start_processing_thread(self):
        """在单独的线程中开始处理文件以避免GUI冻结"""
        if not self.file_paths:
            messagebox.showerror("错误", "请先选择至少一个Word文档！")
            return

        self.process_button.config(state=tk.DISABLED, text="正在处理中...")
        self.select_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)

        thread = threading.Thread(target=self.process_worker, daemon=True)
        thread.start()

    def process_worker(self):
        """工作线程，循环处理所有选中的文档"""
        total_files = len(self.file_paths)
        success_count = 0
        fail_count = 0

        # 创建一个副本进行迭代，这样可以在处理后清空列表
        paths_to_process = list(self.file_paths)

        for i, doc_path in enumerate(paths_to_process):
            self.log("=" * 60)
            self.log(f"开始处理文件 {i + 1}/{total_files}: {os.path.basename(doc_path)}")

            new_file_path = find_and_replace_images_in_doc(doc_path, self.log)

            if new_file_path and new_file_path != "no_images_found":
                success_count += 1
            elif new_file_path is None:
                fail_count += 1

        self.log("=" * 60)
        self.log(f"批量处理完成！")
        summary = f"共处理 {total_files} 个文件。\n\n成功生成: {success_count} 个\n处理失败: {fail_count} 个"
        self.log(summary.replace('\n\n', '\n'))
        messagebox.showinfo("批量处理完成", summary)

        def _reset_ui():
            self.clear_list()  # 处理完成后清空列表
            self.process_button.config(state=tk.DISABLED, text="2. 开始批量处理")  # 列表清空后禁用
            self.select_button.config(state=tk.NORMAL)
            self.clear_button.config(state=tk.NORMAL)

        self.root.after(0, _reset_ui)


if __name__ == "__main__":
    try:
        import docx
        import requests
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "缺少依赖库",
            "运行本程序需要'python-docx'和'requests'库。\n请在命令行中使用以下命令安装:\n\npip install python-docx requests"
        )
        exit()

    root = tk.Tk()
    app = App(root)
    root.mainloop()
