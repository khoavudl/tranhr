"""
UI Helper components and utilities for TranHR Tax Tool.
"""

import os
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from typing import Callable, List, Optional, Dict, Any


def open_path(path: str):
    """Mở file hoặc thư mục trên hệ điều hành (macOS, Windows, Linux)."""
    if not os.path.exists(path):
        return
    sys_name = platform.system()
    try:
        if sys_name == "Darwin":  # macOS
            subprocess.run(["open", path])
        elif sys_name == "Windows":
            os.startfile(path)
        else:  # Linux
            subprocess.run(["xdg-open", path])
    except Exception as e:
        print(f"Không thể mở đường dẫn: {e}")


def open_parent_folder(file_path: str):
    """Mở thư mục chứa file."""
    if not file_path:
        return
    folder = os.path.dirname(file_path) if os.path.isfile(file_path) else file_path
    open_path(folder)


class ToolTab(ttk.Frame):
    """
    Component Tab dùng chung cho cả tool MST và NPT.
    """
    def __init__(
        self,
        parent,
        tool_id: str,
        title: str,
        description: str,
        default_excel_hint: str,
        default_output_name: str,
        process_fn: Callable[..., Dict[str, Any]]
    ):
        super().__init__(parent, padding=12)
        self.tool_id = tool_id
        self.title = title
        self.description = description
        self.default_excel_hint = default_excel_hint
        self.default_output_name = default_output_name
        self.process_fn = process_fn

        self.excel_path = tk.StringVar()
        self.pdf_files: List[str] = []
        self.last_output_file: Optional[str] = None
        self.is_processing = False

        self._build_ui()

    def _build_ui(self):
        # 1. Header info banner
        banner = ttk.LabelFrame(self, text="ℹ️ Thông tin công cụ", padding=8)
        banner.pack(fill=tk.X, pady=(0, 10))

        lbl_desc = ttk.Label(
            banner,
            text=self.description,
            font=("Helvetica", 11),
            foreground="#2B2D42",
            wraplength=720
        )
        lbl_desc.pack(anchor=tk.W)

        # 2. Excel Selection
        excel_frame = ttk.LabelFrame(self, text="1. Chọn file Excel danh sách", padding=8)
        excel_frame.pack(fill=tk.X, pady=(0, 10))

        excel_box = ttk.Frame(excel_frame)
        excel_box.pack(fill=tk.X)

        self.entry_excel = ttk.Entry(excel_box, textvariable=self.excel_path, font=("Helvetica", 10))
        self.entry_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_browse_excel = ttk.Button(
            excel_box,
            text="📂 Chọn file Excel...",
            command=self._choose_excel
        )
        btn_browse_excel.pack(side=tk.RIGHT)

        # 3. PDF Selection
        pdf_frame = ttk.LabelFrame(self, text="2. Chọn file PDF thông báo từ Cục Thuế (Mẫu 14 & Mẫu 40)", padding=8)
        pdf_frame.pack(fill=tk.X, pady=(0, 10))

        pdf_actions = ttk.Frame(pdf_frame)
        pdf_actions.pack(fill=tk.X, pady=(0, 6))

        btn_choose_pdfs = ttk.Button(
            pdf_actions,
            text="📄 Chọn nhiều file PDF...",
            command=self._choose_pdfs
        )
        btn_choose_pdfs.pack(side=tk.LEFT, padx=(0, 8))

        btn_choose_folder = ttk.Button(
            pdf_actions,
            text="📁 Chọn cả thư mục chứa PDF...",
            command=self._choose_pdf_folder
        )
        btn_choose_folder.pack(side=tk.LEFT, padx=(0, 8))

        btn_clear_pdfs = ttk.Button(
            pdf_actions,
            text="🗑️ Xóa danh sách",
            command=self._clear_pdfs
        )
        btn_clear_pdfs.pack(side=tk.LEFT)

        self.lbl_pdf_count = ttk.Label(
            pdf_actions,
            text="Chưa chọn file PDF nào",
            font=("Helvetica", 10, "italic"),
            foreground="#6C757D"
        )
        self.lbl_pdf_count.pack(side=tk.RIGHT)

        # 4. Action bar & Progress
        action_frame = ttk.Frame(self, padding=(0, 5))
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_run = tk.Button(
            action_frame,
            text="🚀 BẮT ĐẦU XỬ LÝ",
            font=("Helvetica", 11, "bold"),
            bg="#2E7D32",
            fg="white",
            activebackground="#1B5E20",
            activeforeground="white",
            padx=16,
            pady=6,
            cursor="hand2",
            relief=tk.RAISED,
            command=self._start_process
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, 12))

        self.btn_open_file = ttk.Button(
            action_frame,
            text="📊 Mở file Excel kết quả",
            state=tk.DISABLED,
            command=lambda: open_path(self.last_output_file) if self.last_output_file else None
        )
        self.btn_open_file.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_folder = ttk.Button(
            action_frame,
            text="📂 Mở thư mục kết quả",
            state=tk.DISABLED,
            command=lambda: open_parent_folder(self.last_output_file) if self.last_output_file else None
        )
        self.btn_open_folder.pack(side=tk.LEFT)

        self.progress_bar = ttk.Progressbar(action_frame, mode="determinate")
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(12, 0))

        # 5. Log Window
        log_frame = ttk.LabelFrame(self, text="3. Tiến trình & Nhật ký (Log)", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(
            log_frame,
            wrap=tk.WORD,
            font=("Menlo" if platform.system() == "Darwin" else "Consolas", 10),
            bg="#1E1E24",
            fg="#E0E1DD",
            insertbackground="white",
            relief=tk.FLAT,
            height=12
        )
        scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll.set)

        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._log(f"Sẵn sàng. Vui lòng chọn file Excel ({self.default_excel_hint}) và các file PDF kết quả từ Cục Thuế.")

    def _log(self, text: str):
        self.txt_log.insert(tk.END, text + "\n")
        self.txt_log.see(tk.END)

    def _choose_excel(self):
        file_path = filedialog.askopenfilename(
            title=f"Chọn file {self.default_excel_hint}",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path.set(file_path)
            self._log(f"✓ Đã chọn file Excel: {file_path}")

    def _choose_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Chọn các file PDF thông báo thuế",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if files:
            # Gộp và loại bỏ trùng lặp
            new_files = list(set(self.pdf_files + list(files)))
            self.pdf_files = new_files
            self._update_pdf_status()
            self._log(f"✓ Đã chọn thêm {len(files)} file PDF (Tổng cộng: {len(self.pdf_files)} file)")

    def _choose_pdf_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa các file PDF")
        if folder:
            found = []
            for root, _, filenames in os.walk(folder):
                for f in filenames:
                    if f.lower().endswith(".pdf"):
                        found.append(os.path.join(root, f))
            if found:
                new_files = list(set(self.pdf_files + found))
                self.pdf_files = new_files
                self._update_pdf_status()
                self._log(f"✓ Đã quét thư mục '{folder}': tìm thấy {len(found)} file PDF (Tổng cộng: {len(self.pdf_files)} file)")
            else:
                messagebox.showwarning("Thông báo", f"Không tìm thấy file .pdf nào trong thư mục:\n{folder}")

    def _clear_pdfs(self):
        self.pdf_files.clear()
        self._update_pdf_status()
        self._log("✓ Đã xóa danh sách file PDF.")

    def _update_pdf_status(self):
        count = len(self.pdf_files)
        if count == 0:
            self.lbl_pdf_count.config(text="Chưa chọn file PDF nào", foreground="#6C757D")
        else:
            self.lbl_pdf_count.config(text=f"Đã chọn {count} file PDF", foreground="#2E7D32")

    def _start_process(self):
        excel = self.excel_path.get().strip()
        if not excel or not os.path.exists(excel):
            messagebox.showerror("Thiếu dữ liệu", "Vui lòng chọn file Excel hợp lệ trước khi bấm chạy!")
            return

        if not self.pdf_files:
            messagebox.showerror("Thiếu dữ liệu", "Vui lòng chọn ít nhất 1 file PDF thông báo từ Cục Thuế!")
            return

        if self.is_processing:
            return

        self.is_processing = True
        self.btn_run.config(text="⏳ Đang xử lý...", state=tk.DISABLED, bg="#757575")
        self.btn_open_file.config(state=tk.DISABLED)
        self.btn_open_folder.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0

        # Chạy trong luồng nền
        thread = threading.Thread(target=self._run_worker, daemon=True)
        thread.start()

    def _run_worker(self):
        excel = self.excel_path.get().strip()
        pdf_list = list(self.pdf_files)

        def log_cb(msg: str):
            self.after(0, lambda: self._log(msg))

        def progress_cb(cur: int, total: int):
            percent = int((cur / total) * 100) if total > 0 else 0
            self.after(0, lambda: self.progress_bar.configure(value=percent))

        try:
            res = self.process_fn(
                excel_file=excel,
                pdf_files=pdf_list,
                output_file=None,
                log_callback=log_cb,
                progress_callback=progress_cb
            )

            self.last_output_file = res["output_file"]

            def on_success():
                self.btn_run.config(text="🚀 BẮT ĐẦU XỬ LÝ", state=tk.NORMAL, bg="#2E7D32")
                self.btn_open_file.config(state=tk.NORMAL)
                self.btn_open_folder.config(state=tk.NORMAL)
                self.progress_bar["value"] = 100
                self.is_processing = False

                split_info = f"\n- Khớp ngắt trang: {res.get('split_matches', 0)}" if 'split_matches' in res else ""
                msg = (
                    f"Xử lý thành công!\n\n"
                    f"• File kết quả: {os.path.basename(res['output_file'])}\n"
                    f"• Tổng số dòng: {res['total_rows']}\n"
                    f"• Cập nhật thành công: {res['updated']}\n"
                    f"  - Cấp mới (Xanh): {res['cap_moi']}{split_info}\n"
                    f"  - Trùng MST (Vàng): {res['duplicate']}\n"
                    f"• Không tìm thấy: {res['not_found']}"
                )
                messagebox.showinfo("Thành công", msg)

            self.after(0, on_success)

        except Exception as e:
            def on_error(err=e):
                self.btn_run.config(text="🚀 BẮT ĐẦU XỬ LÝ", state=tk.NORMAL, bg="#2E7D32")
                self.is_processing = False
                log_cb(f"\n❌ LỖI TRONG QUÁ TRÌNH XỬ LÝ: {err}")
                messagebox.showerror("Lỗi thực thi", f"Đã xảy ra lỗi:\n{err}")

            self.after(0, on_error)
