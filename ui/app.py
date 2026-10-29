"""
Main Application Window for TranHR Tax Tool.
"""

import tkinter as tk
from tkinter import ttk
from .components import ToolTab
from core.mst_processor import process_mst
from core.npt_processor import process_npt


class TaxToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Công Cụ Tự Động Đối Soát Thuế (MST & NPT)")
        self.geometry("860x720")
        self.minsize(780, 600)

        self._apply_style()
        self._build_layout()
        self._center_window()

    def _apply_style(self):
        style = ttk.Style(self)
        try:
            # Ưu tiên theme clam hoặc vista/aqua
            available = style.theme_names()
            if "clam" in available:
                style.theme_use("clam")
        except Exception:
            pass

        # Cấu hình font và khoảng đệm cho Tab Notebook
        style.configure(
            "TNotebook.Tab",
            padding=[18, 8],
            font=("Helvetica", 11, "bold")
        )

    def _build_layout(self):
        # 1. Header Frame
        header = tk.Frame(self, bg="#1E3A8A", height=60)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        title_lbl = tk.Label(
            header,
            text="TAX AUTOMATION TOOL",
            font=("Helvetica", 14, "bold"),
            fg="#FFFFFF",
            bg="#1E3A8A",
            padx=16
        )
        title_lbl.pack(side=tk.LEFT, fill=tk.Y)

        sub_lbl = tk.Label(
            header,
            text="Tự động bóc tách Mẫu 14 & Mẫu 40 từ Cục Thuế",
            font=("Helvetica", 10, "italic"),
            fg="#93C5FD",
            bg="#1E3A8A",
            padx=16
        )
        sub_lbl.pack(side=tk.RIGHT, fill=tk.Y)

        # 2. Notebook (Tabs)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        # Tab 1: MST
        self.tab_mst = ToolTab(
            parent=self.notebook,
            tool_id="mst",
            title="Mã Số Thuế Cá Nhân",
            description=(
                "• Nghiệp vụ: Cấp Mã số thuế cho Người Lao Động.\n"
                "• Quy tắc: Đọc số CCCD tại CỘT M (từ dòng 4), điền MST vào CỘT D.\n"
                "• Nhận diện: Mẫu 14-MST (Cấp mới - Tô xanh) & Mẫu 40/TB (Trùng MST - Tô vàng)."
            ),
            default_excel_hint="Tong hop MST.xlsx",
            default_output_name="Tong_hop_MST_Updated.xlsx",
            process_fn=process_mst
        )
        self.notebook.add(self.tab_mst, text="  🏢 Cấp Mã Số Thuế (MST)  ")

        # Tab 2: NPT
        self.tab_npt = ToolTab(
            parent=self.notebook,
            tool_id="npt",
            title="Người Phụ Thuộc",
            description=(
                "• Nghiệp vụ: Cấp Mã số thuế cho Người Phụ Thuộc (giảm trừ gia cảnh).\n"
                "• Quy tắc: Đọc số CCCD tại CỘT K (từ dòng 4), điền MST vào CỘT D.\n"
                "• Tính năng đặc biệt: Thuật toán Split-match chống lỗi ngắt trang (Page Break) trên file PDF."
            ),
            default_excel_hint="Tong hop NPT.xlsx",
            default_output_name="Tong_hop_NPT_Updated.xlsx",
            process_fn=process_npt
        )
        self.notebook.add(self.tab_npt, text="  👨‍👩‍👧 Người Phụ Thuộc (NPT)  ")

        # 3. Footer / Status bar
        footer = tk.Frame(self, bg="#E2E8F0", height=24)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)

        footer_lbl = tk.Label(
            footer,
            text="Phiên bản desktop 2.0 • Sẵn sàng xử lý dữ liệu",
            font=("Helvetica", 9),
            fg="#475569",
            bg="#E2E8F0",
            padx=10
        )
        footer_lbl.pack(side=tk.LEFT)

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, int((sw - w) / 2))
        y = max(0, int((sh - h) / 2))
        self.geometry(f"{w}x{h}+{x}+{y}")
