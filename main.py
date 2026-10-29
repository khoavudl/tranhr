#!/usr/bin/env python3
"""
TranHR - Desktop Tax Automation Application (Tkinter)
Entry point to run the application.
"""

import sys
import os

# Đảm bảo đường dẫn thư mục gốc có trong sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


def check_dependencies():
    """Kiểm tra các thư viện bắt buộc."""
    missing = []
    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz
        except ImportError:
            missing.append("pymupdf")

    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")

    if missing:
        print("=" * 60)
        print("THIẾU THƯ VIỆN PHỤ THUỘC:")
        print(f"Vui lòng cài đặt bằng lệnh: pip install {' '.join(missing)}")
        print("=" * 60)
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Thiếu thư viện",
                f"Vui lòng cài đặt các thư viện sau trước khi chạy:\n\npip install {' '.join(missing)}"
            )
            root.destroy()
        except Exception:
            pass
        sys.exit(1)


def main():
    check_dependencies()
    from ui import TaxToolApp
    app = TaxToolApp()
    app.mainloop()


if __name__ == "__main__":
    main()
