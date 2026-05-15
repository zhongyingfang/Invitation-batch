#!/usr/bin/env python3
"""图形界面入口：用于选择多个 Excel 文件和模板映射并运行批量生成。"""

import os
import sys
import threading
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
except ImportError:
    raise RuntimeError("未找到 tkinter 模块，请在 Windows 上安装 Python 时勾选 Tkinter 组件。")

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR / "src"))

from batch_processor import BatchProcessor


class InviteGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("批量邀请函生成器")
        self.root.geometry("820x660")
        self.root.resizable(False, False)

        self.mappings = []
        self.output_docs_path = tk.StringVar(value=str(SCRIPT_DIR / "output_documents"))
        self.output_images_path = tk.StringVar(value=str(SCRIPT_DIR / "output_images"))

        self._build_ui()

    def _build_ui(self):
        frame = tk.Frame(self.root, padx=16, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        row = 0
        tk.Label(frame, text="Excel + 模板映射列表:", anchor="w").grid(row=row, column=0, sticky="w", columnspan=4)

        row += 1
        self.mapping_listbox = tk.Listbox(frame, width=110, height=10, selectmode=tk.EXTENDED)
        self.mapping_listbox.grid(row=row, column=0, columnspan=4, sticky="nsew")
        mapping_scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=self.mapping_listbox.yview)
        mapping_scrollbar.grid(row=row, column=4, sticky="ns")
        self.mapping_listbox.config(yscrollcommand=mapping_scrollbar.set)

        row += 1
        tk.Button(frame, text="新增映射", width=14, command=self.add_mapping).grid(row=row, column=0, pady=(8, 0), sticky="w")
        tk.Button(frame, text="删除选中", width=14, command=self.remove_selected_mappings).grid(row=row, column=1, pady=(8, 0), sticky="w")
        tk.Button(frame, text="清空映射", width=14, command=self.clear_mappings).grid(row=row, column=2, pady=(8, 0), sticky="w")

        row += 1
        tk.Label(frame, text="输出文档目录:", anchor="w").grid(row=row, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.output_docs_path, width=70).grid(row=row, column=1, columnspan=2, sticky="we", padx=(8, 0))
        tk.Button(frame, text="浏览...", command=self.select_output_docs).grid(row=row, column=3, padx=(8, 0))

        row += 1
        tk.Label(frame, text="输出图片目录:", anchor="w").grid(row=row, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.output_images_path, width=70).grid(row=row, column=1, columnspan=2, sticky="we", padx=(8, 0))
        tk.Button(frame, text="浏览...", command=self.select_output_images).grid(row=row, column=3, padx=(8, 0))

        row += 1
        self.generate_png_var = tk.BooleanVar(value=True)
        self.png_check = tk.Checkbutton(frame, text="生成 PNG 图片（需安装 LibreOffice，未安装请取消勾选）",
                                         variable=self.generate_png_var)
        self.png_check.grid(row=row, column=0, columnspan=3, pady=(12, 0), sticky="w")

        row += 1
        self.run_button = tk.Button(frame, text="开始生成", width=16, command=self.start_processing, bg="#4CAF50", fg="white")
        self.run_button.grid(row=row, column=0, pady=(12, 0), sticky="w")
        tk.Button(frame, text="打开文档输出目录", command=self.open_output_docs).grid(row=row, column=1, pady=(12, 0), sticky="w")
        tk.Button(frame, text="打开图片输出目录", command=self.open_output_images).grid(row=row, column=2, pady=(12, 0), sticky="w")

        row += 1
        tk.Label(frame, text="执行日志:", anchor="w").grid(row=row, column=0, columnspan=4, sticky="w", pady=(16, 0))

        row += 1
        self.log_text = scrolledtext.ScrolledText(frame, state="disabled", width=110, height=14, wrap=tk.WORD)
        self.log_text.grid(row=row, column=0, columnspan=4, pady=(4, 0), sticky="nsew")

        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=0)

    def add_mapping(self):
        excel_paths = filedialog.askopenfilenames(
            title="选择 Excel 数据文件（可多选）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")]
        )
        if not excel_paths:
            return

        # 询问是否选择 PPTX 模板
        use_pptx = messagebox.askyesno("PPTX 模板", "是否选择 PPTX 模板？\n\n选「是」继续选择 PPTX 文件\n选「否」跳过 PowerPoint 模板")
        pptx_path = None
        if use_pptx:
            pptx_path = filedialog.askopenfilename(
                title="选择 PPTX 模板文件",
                filetypes=[("PowerPoint 文件", "*.pptx")]
            )
            if not pptx_path:
                # 用户取消了，再问一次是否要跳过
                if not messagebox.askyesno("跳过 PPTX", "未选择 PPTX 模板，是否跳过（不生成 PPTX 邀请函）？"):
                    return

        # 询问是否选择 DOCX 模板
        use_docx = messagebox.askyesno("DOCX 模板", "是否选择 DOCX 模板？\n\n选「是」继续选择 DOCX 文件\n选「否」跳过 Word 模板")
        docx_path = None
        if use_docx:
            docx_path = filedialog.askopenfilename(
                title="选择 DOCX 模板文件",
                filetypes=[("Word 文件", "*.docx")]
            )
            if not docx_path:
                if not messagebox.askyesno("跳过 DOCX", "未选择 DOCX 模板，是否跳过（不生成 DOCX 邀请函）？"):
                    return

        if not pptx_path and not docx_path:
            messagebox.showwarning("输入错误", "请至少选择一种模板（PPTX 或 DOCX）。")
            return

        for excel_path in excel_paths:
            mapping = {
                "excel": excel_path,
                "pptx": pptx_path,
                "docx": docx_path,
            }
            self.mappings.append(mapping)
            self.mapping_listbox.insert(tk.END, self._format_mapping(mapping))

        count = len(excel_paths)
        types = []
        if pptx_path:
            types.append("PPTX")
        if docx_path:
            types.append("DOCX")
        self.log(f"已添加 {count} 个 Excel + {'/'.join(types)} 映射。\n")

    def remove_selected_mappings(self):
        selected = list(self.mapping_listbox.curselection())
        if not selected:
            messagebox.showinfo("提示", "请先选择要删除的映射。")
            return

        for index in reversed(selected):
            del self.mappings[index]
            self.mapping_listbox.delete(index)

        self.log(f"已删除 {len(selected)} 个映射。\n")

    def clear_mappings(self):
        self.mappings.clear()
        self.mapping_listbox.delete(0, tk.END)
        self.log("已清空所有映射。\n")

    def _format_mapping(self, mapping):
        excel_name = Path(mapping["excel"]).name
        pptx = Path(mapping["pptx"]).name if mapping["pptx"] else "(未选择)"
        docx = Path(mapping["docx"]).name if mapping["docx"] else "(未选择)"
        return f"Excel: {excel_name} | PPTX: {pptx} | DOCX: {docx}"

    def select_output_docs(self):
        path = filedialog.askdirectory(title="选择输出文档目录")
        if path:
            self.output_docs_path.set(path)
            self.log(f"已设置文档输出目录: {path}\n")

    def select_output_images(self):
        path = filedialog.askdirectory(title="选择输出图片目录")
        if path:
            self.output_images_path.set(path)
            self.log(f"已设置图片输出目录: {path}\n")

    def validate_mappings(self):
        if not self.mappings:
            messagebox.showwarning("输入错误", "请先添加至少一个 Excel 模板映射。")
            return False

        for mapping in self.mappings:
            for key, label in [("excel", "Excel 文件"), ("pptx", "PPTX 模板"), ("docx", "DOCX 模板")]:
                path_value = mapping.get(key)
                if path_value and not Path(path_value).exists():
                    messagebox.showwarning("文件不存在", f"{label} 不存在：{path_value}")
                    return False

        if not self.output_docs_path.get():
            messagebox.showwarning("输入错误", "请先设置输出文档目录。")
            return False
        if not self.output_images_path.get():
            messagebox.showwarning("输入错误", "请先设置输出图片目录。")
            return False

        return True

    def start_processing(self):
        if not self.validate_mappings():
            return

        self.run_button.config(state="disabled")
        png = self.generate_png_var.get()
        self.log(f"开始批量生成（{'生成' if png else '跳过'} PNG），请稍候...\n")
        thread = threading.Thread(target=self._run_processor, daemon=True)
        thread.start()

    def _run_processor(self):
        try:
            output_docs_dir = Path(self.output_docs_path.get())
            output_images_dir = Path(self.output_images_path.get())
            output_docs_dir.mkdir(parents=True, exist_ok=True)
            output_images_dir.mkdir(parents=True, exist_ok=True)

            generate_png = self.generate_png_var.get()

            for idx, mapping in enumerate(self.mappings, start=1):
                types = []
                if mapping["pptx"]:
                    types.append("PPTX")
                if mapping["docx"]:
                    types.append("DOCX")

                self.log(f"\n正在处理映射 {idx}/{len(self.mappings)}:\n")
                self.log(f"  Excel: {mapping['excel']}\n")
                self.log(f"  PPTX: {mapping['pptx'] or '(跳过)'}\n")
                self.log(f"  DOCX: {mapping['docx'] or '(跳过)'}\n")

                processor = BatchProcessor(
                    mapping["excel"],
                    mapping["pptx"],
                    mapping["docx"],
                    str(output_docs_dir),
                    str(output_images_dir)
                )
                self.log("  读取 Excel 数据...\n")
                processor.read_data()
                self.log(f"  读取完成：{len(processor.data)} 条记录。\n")
                processor.process(generate_png=generate_png)

            self.log("\n所有映射处理完成。输出已保存到指定目录。\n")
            messagebox.showinfo("完成", "批量生成已完成！")
        except Exception as exc:
            self.log(f"错误：{exc}\n")
            messagebox.showerror("执行失败", str(exc))
        finally:
            self.run_button.config(state="normal")

    def log(self, message: str):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def open_output_docs(self):
        path = self.output_docs_path.get()
        if path and Path(path).exists():
            self._open_folder(path)

    def open_output_images(self):
        path = self.output_images_path.get()
        if path and Path(path).exists():
            self._open_folder(path)

    def _open_folder(self, path: str):
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            os.system(f"open \"{path}\"")
        else:
            os.system(f"xdg-open \"{path}\"")


def main():
    root = tk.Tk()
    app = InviteGeneratorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
