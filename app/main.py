from __future__ import annotations

import queue
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from converter import ConversionError, PdfToWordConverter


class PdfToWordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF 转可编辑 Word")
        self.minsize(720, 500)
        self.geometry("800x560")
        self.files: list[Path] = []
        self.output_dir = tk.StringVar(value=str(Path.cwd() / "output"))
        self.status = tk.StringVar(value="请选择 PDF 文件。")
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self._build_ui()
        self.after(100, self._drain_events)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=18)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        ttk.Label(root, text="PDF 转可编辑 Word", font=("Microsoft YaHei", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(root, text="本地处理，不上传文件；复杂表格将重建为可编辑 Word 表格。", foreground="#555555").grid(row=0, column=0, sticky="e")

        file_frame = ttk.LabelFrame(root, text="待转换 PDF", padding=10)
        file_frame.grid(row=1, column=0, sticky="nsew", pady=(16, 10))
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(0, weight=1)
        self.file_list = tk.Listbox(file_frame, activestyle="none", selectmode=tk.EXTENDED)
        self.file_list.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_list.configure(yscrollcommand=scrollbar.set)
        buttons = ttk.Frame(file_frame)
        buttons.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(buttons, text="选择 PDF", command=self.choose_files).pack(side=tk.LEFT)
        ttk.Button(buttons, text="移除选中项", command=self.remove_selected).pack(side=tk.LEFT, padx=8)
        ttk.Button(buttons, text="清空", command=self.clear_files).pack(side=tk.LEFT)

        output = ttk.Frame(root)
        output.grid(row=2, column=0, sticky="ew", pady=8)
        output.columnconfigure(1, weight=1)
        ttk.Label(output, text="输出文件夹：").grid(row=0, column=0, sticky="w")
        ttk.Entry(output, textvariable=self.output_dir).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(output, text="选择", command=self.choose_output_dir).grid(row=0, column=2)
        ttk.Button(output, text="打开", command=self.open_output_dir).grid(row=0, column=3, padx=(8, 0))

        bottom = ttk.Frame(root)
        bottom.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.status, foreground="#444444").grid(row=0, column=0, sticky="w")
        self.convert_button = ttk.Button(bottom, text="开始转换", command=self.start_conversion)
        self.convert_button.grid(row=0, column=1, sticky="e")

    def choose_files(self) -> None:
        selected = filedialog.askopenfilenames(title="选择 PDF 文件", filetypes=[("PDF 文件", "*.pdf")])
        for name in selected:
            path = Path(name)
            if path not in self.files:
                self.files.append(path)
        self._refresh_files()

    def remove_selected(self) -> None:
        selected = set(self.file_list.curselection())
        self.files = [path for index, path in enumerate(self.files) if index not in selected]
        self._refresh_files()

    def clear_files(self) -> None:
        self.files.clear()
        self._refresh_files()

    def choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择输出文件夹", initialdir=self.output_dir.get())
        if selected:
            self.output_dir.set(selected)


    def open_output_dir(self) -> None:
        output = Path(self.output_dir.get()).expanduser()
        try:
            output.mkdir(parents=True, exist_ok=True)
            os.startfile(output)  # type: ignore[attr-defined]
        except OSError as exc:
            messagebox.showerror("无法打开输出文件夹", str(exc))
    def _refresh_files(self) -> None:
        self.file_list.delete(0, tk.END)
        for path in self.files:
            self.file_list.insert(tk.END, str(path))
        self.status.set(f"已选择 {len(self.files)} 个 PDF 文件。" if self.files else "请选择 PDF 文件。")

    def start_conversion(self) -> None:
        if not self.files:
            messagebox.showwarning("没有文件", "请先选择至少一个 PDF 文件。")
            return
        output = Path(self.output_dir.get()).expanduser()
        self.convert_button.configure(state=tk.DISABLED)
        threading.Thread(target=self._convert_worker, args=(list(self.files), output), daemon=True).start()

    def _convert_worker(self, files: list[Path], output: Path) -> None:
        converter = PdfToWordConverter()
        succeeded, failures = 0, []
        for source in files:
            destination = self._available_destination(output, source.stem)
            try:
                converter.convert(source, destination, lambda message: self.events.put(("progress", message)))
                succeeded += 1
            except ConversionError as exc:
                failures.append(f"{source.name}：{exc}")
        if failures:
            message = f"已完成 {succeeded}/{len(files)} 个文件。\n\n" + "\n".join(failures)
            self.events.put(("error", message))
        else:
            self.events.put(("done", f"已完成 {succeeded} 个文件，输出位置：{output}"))

    @staticmethod
    def _available_destination(output: Path, stem: str) -> Path:
        output.mkdir(parents=True, exist_ok=True)
        candidate = output / f"{stem}.docx"
        counter = 2
        while candidate.exists():
            candidate = output / f"{stem} ({counter}).docx"
            counter += 1
        return candidate

    def _drain_events(self) -> None:
        try:
            while True:
                kind, message = self.events.get_nowait()
                if kind == "progress":
                    self.status.set(message)
                elif kind == "done":
                    self.status.set(message)
                    messagebox.showinfo("转换完成", message)
                    self.convert_button.configure(state=tk.NORMAL)
                else:
                    self.status.set("转换完成，但有文件失败。")
                    messagebox.showerror("转换未完全完成", message)
                    self.convert_button.configure(state=tk.NORMAL)
        except queue.Empty:
            pass
        self.after(100, self._drain_events)


if __name__ == "__main__":
    PdfToWordApp().mainloop()
