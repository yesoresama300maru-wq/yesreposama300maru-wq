"""Desktop app for creating Excel minutes from a transcript.

Run with: python app.py
"""

from __future__ import annotations

import importlib
import importlib.util
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from minutes_maker import save_minutes_excel

DND_AVAILABLE = False

try:
    tkinterdnd2 = importlib.import_module("tkinterdnd2")
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False


class MinutesApp:
    def __init__(self) -> None:
        root_class = tkinterdnd2.TkinterDnD.Tk if DND_AVAILABLE else tk.Tk
        self.root = root_class()
        self.root.title("トランスクリプト議事録メーカー")
        self.root.geometry("820x620")
        self.output_path = tk.StringVar(value=str(Path.cwd() / "minutes.xlsx"))
        self.status = tk.StringVar(value="テキストを貼り付けるか、.txtファイルをドラッグ＆ドロップしてください。")
        self._build_ui()
        self._setup_dnd()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)

        title = ttk.Label(frame, text="トランスクリプトからExcel議事録を作成", font=("Helvetica", 16, "bold"))
        title.pack(anchor="w")

        help_text = (
            "1. 文字起こし文章を貼り付け、または.txtファイルをドラッグ＆ドロップ\n"
            "2. 出力先を選択\n"
            "3. 「Excel議事録を作成」を押してください"
        )
        ttk.Label(frame, text=help_text).pack(anchor="w", pady=(8, 12))

        self.text = tk.Text(frame, wrap="word", height=22, undo=True)
        self.text.pack(fill="both", expand=True)
        self.text.insert("1.0", "会議名：サンプル定例会\n開催日：2026年5月30日\n参加者：田中、佐藤\n議題：進捗確認\n決定：次回までに仕様を確定する\n担当：田中 仕様書を更新（6/5）")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=12)
        ttk.Button(buttons, text="テキストファイルを選択", command=self.select_input_file).pack(side="left")
        ttk.Button(buttons, text="出力先を選択", command=self.select_output_file).pack(side="left", padx=8)
        ttk.Button(buttons, text="Excel議事録を作成", command=self.create_excel).pack(side="right")

        output_frame = ttk.Frame(frame)
        output_frame.pack(fill="x")
        ttk.Label(output_frame, text="出力先:").pack(side="left")
        ttk.Entry(output_frame, textvariable=self.output_path).pack(side="left", fill="x", expand=True, padx=(8, 0))

        ttk.Label(frame, textvariable=self.status, foreground="#2454A6").pack(anchor="w", pady=(10, 0))
        if not DND_AVAILABLE:
            ttk.Label(
                frame,
                text="ドラッグ＆ドロップを有効にするには `pip install tkinterdnd2` を実行してください。",
                foreground="#A65E00",
            ).pack(anchor="w")

    def _setup_dnd(self) -> None:
        if not DND_AVAILABLE:
            return
        self.text.drop_target_register(tkinterdnd2.DND_FILES, tkinterdnd2.DND_TEXT)
        self.text.dnd_bind("<<Drop>>", self.handle_drop)

    def handle_drop(self, event) -> None:
        data = event.data.strip()
        candidate = Path(data.strip("{}"))
        if candidate.exists() and candidate.is_file():
            self.load_file(candidate)
        else:
            self.text.delete("1.0", "end")
            self.text.insert("1.0", data)
            self.status.set("ドロップされたテキストを読み込みました。")

    def select_input_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            self.load_file(Path(path))

    def load_file(self, path: Path) -> None:
        text = path.read_text(encoding="utf-8")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", text)
        self.status.set(f"読み込みました: {path}")

    def select_output_file(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="minutes.xlsx",
        )
        if path:
            self.output_path.set(path)

    def create_excel(self) -> None:
        transcript = self.text.get("1.0", "end").strip()
        if not transcript:
            messagebox.showwarning("入力がありません", "トランスクリプト文章を入力してください。")
            return
        output = Path(self.output_path.get()).expanduser()
        save_minutes_excel(transcript, output)
        self.status.set(f"作成しました: {output}")
        messagebox.showinfo("完了", f"Excel議事録を作成しました。\n{output}")

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    MinutesApp().run()
