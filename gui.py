"""
gui.py
──────
PPT 自動配音轉影片與資通院交付物生成圖形化介面 (GUI)
提供雙模式視覺化操作：
  【模式 1】單一簡報處理 (Single PPT)：精細調參、單元測試、講稿/備註微調
  【模式 2】批次全自動處理 (Batch Processing)：支援 20+ 份 PPT 一鍵排隊轉換、自動跳過已完成、自動清理暫存
"""

import os
import sys
import glob
import time
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from core.utils import (
    load_settings, save_pptx_settings,
    get_pptx_stem, get_default_clips_dir, get_default_deliverables_dir,
    parse_range_str,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PPT 自動配音轉影片 & 資通院教材交付物工具 (單檔 & 20+ 批次全自動版)")
        self.geometry("980x820")
        self.minsize(880, 680)

        # 設定樣式
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("vista")
        except Exception:
            pass

        self.log_queue = queue.Queue()
        self.current_process = None
        self.is_running = False

        self._build_ui()
        self._auto_detect_pptx()
        self._poll_log_queue()

    def _build_ui(self):
        # 建立上方分頁 (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="x", padx=10, pady=(6, 0))

        # ── 分頁 1：單一簡報處理 ───────────────────────────
        self.tab_single = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_single, text="  🖥️ 單一簡報處理 (Single Mode)  ")
        self._build_single_tab(self.tab_single)

        # ── 分頁 2：批次全自動處理 ─────────────────────────
        self.tab_batch = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_batch, text="  📦 批次全自動處理 (Batch Mode, 20+ PPT)  ")
        self._build_batch_tab(self.tab_batch)

        # ── 底部通用區：即時日誌與狀態列 ───────────────────
        log_frame = ttk.LabelFrame(self, text="  執行日誌  ", padding=6)
        log_frame.pack(fill="both", expand=True, padx=10, pady=6)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap="word", font=("Consolas", 9),
            bg="#1e1e1e", fg="#e0e0e0", insertbackground="white"
        )
        self.log_text.pack(fill="both", expand=True)

        # 狀態列
        self.status_var = tk.StringVar(value="就緒")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=(6, 3))
        status_bar.pack(fill="x", side="bottom")

    # ══════════════════════════════════════════════════════════
    # 分頁 1：單一簡報介面建構
    # ══════════════════════════════════════════════════════════

    def _build_single_tab(self, parent):
        # 1. 檔案選擇區
        file_frame = ttk.LabelFrame(parent, text="  1. 簡報檔案選擇 (PPTX)  ", padding=8)
        file_frame.pack(fill="x", padx=4, pady=4)

        ttk.Label(file_frame, text="PPT 檔案：").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.pptx_var = tk.StringVar()
        self.pptx_entry = ttk.Entry(file_frame, textvariable=self.pptx_var, font=("Segoe UI", 9))
        self.pptx_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
        self.pptx_var.trace_add("write", lambda *args: self._on_pptx_changed())

        browse_btn = ttk.Button(file_frame, text=" 瀏覽 (Browse)... ", command=self._browse_pptx)
        browse_btn.grid(row=0, column=2, padx=4, pady=2)
        file_frame.columnconfigure(1, weight=1)

        self.ppt_info_var = tk.StringVar(value="（尚未選取 PPT）")
        ppt_info_label = ttk.Label(file_frame, textvariable=self.ppt_info_var, foreground="#2563eb", font=("Segoe UI", 8, "italic"))
        ppt_info_label.grid(row=1, column=1, sticky="w", padx=4, pady=0)

        # 2. 參數設定區
        settings_frame = ttk.LabelFrame(parent, text="  2. PPT 專屬參數與範圍設定 (自動物理隔離)  ", padding=8)
        settings_frame.pack(fill="x", padx=4, pady=4)

        # 第 0 列：notes_slides
        ttk.Label(settings_frame, text="唸備註頁碼：", foreground="#b45309", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=4, pady=3)
        self.notes_slides_var = tk.StringVar()
        notes_entry = ttk.Entry(settings_frame, textvariable=self.notes_slides_var, width=20)
        notes_entry.grid(row=0, column=1, sticky="w", padx=4, pady=3)

        ttk.Label(settings_frame, text="（例：1,2,680,681 或 1-5，留空代表不唸）", foreground="gray").grid(row=0, column=2, columnspan=2, sticky="w", padx=2, pady=3)

        self.btn_save_settings = ttk.Button(settings_frame, text="💾 儲存此 PPT 參數", command=self._save_ppt_settings)
        self.btn_save_settings.grid(row=0, column=4, padx=4, pady=3, sticky="e")

        # 第 1 列：指定頁碼與指定標題
        ttk.Label(settings_frame, text="指定處理頁碼：").grid(row=1, column=0, sticky="w", padx=4, pady=3)
        self.pages_var = tk.StringVar()
        pages_entry = ttk.Entry(settings_frame, textvariable=self.pages_var, width=20)
        pages_entry.grid(row=1, column=1, sticky="w", padx=4, pady=3)
        ttk.Label(settings_frame, text="（例如：1-3,5 或留空代表全部）", foreground="gray").grid(row=1, column=2, sticky="w", padx=2, pady=3)

        ttk.Label(settings_frame, text="指定標題編號：").grid(row=1, column=3, sticky="w", padx=8, pady=3)
        self.range_var = tk.StringVar()
        range_entry = ttk.Entry(settings_frame, textvariable=self.range_var, width=14)
        range_entry.grid(row=1, column=4, sticky="w", padx=4, pady=3)

        # 第 2 列：平行數與 GPU
        ttk.Label(settings_frame, text="平行合成數：").grid(row=2, column=0, sticky="w", padx=4, pady=3)
        self.workers_var = tk.StringVar(value="4")
        workers_spin = ttk.Spinbox(settings_frame, from_=1, to=16, textvariable=self.workers_var, width=5)
        workers_spin.grid(row=2, column=1, sticky="w", padx=4, pady=3)

        self.gpu_var = tk.BooleanVar(value=True)
        gpu_check = ttk.Checkbutton(settings_frame, text="啟用 NVIDIA GPU 加速編碼 (nvenc)", variable=self.gpu_var)
        gpu_check.grid(row=2, column=2, columnspan=3, sticky="w", padx=4, pady=3)

        # 第 3 列：輸出影片與交付資料夾
        ttk.Label(settings_frame, text="輸出影片檔名：").grid(row=3, column=0, sticky="w", padx=4, pady=3)
        self.output_video_var = tk.StringVar(value="教學簡報測試.mp4")
        out_video_entry = ttk.Entry(settings_frame, textvariable=self.output_video_var, width=20)
        out_video_entry.grid(row=3, column=1, sticky="w", padx=4, pady=3)

        ttk.Label(settings_frame, text="交付資料夾：").grid(row=3, column=2, sticky="w", padx=6, pady=3)
        self.output_dir_var = tk.StringVar(value="deliverables/教學簡報測試")
        out_dir_entry = ttk.Entry(settings_frame, textvariable=self.output_dir_var, width=30)
        out_dir_entry.grid(row=3, column=3, columnspan=2, sticky="ew", padx=4, pady=3)

        # 第 4 列：clips 資料夾
        ttk.Label(settings_frame, text="clips 資料夾：").grid(row=4, column=0, sticky="w", padx=4, pady=3)
        self.clips_dir_var = tk.StringVar(value="clips/教學簡報測試")
        clips_dir_entry = ttk.Entry(settings_frame, textvariable=self.clips_dir_var, width=20)
        clips_dir_entry.grid(row=4, column=1, sticky="w", padx=4, pady=3)
        ttk.Label(settings_frame, text="（已依 PPT 檔名自動物理隔離，避免換檔覆蓋）", foreground="#059669", font=("Segoe UI", 8)).grid(row=4, column=2, columnspan=3, sticky="w", padx=2, pady=3)

        settings_frame.columnconfigure(3, weight=1)

        # 3. 執行操作區
        btn_frame = ttk.LabelFrame(parent, text="  3. 執行操作  ", padding=8)
        btn_frame.pack(fill="x", padx=4, pady=4)

        self.btn_all = tk.Button(
            btn_frame, text="🚀 一鍵完整執行 (Phase 1 + 2 + 交付物)",
            bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"),
            relief="raised", padx=8, pady=5, command=self._run_all_pipeline
        )
        self.btn_all.grid(row=0, column=0, padx=4, pady=3, sticky="ew")

        self.btn_p1 = ttk.Button(btn_frame, text="🎬 步驟 1：生成 clips (Phase 1)", command=self._run_phase1)
        self.btn_p1.grid(row=0, column=1, padx=3, pady=3, sticky="ew")

        self.btn_p2 = ttk.Button(btn_frame, text="🎞️ 步驟 2：串接與交付物 (Phase 2)", command=self._run_phase2)
        self.btn_p2.grid(row=0, column=2, padx=3, pady=3, sticky="ew")

        self.btn_deliv = tk.Button(
            btn_frame, text="📑 僅產出交付物 (秒級)",
            bg="#059669", fg="white", font=("Segoe UI", 9, "bold"),
            relief="raised", padx=6, pady=3, command=self._run_deliverables_only
        )
        self.btn_deliv.grid(row=0, column=3, padx=3, pady=3, sticky="ew")

        sub_btn_frame = ttk.Frame(btn_frame)
        sub_btn_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(4, 0))

        self.btn_open_deliv = ttk.Button(sub_btn_frame, text="📂 開啟此 PPT 交付資料夾", command=self._open_deliverables_folder)
        self.btn_open_deliv.pack(side="left", padx=3)

        self.btn_open_clips = ttk.Button(sub_btn_frame, text="📁 開啟此 PPT clips 短片資料夾", command=self._open_clips_folder)
        self.btn_open_clips.pack(side="left", padx=3)

        self.btn_stop = tk.Button(
            sub_btn_frame, text="⏹️ 中止執行",
            bg="#dc2626", fg="white", font=("Segoe UI", 9, "bold"),
            state="disabled", command=self._stop_execution
        )
        self.btn_stop.pack(side="right", padx=3)

        self.btn_clear_log = ttk.Button(sub_btn_frame, text="🧹 清空日誌", command=self._clear_log)
        self.btn_clear_log.pack(side="right", padx=3)

        for c in range(4):
            btn_frame.columnconfigure(c, weight=1)

    # ══════════════════════════════════════════════════════════
    # 分頁 2：批次全自動處理介面建構
    # ══════════════════════════════════════════════════════════

    def _build_batch_tab(self, parent):
        # 1. 資料夾選取區
        folder_frame = ttk.LabelFrame(parent, text="  1. PPT 來源資料夾選擇  ", padding=8)
        folder_frame.pack(fill="x", padx=4, pady=4)

        ttk.Label(folder_frame, text="PPT 資料夾：").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.batch_dir_var = tk.StringVar(value=os.path.abspath("."))
        self.batch_dir_entry = ttk.Entry(folder_frame, textvariable=self.batch_dir_var, font=("Segoe UI", 9))
        self.batch_dir_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=2)

        browse_folder_btn = ttk.Button(folder_frame, text=" 瀏覽資料夾... ", command=self._browse_batch_dir)
        browse_folder_btn.grid(row=0, column=2, padx=4, pady=2)

        scan_btn = ttk.Button(folder_frame, text=" 🔄 重新掃描 ", command=self._scan_batch_pptx)
        scan_btn.grid(row=0, column=3, padx=4, pady=2)
        folder_frame.columnconfigure(1, weight=1)

        # 2. PPT 檔案清單表格 (Treeview)
        list_frame = ttk.LabelFrame(parent, text="  2. 掃描到的 PPT 簡報清單  ", padding=6)
        list_frame.pack(fill="both", expand=True, padx=4, pady=4)

        tree_scroll = ttk.Scrollbar(list_frame)
        tree_scroll.pack(side="right", fill="y")

        columns = ("idx", "name", "notes", "status", "output")
        self.batch_tree = ttk.Treeview(
            list_frame, columns=columns, show="headings",
            yscrollcommand=tree_scroll.set, height=6
        )
        tree_scroll.config(command=self.batch_tree.yview)

        self.batch_tree.heading("idx", text="#")
        self.batch_tree.heading("name", text="PPT 檔案名稱")
        self.batch_tree.heading("notes", text="專屬 notes_slides")
        self.batch_tree.heading("status", text="預計狀態")
        self.batch_tree.heading("output", text="輸出目標影片")

        self.batch_tree.column("idx", width=40, anchor="center")
        self.batch_tree.column("name", width=220, anchor="w")
        self.batch_tree.column("notes", width=160, anchor="center")
        self.batch_tree.column("status", width=90, anchor="center")
        self.batch_tree.column("output", width=260, anchor="w")

        self.batch_tree.pack(fill="both", expand=True)

        self.batch_summary_var = tk.StringVar(value="尚未掃描資料夾")
        ttk.Label(list_frame, textvariable=self.batch_summary_var, foreground="#2563eb", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 0))

        # 3. 批次選項與執行
        batch_opt_frame = ttk.LabelFrame(parent, text="  3. 批次轉換選項與控制  ", padding=8)
        batch_opt_frame.pack(fill="x", padx=4, pady=4)

        self.batch_skip_var = tk.BooleanVar(value=True)
        skip_chk = ttk.Checkbutton(batch_opt_frame, text="⏩ 自動跳過已完成之 PPT（若輸出影片已存在）", variable=self.batch_skip_var)
        skip_chk.grid(row=0, column=0, sticky="w", padx=6, pady=3)

        self.batch_clean_var = tk.BooleanVar(value=False)
        clean_chk = ttk.Checkbutton(batch_opt_frame, text="🧹 每份轉換完成後自動清理 clips 暫存（節省數十 GB 空間）", variable=self.batch_clean_var)
        clean_chk.grid(row=0, column=1, columnspan=2, sticky="w", padx=6, pady=3)

        self.batch_gpu_var = tk.BooleanVar(value=True)
        batch_gpu_chk = ttk.Checkbutton(batch_opt_frame, text="啟用 NVIDIA GPU 加速", variable=self.batch_gpu_var)
        batch_gpu_chk.grid(row=1, column=0, sticky="w", padx=6, pady=3)

        ttk.Label(batch_opt_frame, text="平行合成數：").grid(row=1, column=1, sticky="w", padx=6, pady=3)
        self.batch_workers_var = tk.StringVar(value="4")
        batch_workers_spin = ttk.Spinbox(batch_opt_frame, from_=1, to=16, textvariable=self.batch_workers_var, width=5)
        batch_workers_spin.grid(row=1, column=2, sticky="w", padx=2, pady=3)

        # 批次按鈕
        btn_box = ttk.Frame(batch_opt_frame)
        btn_box.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))

        self.btn_batch_start = tk.Button(
            btn_box, text="🚀 開始批次一鍵全自動轉換 (Phase 1 + 2 全部排隊執行)",
            bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"),
            relief="raised", padx=10, pady=6, command=self._run_batch_pipeline
        )
        self.btn_batch_start.pack(side="left", padx=4, fill="x", expand=True)

        self.btn_batch_open_deliv = ttk.Button(btn_box, text="📂 開啟交付物總目錄 (deliverables)", command=self._open_deliverables_folder)
        self.btn_batch_open_deliv.pack(side="left", padx=4)

        self.btn_batch_stop = tk.Button(
            btn_box, text="⏹️ 中止批次執行",
            bg="#dc2626", fg="white", font=("Segoe UI", 9, "bold"),
            state="disabled", command=self._stop_execution
        )
        self.btn_batch_stop.pack(side="right", padx=4)

    # ══════════════════════════════════════════════════════════
    # 單檔 PPT 切換與設定聯動
    # ══════════════════════════════════════════════════════════

    def _on_pptx_changed(self):
        path = self.pptx_var.get().strip()
        if not path:
            self.ppt_info_var.set("（尚未選取 PPT）")
            return

        stem = get_pptx_stem(path)
        self.ppt_info_var.set(f"★ 當前 PPT：{stem}（已套用專屬路徑隔離）")

        self.output_video_var.set(f"{stem}.mp4")
        self.output_dir_var.set(get_default_deliverables_dir(path))
        self.clips_dir_var.set(get_default_clips_dir(path))

        settings = load_settings("settings.yaml", path)
        notes = settings.get("notes_slides", "")

        if isinstance(notes, list):
            notes_str = ",".join(str(n) for n in notes)
        elif notes is not None:
            notes_str = str(notes).strip("[]")
        else:
            notes_str = ""

        self.notes_slides_var.set(notes_str)

    def _save_ppt_settings(self):
        pptx = self.pptx_var.get().strip()
        if not pptx:
            messagebox.showwarning("提示", "請先選擇 PPT 檔案！")
            return

        stem = get_pptx_stem(pptx)
        raw_notes = self.notes_slides_var.get().strip()

        if not raw_notes:
            parsed_notes = []
        elif "-" in raw_notes:
            parsed_notes = raw_notes
        else:
            try:
                parsed_notes = [int(p.strip()) for p in raw_notes.split(",") if p.strip().isdigit()]
            except Exception:
                parsed_notes = raw_notes

        success = save_pptx_settings(pptx, {"notes_slides": parsed_notes}, "settings.yaml")
        if success:
            msg = f"已成功將 PPT【{stem}】專屬設定儲存至 settings.yaml！\nnotes_slides: {parsed_notes}"
            self._log(f"\n[設定更新] {msg}\n")
            messagebox.showinfo("儲存成功", msg)
        else:
            messagebox.showerror("錯誤", "儲存設定失敗，請確認 settings.yaml 權限。")

    def _auto_detect_pptx(self):
        candidates = ["教學簡報測試.pptx", "測試.pptx"]
        for c in candidates:
            if os.path.exists(c):
                self.pptx_var.set(os.path.abspath(c))
                self._scan_batch_pptx()
                return
        all_pptx = glob.glob("*.pptx")
        if all_pptx:
            self.pptx_var.set(os.path.abspath(all_pptx[0]))
        self._scan_batch_pptx()

    def _browse_pptx(self):
        path = filedialog.askopenfilename(
            title="請選擇 PPT 簡報檔案",
            filetypes=[("PowerPoint 檔案", "*.pptx"), ("所有檔案", "*.*")]
        )
        if path:
            self.pptx_var.set(path)

    # ══════════════════════════════════════════════════════════
    # 批次資料夾與掃描邏輯
    # ══════════════════════════════════════════════════════════

    def _browse_batch_dir(self):
        folder = filedialog.askdirectory(title="請選擇存放多個 PPT 簡報的資料夾")
        if folder:
            self.batch_dir_var.set(folder)
            self._scan_batch_pptx()

    def _scan_batch_pptx(self):
        folder = self.batch_dir_var.get().strip()
        if not folder or not os.path.isdir(folder):
            return

        for item in self.batch_tree.get_children():
            self.batch_tree.delete(item)

        files = glob.glob(os.path.join(folder, "*.pptx"))
        valid_files = [f for f in files if not os.path.basename(f).startswith("~$")]
        valid_files.sort(key=lambda x: os.path.basename(x).lower())

        completed_count = 0
        for idx, f in enumerate(valid_files, start=1):
            stem = get_pptx_stem(f)
            fname = os.path.basename(f)
            cfg = load_settings("settings.yaml", f)
            notes = str(cfg.get("notes_slides", []))

            deliv_dir = get_default_deliverables_dir(f)
            target_vid = os.path.join(deliv_dir, f"{stem}.mp4")

            if os.path.exists(target_vid) and os.path.getsize(target_vid) > 1024:
                status = "✅ 已完成"
                completed_count += 1
            else:
                status = "⏳ 待處理"

            rel_target = os.path.relpath(target_vid) if target_vid.startswith(os.getcwd()) else target_vid
            self.batch_tree.insert("", "end", values=(idx, fname, notes, status, rel_target))

        self.batch_summary_var.set(
            f"共掃描到 {len(valid_files)} 份 PPT 簡報（其中 {completed_count} 份已存在交付影片，{len(valid_files) - completed_count} 份待生成）"
        )

    # ══════════════════════════════════════════════════════════
    # 日誌與背景執行
    # ══════════════════════════════════════════════════════════

    def _log(self, text: str):
        self.log_queue.put(text)

    def _poll_log_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.insert(tk.END, msg)
            self.log_text.see(tk.END)
        self.after(100, self._poll_log_queue)

    def _clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def _set_ui_state(self, running: bool):
        self.is_running = running
        state = "disabled" if running else "normal"
        self.btn_all.config(state=state)
        self.btn_p1.config(state=state)
        self.btn_p2.config(state=state)
        self.btn_deliv.config(state=state)
        self.btn_save_settings.config(state=state)
        self.btn_batch_start.config(state=state)
        self.btn_stop.config(state="normal" if running else "disabled")
        self.btn_batch_stop.config(state="normal" if running else "disabled")

    def _open_deliverables_folder(self):
        folder = os.path.abspath(self.output_dir_var.get().strip() or "deliverables")
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def _open_clips_folder(self):
        folder = os.path.abspath(self.clips_dir_var.get().strip() or "clips")
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def _stop_execution(self):
        if self.current_process and self.is_running:
            if messagebox.askyesno("確認中止", "確定要中止目前正在執行的任務嗎？"):
                try:
                    self.current_process.terminate()
                    self._log("\n\n[使用者中止] 任務已被終止。\n")
                    self.status_var.set("已中止")
                except Exception as e:
                    self._log(f"\n[錯誤] 中止處理程序失敗：{e}\n")

    def _run_commands_async(self, cmds: list[list[str]], title: str):
        if self.is_running:
            messagebox.showwarning("執行中", "目前已有任務正在執行，請稍候或點擊中止。")
            return

        def _worker():
            self._set_ui_state(True)
            self.status_var.set(f"正在執行：{title}...")
            self._log(f"\n{'='*75}\n[開始執行] {title}\n{'='*75}\n")

            all_success = True
            for cmd_idx, cmd in enumerate(cmds, start=1):
                cmd_display = " ".join(f'"{c}"' if " " in c else c for c in cmd)
                self._log(f"\n> {cmd_display}\n\n")

                try:
                    self.current_process = subprocess.Popen(
                        cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        bufsize=1,
                    )

                    for line in self.current_process.stdout:
                        self._log(line)

                    self.current_process.wait()
                    code = self.current_process.returncode

                    if code != 0:
                        self._log(f"\n[錯誤] 指令執行失敗，結束代碼：{code}\n")
                        all_success = False
                        break

                except Exception as e:
                    self._log(f"\n[例外錯誤] 執行失敗：{e}\n")
                    all_success = False
                    break

            self.current_process = None
            self._set_ui_state(False)
            self._scan_batch_pptx()

            if all_success:
                self.status_var.set(f"【成功】{title} 已完成！")
                self._log(f"\n{'='*75}\n[完成] {title} 全部順利執行完畢！\n{'='*75}\n")
                messagebox.showinfo("執行完成", f"【{title}】已順利完成！")
            else:
                self.status_var.set(f"【失敗】{title} 執行中斷。")
                messagebox.showerror("執行失敗", f"【{title}】執行過程中發生錯誤，請查看日誌。")

        threading.Thread(target=_worker, daemon=True).start()

    # ══════════════════════════════════════════════════════════
    # 單檔模式按鈕任務
    # ══════════════════════════════════════════════════════════

    def _run_phase1(self):
        pptx = self.pptx_var.get().strip()
        if not pptx or not os.path.exists(pptx):
            messagebox.showerror("錯誤", "請先選擇正確的 PPT 檔案！")
            return

        clips_dir = self.clips_dir_var.get().strip() or get_default_clips_dir(pptx)
        cmd = [
            sys.executable, "-u", "phase1_generate.py",
            "--pptx", pptx,
            "--clips", clips_dir,
        ]
        notes = self.notes_slides_var.get().strip()
        if notes:
            cmd.extend(["--notes-slides", notes])
        pages = self.pages_var.get().strip()
        if pages:
            cmd.extend(["--pages", pages])
        rng = self.range_var.get().strip()
        if rng:
            cmd.extend(["--range", rng])
        workers = self.workers_var.get().strip()
        if workers:
            cmd.extend(["--workers", workers])
        if not self.gpu_var.get():
            cmd.append("--no-gpu")

        self._run_commands_async([cmd], f"單檔 Phase 1：生成 clips ({get_pptx_stem(pptx)})")

    def _run_phase2(self):
        pptx = self.pptx_var.get().strip()
        clips_dir = self.clips_dir_var.get().strip() or get_default_clips_dir(pptx)
        out_vid = self.output_video_var.get().strip() or f"{get_pptx_stem(pptx)}.mp4"
        out_dir = self.output_dir_var.get().strip() or get_default_deliverables_dir(pptx)

        cmd = [
            sys.executable, "-u", "phase2_compose.py",
            "--clips", clips_dir,
            "--output", out_vid,
            "--output-dir", out_dir,
        ]
        if pptx and os.path.exists(pptx):
            cmd.extend(["--pptx", pptx])

        self._run_commands_async([cmd], f"單檔 Phase 2：串接影片與交付物 ({get_pptx_stem(pptx)})")

    def _run_deliverables_only(self):
        pptx = self.pptx_var.get().strip()
        if not pptx or not os.path.exists(pptx):
            messagebox.showerror("錯誤", "請先選擇正確的 PPT 檔案！")
            return

        clips_dir = self.clips_dir_var.get().strip() or get_default_clips_dir(pptx)
        out_dir = self.output_dir_var.get().strip() or get_default_deliverables_dir(pptx)

        cmd = [
            sys.executable, "-u", "generate_deliverables.py",
            "--pptx", pptx,
            "--clips", clips_dir,
            "--output-dir", out_dir,
        ]
        self._run_commands_async([cmd], f"單檔僅產出資通院交付檔案 ({get_pptx_stem(pptx)})")

    def _run_all_pipeline(self):
        pptx = self.pptx_var.get().strip()
        if not pptx or not os.path.exists(pptx):
            messagebox.showerror("錯誤", "請先選擇正確的 PPT 檔案！")
            return

        clips_dir = self.clips_dir_var.get().strip() or get_default_clips_dir(pptx)
        out_vid = self.output_video_var.get().strip() or f"{get_pptx_stem(pptx)}.mp4"
        out_dir = self.output_dir_var.get().strip() or get_default_deliverables_dir(pptx)

        cmd_p1 = [
            sys.executable, "-u", "phase1_generate.py",
            "--pptx", pptx,
            "--clips", clips_dir,
        ]
        notes = self.notes_slides_var.get().strip()
        if notes:
            cmd_p1.extend(["--notes-slides", notes])
        pages = self.pages_var.get().strip()
        if pages:
            cmd_p1.extend(["--pages", pages])
        rng = self.range_var.get().strip()
        if rng:
            cmd_p1.extend(["--range", rng])
        workers = self.workers_var.get().strip()
        if workers:
            cmd_p1.extend(["--workers", workers])
        if not self.gpu_var.get():
            cmd_p1.append("--no-gpu")

        cmd_p2 = [
            sys.executable, "-u", "phase2_compose.py",
            "--pptx", pptx,
            "--clips", clips_dir,
            "--output", out_vid,
            "--output-dir", out_dir,
        ]

        self._run_commands_async([cmd_p1, cmd_p2], f"單檔全自動一鍵轉換 ({get_pptx_stem(pptx)})")

    # ══════════════════════════════════════════════════════════
    # 批次模式按鈕任務
    # ══════════════════════════════════════════════════════════

    def _run_batch_pipeline(self):
        folder = self.batch_dir_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("錯誤", "請先選擇有效的 PPT 資料夾！")
            return

        files = glob.glob(os.path.join(folder, "*.pptx"))
        valid_files = [f for f in files if not os.path.basename(f).startswith("~$")]
        if not valid_files:
            messagebox.showwarning("提示", f"在資料夾【{folder}】中找不到任何 PPTX 檔案！")
            return

        cmd = [
            sys.executable, "-u", "batch_process.py",
            "--dir", folder,
            "--workers", self.batch_workers_var.get().strip() or "4",
        ]

        if self.batch_skip_var.get():
            cmd.append("--skip-existing")
        if self.batch_clean_var.get():
            cmd.append("--clean")
        if not self.batch_gpu_var.get():
            cmd.append("--no-gpu")

        self._run_commands_async([cmd], f"批次全自動轉換（共 {len(valid_files)} 份 PPT）")


if __name__ == "__main__":
    app = App()
    app.mainloop()
