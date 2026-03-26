"""
gui.py — Interfaccia grafica per getFV.py
Avvia tramite Automator o doppio click su Aggiorna Valori.command
"""

import queue
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox
from typing import Optional

# — importa la logica di getFV dalla stessa cartella —
sys.path.insert(0, str(Path(__file__).parent))
import getFV as core
from openpyxl import load_workbook
import random

# ─────────────────────────────────────────────
# Palette
# ─────────────────────────────────────────────
BG        = "#0f1117"
SURFACE   = "#1a1d27"
BORDER    = "#2a2d3a"
GREEN     = "#00c896"
GREEN_DIM = "#00856a"
RED       = "#ff5c5c"
TEXT      = "#e8eaf0"
MUTED     = "#6b7280"
TICKER_FG = "#a5b4fc"

FONT_TITLE  = ("Helvetica Neue", 15, "bold")
FONT_COUNT  = ("Helvetica Neue", 32, "bold")
FONT_TICKER = ("Helvetica Neue", 12)
FONT_SMALL  = ("Helvetica Neue", 10)
FONT_MONO   = ("Menlo", 10)
FONT_BTN    = ("Helvetica Neue", 12, "bold")

# ─────────────────────────────────────────────
# Conteggio ticker prima del run
# ─────────────────────────────────────────────

def count_total_tickers() -> tuple[int, str]:
    """
    Conta i ticker nel file Excel.
    Restituisce (conteggio, messaggio_errore).
    """
    if not core.EXCEL_FILE_PATH.exists():
        return 0, f"File non trovato:\n{core.EXCEL_FILE_PATH}"
    try:
        wb = load_workbook(core.EXCEL_FILE_PATH, read_only=True)
        total = 0
        for sheet_name in wb.sheetnames:
            if sheet_name in core.SHEETS_TO_SKIP:
                continue
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=5, min_col=3, max_col=3, values_only=True):
                if row[0] is not None and str(row[0]).strip():
                    total += 1
        wb.close()
        return total, ""
    except Exception as e:
        return 0, str(e)

# ─────────────────────────────────────────────
# Worker — process_sheets con callback GUI
# ─────────────────────────────────────────────

def run_worker(
    q: queue.Queue,
    stop_event: threading.Event,
    dry_run: bool = False,
) -> None:
    """
    Esegue l'elaborazione in un thread secondario.
    Comunica con la GUI esclusivamente tramite la queue q.
    Non chiama mai sys.exit() — usa q.put(("error", msg)) per gli errori.
    """
    def put(*args):
        q.put(args)

    # — validazione cookie manuale (non sys.exit) —
    if not core.COOKIE or not core.COOKIE.strip():
        put("error", "Cookie non trovato.\nImposta COOKIE nel file .env")
        return

    if not core.EXCEL_FILE_PATH.exists():
        put("error", f"File Excel non trovato:\n{core.EXCEL_FILE_PATH}")
        return

    try:
        session = core.create_session()
        wb = load_workbook(core.EXCEL_FILE_PATH)

        for sheet_name in wb.sheetnames:
            if stop_event.is_set():
                break
            if sheet_name in core.SHEETS_TO_SKIP:
                continue

            ws = wb[sheet_name]
            offset = core.SHEET_COLUMN_OFFSET.get(sheet_name, 0)

            tickers: list[tuple[int, str]] = []
            for row_idx, row in enumerate(
                ws.iter_rows(min_row=5, min_col=3, max_col=3, values_only=True), start=5
            ):
                val = row[0]
                if val is not None and str(val).strip():
                    tickers.append((row_idx, str(val).strip()))

            if not tickers:
                continue

            put("status", f"Sheet: {sheet_name}  —  {len(tickers)} ticker")

            for batch_start in range(0, len(tickers), core.BATCH_SIZE):
                if stop_event.is_set():
                    break
                batch = tickers[batch_start: batch_start + core.BATCH_SIZE]

                for row_number, ticker in batch:
                    if stop_event.is_set():
                        break

                    put("ticker_start", ticker)
                    result = core.extract_all(ticker, session)

                    if result is not None:
                        core.write_result(ws, row_number, result, offset, dry_run)
                        put("ticker_done", ticker, True)
                    else:
                        put("ticker_done", ticker, False)

                    time.sleep(random.uniform(core.TICKER_DELAY_MIN, core.TICKER_DELAY_MAX))

                if not dry_run and not stop_event.is_set():
                    wb.save(core.EXCEL_FILE_PATH)

                if batch_start + core.BATCH_SIZE < len(tickers) and not stop_event.is_set():
                    wait = random.uniform(core.BATCH_WAIT_MIN, core.BATCH_WAIT_MAX)
                    put("status", f"Pausa batch — {wait:.0f}s...")
                    time.sleep(wait)

            if not dry_run and not stop_event.is_set():
                wb.save(core.EXCEL_FILE_PATH)

        put("done", None)

    except Exception as e:
        put("error", str(e))

# ─────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("LPZ Investing")
        self.configure(bg=BG)
        self.resizable(False, False)

        self._q: queue.Queue      = queue.Queue()
        self._stop                = threading.Event()
        self._thread: Optional[threading.Thread] = None

        self._total      = 0
        self._done       = 0
        self._errors     = 0
        self._t_start: Optional[float] = None

        self._build()
        self._load_total()
        self._poll()

    # ── build ────────────────────────────────

    def _build(self):
        W = 520
        self.geometry(f"{W}x640")

        root = tk.Frame(self, bg=BG)
        root.pack(fill="both", expand=True, padx=28, pady=24)

        # header
        hdr = tk.Frame(root, bg=BG)
        hdr.pack(fill="x", pady=(0, 20))
        tk.Label(hdr, text="LPZ Investing", font=FONT_TITLE,
                 bg=BG, fg=TEXT).pack(side="left")
        self._badge = tk.Label(hdr, text="pronto", font=FONT_SMALL,
                               bg=SURFACE, fg=MUTED, padx=10, pady=3)
        self._badge.pack(side="right")

        # card contatore
        card = tk.Frame(root, bg=SURFACE, highlightbackground=BORDER,
                        highlightthickness=1)
        card.pack(fill="x", pady=(0, 16))
        inner = tk.Frame(card, bg=SURFACE)
        inner.pack(padx=20, pady=18)

        self._lbl_count = tk.Label(inner, text="— / —",
                                   font=FONT_COUNT, bg=SURFACE, fg=TEXT)
        self._lbl_count.pack()

        self._lbl_ticker = tk.Label(inner, text="In attesa di avvio…",
                                    font=FONT_TICKER, bg=SURFACE, fg=TICKER_FG)
        self._lbl_ticker.pack(pady=(4, 0))

        # barra
        bar_frame = tk.Frame(root, bg=BG)
        bar_frame.pack(fill="x", pady=(0, 6))

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("lpz.Horizontal.TProgressbar",
                         troughcolor=SURFACE, background=GREEN,
                         bordercolor=BG, lightcolor=GREEN, darkcolor=GREEN,
                         thickness=6)
        self._bar = ttk.Progressbar(bar_frame, orient="horizontal",
                                     length=W - 56, mode="determinate",
                                     style="lpz.Horizontal.TProgressbar")
        self._bar.pack()

        # tempo
        time_row = tk.Frame(root, bg=BG)
        time_row.pack(fill="x", pady=(6, 0))
        self._lbl_elapsed   = tk.Label(time_row, text="Trascorso  —",
                                       font=FONT_SMALL, bg=BG, fg=MUTED)
        self._lbl_remaining = tk.Label(time_row, text="Rimanente  —",
                                       font=FONT_SMALL, bg=BG, fg=MUTED)
        self._lbl_elapsed.pack(side="left")
        self._lbl_remaining.pack(side="right")

        # status
        self._lbl_status = tk.Label(root, text="",
                                    font=FONT_SMALL, bg=BG, fg=MUTED,
                                    anchor="w")
        self._lbl_status.pack(fill="x", pady=(8, 0))

        # log
        log_outer = tk.Frame(root, bg=SURFACE, highlightbackground=BORDER,
                             highlightthickness=1)
        log_outer.pack(fill="x", pady=(12, 0))

        self._log = tk.Text(log_outer, width=58, height=10,
                            font=FONT_MONO, bg=SURFACE, fg="#9ca3af",
                            insertbackground=TEXT, relief="flat",
                            state="disabled", wrap="none",
                            padx=14, pady=10)
        self._log.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(log_outer, orient="vertical",
                          command=self._log.yview, bg=SURFACE,
                          troughcolor=SURFACE, width=8)
        sb.pack(side="right", fill="y")
        self._log.config(yscrollcommand=sb.set)

        self._log.tag_config("ok",  foreground=GREEN)
        self._log.tag_config("err", foreground=RED)
        self._log.tag_config("inf", foreground=MUTED)
        self._log.tag_config("ts",  foreground="#374151")

        # bottoni
        btn_row = tk.Frame(root, bg=BG)
        btn_row.pack(fill="x", pady=(20, 0))

        self._btn_start = tk.Button(
            btn_row, text="Avvia aggiornamento",
            font=FONT_BTN, bg=GREEN, fg="#0a0f0d",
            activebackground=GREEN_DIM, activeforeground="#0a0f0d",
            relief="flat", padx=20, pady=10, cursor="hand2",
            command=self._on_start
        )
        self._btn_start.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self._btn_stop = tk.Button(
            btn_row, text="Interrompi",
            font=FONT_BTN, bg=BORDER, fg=MUTED,
            activebackground="#3a3d4a", activeforeground=TEXT,
            relief="flat", padx=20, pady=10, cursor="hand2",
            state="disabled", command=self._on_stop
        )
        self._btn_stop.pack(side="left", fill="x", expand=True)

    # ── logica ───────────────────────────────

    def _load_total(self):
        self._log_line("Lettura file Excel…", "inf")
        def _work():
            total, err = count_total_tickers()
            self._q.put(("total", total, err))
        threading.Thread(target=_work, daemon=True).start()

    def _on_start(self):
        if self._total == 0:
            messagebox.showwarning(
                "File non trovato",
                f"Nessun ticker trovato.\nControlla il percorso:\n{core.EXCEL_FILE_PATH}"
            )
            return

        self._stop.clear()
        self._done = 0
        self._errors = 0
        self._t_start = time.time()
        self._bar["value"] = 0
        self._bar["maximum"] = self._total
        self._set_badge("in corso", GREEN)
        self._btn_start.config(state="disabled")
        self._btn_stop.config(state="normal", bg="#3a1a1a", fg=RED)
        self._log_line("Avvio elaborazione…", "inf")

        self._thread = threading.Thread(
            target=run_worker,
            args=(self._q, self._stop),
            daemon=True
        )
        self._thread.start()

    def _on_stop(self):
        self._stop.set()
        self._btn_stop.config(state="disabled")
        self._log_line("Interruzione in corso…", "inf")

    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                self._dispatch(msg)
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def _dispatch(self, msg: tuple):
        kind = msg[0]

        if kind == "total":
            _, total, err = msg
            self._total = total
            if err:
                self._log_line(f"Errore lettura: {err}", "err")
                self._lbl_count.config(text="Errore")
            else:
                self._lbl_count.config(text=f"0 / {total}")
                self._log_line(f"{total} ticker trovati.", "inf")

        elif kind == "ticker_start":
            _, ticker = msg
            self._lbl_ticker.config(text=ticker)

        elif kind == "ticker_done":
            _, ticker, ok = msg
            self._done += 1
            if not ok:
                self._errors += 1
            self._bar["value"] = self._done
            self._lbl_count.config(text=f"{self._done} / {self._total}")
            self._log_line(
                f"{'✓' if ok else '✗'}  {ticker}",
                "ok" if ok else "err"
            )
            self._refresh_time()

        elif kind == "status":
            _, txt = msg
            self._lbl_status.config(text=txt)

        elif kind == "done":
            self._finish(interrupted=False)

        elif kind == "error":
            _, err = msg
            self._log_line(f"ERRORE: {err}", "err")
            messagebox.showerror("Errore", err)
            self._finish(interrupted=True)

    def _finish(self, interrupted: bool):
        self._btn_start.config(state="normal")
        self._btn_stop.config(state="disabled", bg=BORDER, fg=MUTED)
        label = "interrotto" if interrupted else "completato"
        self._set_badge(label, MUTED)
        self._lbl_ticker.config(text="Operazione terminata")
        summary = f"{self._done} elaborati  ·  {self._errors} errori"
        self._lbl_status.config(text=summary)
        self._log_line(f"— {summary} —", "inf")

    def _refresh_time(self):
        if not self._t_start or self._done == 0:
            return
        elapsed = time.time() - self._t_start
        avg = elapsed / self._done
        remaining = avg * (self._total - self._done)
        self._lbl_elapsed.config(text=f"Trascorso  {self._fmt(elapsed)}")
        self._lbl_remaining.config(text=f"Rimanente  {self._fmt(remaining)}")

    @staticmethod
    def _fmt(sec: float) -> str:
        h, rem = divmod(int(sec), 3600)
        m, s   = divmod(rem, 60)
        if h:   return f"{h}h {m:02d}m"
        if m:   return f"{m}m {s:02d}s"
        return  f"{s}s"

    def _set_badge(self, text: str, color: str):
        self._badge.config(text=text, fg=color)

    def _log_line(self, text: str, tag: str = ""):
        self._log.config(state="normal")
        ts = time.strftime("%H:%M:%S")
        self._log.insert("end", f"[{ts}]  ", "ts")
        self._log.insert("end", f"{text}\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")


# ─────────────────────────────────────────────
# Entrypoint
# ─────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()