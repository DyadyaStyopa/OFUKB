#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Графическая оболочка для OFUKB_CBR_PQ_alt_parser.py.

GUI запускает backend как импортируемый Python-модуль в отдельном процессе.
Это удобнее для упаковки в .app/.exe: пользователю не нужен терминал, IDE или
отдельный выбор Python-интерпретатора.
"""

from __future__ import annotations

import contextlib
import importlib.util
import json
import multiprocessing as mp
import os
import queue
import re
import shlex
import subprocess
import sys
import time
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

APP_TITLE = "Заполнение Excel из данных ЦБ"
BACKEND_FILENAME = "OFUKB_CBR_PQ_alt_parser.py"
SQLITE_EXPORTER_FILENAME = "cbr_sqlite_export.py"
SETTINGS_FILE = Path.home() / ".pq_excel_gui_settings.json"
PROCESS_DONE = "__PROCESS_DONE__"


# -----------------------------------------------------------------------------
# Вспомогательные функции
# -----------------------------------------------------------------------------

def app_dir() -> Path:
    """Папка, где лежит GUI-скрипт или распакованные ресурсы приложения."""
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        return Path(bundle_dir).resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def find_backend_script() -> Optional[Path]:
    """Пытается найти backend рядом с GUI."""
    return find_python_script(BACKEND_FILENAME, "OFUKB_CBR_PQ_alt_parser*.py")


def find_sqlite_exporter_script() -> Optional[Path]:
    """Пытается найти SQLite-экспортёр рядом с GUI."""
    return find_python_script(SQLITE_EXPORTER_FILENAME, "cbr_sqlite_export*.py")


def find_python_script(filename: str, pattern: str) -> Optional[Path]:
    """Пытается найти Python-скрипт рядом с GUI."""
    folder = app_dir()
    exact = folder / filename
    if exact.exists():
        return exact

    candidates = sorted(
        folder.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def resolve_backend_script(value: str) -> Optional[Path]:
    """Возвращает реальный путь backend, не требуя показывать absolute path в GUI."""
    raw = value.strip()
    if not raw or raw == BACKEND_FILENAME:
        return find_backend_script()
    candidate = Path(raw).expanduser()
    if candidate.exists():
        return candidate.resolve()
    local = app_dir() / raw
    if local.exists():
        return local.resolve()
    return None


def resolve_sqlite_exporter_script() -> Optional[Path]:
    return find_sqlite_exporter_script()


def quote_command(args: List[str]) -> str:
    """Красиво экранирует команду для отображения/копирования."""
    return " ".join(shlex.quote(str(a)) for a in args)


def default_output_path(xlsx_path: str, regnum: str) -> str:
    """Строит путь выходного файла как в backend-скрипте."""
    if not xlsx_path:
        return ""
    p = Path(xlsx_path).expanduser()
    suffix = f"_regnum_{regnum}_python_filled.xlsx" if regnum else "_python_filled.xlsx"
    return str(p.with_name(p.stem + suffix))


def default_sqlite_output_path(xlsx_path: str) -> str:
    """Строит путь SQLite-базы рядом с исходной книгой."""
    if not xlsx_path:
        return ""
    p = Path(xlsx_path).expanduser()
    return str(p.with_name(p.stem + "_all_active_banks.sqlite"))


def open_path(path: Path) -> None:
    """Открывает файл/папку системным способом."""
    path = Path(path).expanduser().resolve()
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    elif os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(path)])


def load_settings() -> Dict[str, str]:
    try:
        if SETTINGS_FILE.exists():
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {}


def save_settings(data: Dict[str, str]) -> None:
    try:
        SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        # Настройки необязательны, не мешаем основной работе.
        pass


class QueueTextWriter:
    """Минимальный file-like writer, отправляющий stdout/stderr в multiprocessing.Queue."""

    def __init__(self, out_queue: mp.Queue):
        self.out_queue = out_queue

    def write(self, text: str) -> int:
        if text:
            self.out_queue.put(text)
        return len(text)

    def flush(self) -> None:
        pass


def run_backend_worker(backend_path: str, backend_args: List[str], cwd: str, out_queue: mp.Queue) -> None:
    """Запускается в дочернем процессе и вызывает backend.main(argv)."""
    writer = QueueTextWriter(out_queue)
    code = 1
    try:
        os.chdir(cwd)
        backend = Path(backend_path).expanduser().resolve()
        sys.path.insert(0, str(backend.parent))
        module_name = "ofukb_gui_worker_" + re.sub(r"\W+", "_", backend.stem)
        spec = importlib.util.spec_from_file_location(module_name, backend)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"Не удалось загрузить backend-модуль: {backend}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
        if not hasattr(module, "main"):
            raise RuntimeError(f"В backend-модуле нет функции main(argv): {backend}")

        with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
            result = module.main(backend_args)
        code = int(result or 0)
    except SystemExit as exc:
        code = int(exc.code or 0) if isinstance(exc.code, int) else 1
    except BaseException:
        out_queue.put("\n[Ошибка GUI при запуске backend]\n")
        out_queue.put(traceback.format_exc())
        code = 1
    finally:
        out_queue.put(f"\n[Процесс завершён с кодом {code}]\n")
        out_queue.put({"type": PROCESS_DONE, "code": code})


# -----------------------------------------------------------------------------
# Главное окно
# -----------------------------------------------------------------------------

class PowerQueryExcelGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1060x760")
        self.minsize(900, 640)

        self.proc: Optional[mp.Process] = None
        self.output_queue: Optional[mp.Queue] = None
        self.start_time: Optional[float] = None
        self.last_return_code: Optional[int] = None

        self.settings = load_settings()

        # Не подставляем сохранённые абсолютные пути при старте: окно не должно
        # показывать файловую систему пользователя до явного выбора файлов.
        self.var_backend = tk.StringVar(value=BACKEND_FILENAME)
        self.var_xlsx = tk.StringVar(value="")
        self.var_regnum = tk.StringVar(value=self.settings.get("regnum", ""))
        self.var_output = tk.StringVar(value="")
        self.var_cache_dir = tk.StringVar(value="")
        self.var_log_file = tk.StringVar(value="")
        self.var_debug_dir = tk.StringVar(value="")

        self._last_regnum = self.var_regnum.get()

        self.var_verbose = tk.BooleanVar(value=self.settings.get("verbose", "1") == "1")
        self.var_debug = tk.BooleanVar(value=self.settings.get("debug", "0") == "1")
        self.var_no_cache = tk.BooleanVar(value=self.settings.get("no_cache", "0") == "1")
        self.var_dump_m = tk.BooleanVar(value=self.settings.get("dump_m", "0") == "1")
        self.var_list_only = tk.BooleanVar(value=False)
        self.var_sqlite_all_banks = tk.BooleanVar(value=self.settings.get("sqlite_all_banks", "0") == "1")

        self._build_ui()
        self._bind_events()
        self.after(100, self._poll_process_output)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self, padding=12)
        top.grid(row=0, column=0, sticky="nsew")
        top.columnconfigure(1, weight=1)

        self._add_path_row(top, 0, "Excel-файл:", self.var_xlsx, self.choose_xlsx, "Выбрать .xlsx")
        self._add_path_row(top, 1, "Результат:", self.var_output, self.choose_output, "Куда сохранить")

        reg_frame = ttk.Frame(top)
        reg_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        reg_frame.columnconfigure(1, weight=1)
        ttk.Label(reg_frame, text="regnum банка:").grid(row=0, column=0, sticky="w")
        ttk.Entry(reg_frame, textvariable=self.var_regnum, width=16).grid(row=0, column=1, sticky="w", padx=(8, 6))
        ttk.Button(reg_frame, text="Авто-путь выхода", command=self.set_auto_output).grid(row=0, column=2, padx=(0, 6))
        ttk.Label(
            reg_frame,
            text="Например: 1000, 1481, 2673. Пусто = оставить regnum из M-кода.",
            foreground="#555",
        ).grid(row=0, column=3, sticky="w")

        options = ttk.LabelFrame(top, text="Режимы", padding=10)
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        for i in range(5):
            options.columnconfigure(i, weight=1)
        ttk.Checkbutton(options, text="Подробный лог (--verbose)", variable=self.var_verbose).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(options, text="Debug-режим (--debug)", variable=self.var_debug).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(options, text="Не использовать кэш (--no-cache)", variable=self.var_no_cache).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(options, text="Сохранить M-код (--dump-m)", variable=self.var_dump_m).grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(options, text="Только список таблиц (--list)", variable=self.var_list_only).grid(row=0, column=4, sticky="w")
        ttk.Checkbutton(options, text="SQLite по всем действующим банкам", variable=self.var_sqlite_all_banks).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        advanced = ttk.LabelFrame(top, text="Дополнительные пути", padding=10)
        advanced.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        advanced.columnconfigure(1, weight=1)
        self._add_path_row(advanced, 0, "Backend-скрипт:", self.var_backend, self.choose_backend, "Выбрать скрипт")
        self._add_path_row(advanced, 1, "Папка кэша:", self.var_cache_dir, self.choose_cache_dir, "Выбрать папку")
        self._add_path_row(advanced, 2, "Файл лога:", self.var_log_file, self.choose_log_file, "Выбрать файл")
        self._add_path_row(advanced, 3, "Debug-папка:", self.var_debug_dir, self.choose_debug_dir, "Выбрать папку")

        command_frame = ttk.LabelFrame(self, text="Эквивалентная команда", padding=10)
        command_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        command_frame.columnconfigure(0, weight=1)
        self.command_text = tk.Text(command_frame, height=2, wrap="word", font=("Menlo", 11))
        self.command_text.grid(row=0, column=0, sticky="ew")
        ttk.Button(command_frame, text="Копировать", command=self.copy_command).grid(row=0, column=1, padx=(8, 0), sticky="ns")

        log_frame = ttk.LabelFrame(self, text="Лог выполнения", padding=10)
        log_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 8))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = scrolledtext.ScrolledText(log_frame, wrap="word", font=("Menlo", 11))
        self.log.grid(row=0, column=0, sticky="nsew")

        bottom = ttk.Frame(self, padding=(12, 0, 12, 12))
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.columnconfigure(4, weight=1)

        self.btn_run = ttk.Button(bottom, text="Запустить", command=self.run_script)
        self.btn_run.grid(row=0, column=0, padx=(0, 8))
        self.btn_stop = ttk.Button(bottom, text="Остановить", command=self.stop_script, state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=(0, 8))
        ttk.Button(bottom, text="Очистить лог", command=self.clear_log).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(bottom, text="Сохранить лог из окна", command=self.save_visible_log).grid(row=0, column=3, padx=(0, 8))

        self.status = tk.StringVar(value="Готово к запуску")
        ttk.Label(bottom, textvariable=self.status).grid(row=0, column=4, sticky="w", padx=(8, 8))
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=160)
        self.progress.grid(row=0, column=5, padx=(0, 8))
        ttk.Button(bottom, text="Открыть результат", command=self.open_output).grid(row=0, column=6, padx=(0, 8))
        ttk.Button(bottom, text="Открыть папку", command=self.open_output_folder).grid(row=0, column=7)

        self.update_command_preview()

    def _add_path_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, command, button_text: str) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=3)
        ttk.Button(parent, text=button_text, command=command).grid(row=row, column=2, sticky="e", pady=3)

    def _bind_events(self) -> None:
        variables = [
            self.var_backend,
            self.var_xlsx,
            self.var_output,
            self.var_cache_dir,
            self.var_log_file,
            self.var_debug_dir,
        ]
        for var in variables:
            var.trace_add("write", lambda *_: self.update_command_preview())
        self.var_regnum.trace_add("write", lambda *_: self._on_regnum_changed())
        for var in [self.var_verbose, self.var_debug, self.var_no_cache, self.var_dump_m, self.var_list_only]:
            var.trace_add("write", lambda *_: self.update_command_preview())
        self.var_sqlite_all_banks.trace_add("write", lambda *_: self._on_mode_changed())

    def _on_regnum_changed(self) -> None:
        old_regnum = getattr(self, "_last_regnum", "")
        new_regnum = self.var_regnum.get().strip()
        xlsx = self.var_xlsx.get().strip()
        output = self.var_output.get().strip()
        old_auto_output = default_output_path(xlsx, old_regnum) if xlsx else ""
        if xlsx and (not output or output == old_auto_output):
            self.var_output.set(self.default_output_for_current_mode(xlsx, new_regnum))
        self._last_regnum = new_regnum
        self.update_command_preview()

    def _on_mode_changed(self) -> None:
        xlsx = self.var_xlsx.get().strip()
        output = self.var_output.get().strip()
        old_regnum = getattr(self, "_last_regnum", self.var_regnum.get().strip())
        excel_auto = default_output_path(xlsx, old_regnum) if xlsx else ""
        sqlite_auto = default_sqlite_output_path(xlsx) if xlsx else ""
        if xlsx and (not output or output in {excel_auto, sqlite_auto}):
            self.var_output.set(self.default_output_for_current_mode(xlsx, self.var_regnum.get().strip()))
        self.update_command_preview()

    # ------------------------------------------------------------------
    # Выбор файлов/папок
    # ------------------------------------------------------------------

    def choose_xlsx(self) -> None:
        old_xlsx = self.var_xlsx.get().strip()
        old_regnum = self.var_regnum.get().strip()
        old_output = self.var_output.get().strip()
        old_auto_output = default_output_path(old_xlsx, old_regnum) if old_xlsx else ""
        output_was_auto = not old_output or old_output == old_auto_output

        initialdir = str(Path(old_xlsx).expanduser().parent) if old_xlsx else str(Path.home())
        path = filedialog.askopenfilename(
            title="Выберите исходный Excel-файл",
            initialdir=initialdir,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_xlsx.set(path)
            if output_was_auto:
                self.set_auto_output()

    def choose_output(self) -> None:
        initialdir = str(Path(self.var_xlsx.get()).expanduser().parent) if self.var_xlsx.get() else str(Path.home())
        path = filedialog.asksaveasfilename(
            title="Куда сохранить результат",
            initialdir=initialdir,
            defaultextension=".sqlite" if self.var_sqlite_all_banks.get() else ".xlsx",
            filetypes=[("SQLite files", "*.sqlite *.db"), ("All files", "*.*")]
            if self.var_sqlite_all_banks.get()
            else [("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_output.set(path)

    def choose_backend(self) -> None:
        initialdir = str(app_dir())
        path = filedialog.askopenfilename(
            title=f"Выберите backend-скрипт {BACKEND_FILENAME}",
            initialdir=initialdir,
            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
        )
        if path:
            self.var_backend.set(path)

    def choose_cache_dir(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку HTML-кэша")
        if path:
            self.var_cache_dir.set(path)

    def choose_log_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Выберите файл лога",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            self.var_log_file.set(path)

    def choose_debug_dir(self) -> None:
        path = filedialog.askdirectory(title="Выберите debug-папку")
        if path:
            self.var_debug_dir.set(path)

    # ------------------------------------------------------------------
    # Команда и валидация
    # ------------------------------------------------------------------

    def build_backend_args(self) -> List[str]:
        xlsx = self.var_xlsx.get().strip()
        if self.var_sqlite_all_banks.get():
            args = [xlsx, "--all-banks", "--workers", "4"]
            output = self.var_output.get().strip()
            if output:
                args += ["--output", output]
            args.append("--replace")
            if self.var_no_cache.get():
                args.append("--no-cache")
            if self.var_verbose.get():
                args.append("--verbose")
            cache_dir = self.var_cache_dir.get().strip()
            if cache_dir:
                args += ["--cache-dir", cache_dir]
            return args

        args = [xlsx]

        output = self.var_output.get().strip()
        if output and not self.var_list_only.get():
            args += ["--output", output]

        regnum = self.var_regnum.get().strip()
        if regnum:
            args += ["--regnum", regnum]

        if self.var_list_only.get():
            args.append("--list")
        if self.var_dump_m.get():
            args.append("--dump-m")
        if self.var_no_cache.get():
            args.append("--no-cache")
        if self.var_verbose.get():
            args.append("--verbose")
        if self.var_debug.get():
            args.append("--debug")

        cache_dir = self.var_cache_dir.get().strip()
        if cache_dir:
            args += ["--cache-dir", cache_dir]
        log_file = self.var_log_file.get().strip()
        if log_file:
            args += ["--log-file", log_file]
        debug_dir = self.var_debug_dir.get().strip()
        if debug_dir:
            args += ["--debug-dir", debug_dir]

        return args

    def build_display_command(self) -> List[str]:
        if self.var_sqlite_all_banks.get():
            script = str(resolve_sqlite_exporter_script() or SQLITE_EXPORTER_FILENAME)
        else:
            script = self.var_backend.get().strip() or BACKEND_FILENAME
        return ["python3", script] + self.build_backend_args()

    def validate_inputs(self) -> bool:
        backend = resolve_backend_script(self.var_backend.get())
        xlsx = Path(self.var_xlsx.get().strip()).expanduser()
        regnum = self.var_regnum.get().strip()

        if self.var_sqlite_all_banks.get():
            exporter = resolve_sqlite_exporter_script()
            if exporter is None:
                messagebox.showerror("Ошибка", f"Не найден SQLite-экспортёр:\n{SQLITE_EXPORTER_FILENAME}\n\nПоложите GUI рядом с {SQLITE_EXPORTER_FILENAME}.")
                return False
        elif backend is None:
            messagebox.showerror("Ошибка", f"Не найден backend-скрипт:\n{self.var_backend.get().strip() or BACKEND_FILENAME}\n\nПоложите GUI рядом с {BACKEND_FILENAME} или выберите скрипт вручную.")
            return False
        if not xlsx.exists():
            messagebox.showerror("Ошибка", f"Не найден Excel-файл:\n{xlsx}")
            return False
        if xlsx.suffix.lower() != ".xlsx":
            if not messagebox.askyesno("Предупреждение", "Файл не имеет расширения .xlsx. Всё равно продолжить?"):
                return False
        if not self.var_sqlite_all_banks.get() and regnum and not re.fullmatch(r"\d+", regnum):
            messagebox.showerror("Ошибка", "regnum должен состоять только из цифр, например 1000 или 1481.")
            return False
        return True

    def update_command_preview(self) -> None:
        command = quote_command(self.build_display_command())
        self.command_text.configure(state="normal")
        self.command_text.delete("1.0", "end")
        self.command_text.insert("1.0", command)
        self.command_text.configure(state="disabled")

    def set_auto_output(self) -> None:
        self.var_output.set(self.default_output_for_current_mode(self.var_xlsx.get().strip(), self.var_regnum.get().strip()))

    def default_output_for_current_mode(self, xlsx: str, regnum: str) -> str:
        if self.var_sqlite_all_banks.get():
            return default_sqlite_output_path(xlsx)
        return default_output_path(xlsx, regnum)

    def copy_command(self) -> None:
        command = quote_command(self.build_display_command())
        self.clipboard_clear()
        self.clipboard_append(command)
        self.status.set("Команда скопирована")

    # ------------------------------------------------------------------
    # Запуск процесса
    # ------------------------------------------------------------------

    def run_script(self) -> None:
        if self.proc is not None:
            messagebox.showinfo("Уже выполняется", "Скрипт уже запущен.")
            return
        if not self.validate_inputs():
            return

        self.save_current_settings()
        self.clear_log()
        backend_args = self.build_backend_args()
        self.append_log("Эквивалентная команда:\n" + quote_command(self.build_display_command()) + "\n\n")

        backend = resolve_sqlite_exporter_script() if self.var_sqlite_all_banks.get() else resolve_backend_script(self.var_backend.get())
        if backend is None:
            missing = SQLITE_EXPORTER_FILENAME if self.var_sqlite_all_banks.get() else (self.var_backend.get().strip() or BACKEND_FILENAME)
            messagebox.showerror("Ошибка", f"Не найден скрипт:\n{missing}")
            return
        backend_path = str(backend)
        cwd = str(Path(self.var_xlsx.get()).expanduser().parent)
        self.output_queue = mp.Queue()
        self.last_return_code = None

        try:
            self.proc = mp.Process(
                target=run_backend_worker,
                args=(backend_path, backend_args, cwd, self.output_queue),
                daemon=True,
            )
            self.proc.start()
        except Exception as exc:
            self.proc = None
            self.output_queue = None
            messagebox.showerror("Не удалось запустить", str(exc))
            return

        self.start_time = time.time()
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.progress.start(12)
        self.status.set("Выполняется...")

    def _poll_process_output(self) -> None:
        if self.output_queue is not None:
            try:
                while True:
                    item = self.output_queue.get_nowait()
                    if isinstance(item, dict) and item.get("type") == PROCESS_DONE:
                        self.last_return_code = int(item.get("code", 1))
                        self._on_process_done()
                    else:
                        self.append_log(str(item))
            except queue.Empty:
                pass

        if self.proc is not None and not self.proc.is_alive() and self.last_return_code is None:
            self.last_return_code = self.proc.exitcode
            self.append_log(f"\n[Процесс завершён с кодом {self.last_return_code}]\n")
            self._on_process_done()

        self.after(100, self._poll_process_output)

    def _on_process_done(self) -> None:
        if self.proc is None:
            return
        code = self.last_return_code if self.last_return_code is not None else self.proc.exitcode
        elapsed = int(time.time() - self.start_time) if self.start_time else 0
        self.proc.join(timeout=0.2)
        self.proc = None
        self.output_queue = None
        self.start_time = None
        self.btn_run.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.progress.stop()

        if code == 0:
            self.status.set(f"Готово. Время: {elapsed} сек.")
            messagebox.showinfo("Готово", "Скрипт завершился успешно.")
        else:
            self.status.set(f"Ошибка. Код завершения: {code}")
            messagebox.showerror("Ошибка", "Скрипт завершился с ошибкой. Подробности смотри в логе.")
        self.last_return_code = None

    def stop_script(self) -> None:
        if self.proc is None:
            return
        if not messagebox.askyesno(
            "Остановить",
            "Остановить выполнение скрипта?\n\nПри Excel-режиме временный .tmp.xlsx может остаться рядом с результатом. При SQLite-выгрузке уже записанные строки останутся в базе.",
        ):
            return
        try:
            self.proc.terminate()
            self.append_log("\n[Отправлена команда остановки процесса]\n")
        except Exception as exc:
            messagebox.showerror("Ошибка остановки", str(exc))

    # ------------------------------------------------------------------
    # Лог и открытие результата
    # ------------------------------------------------------------------

    def append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="normal")

    def clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def save_visible_log(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Сохранить лог из окна",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        Path(path).write_text(self.log.get("1.0", "end"), encoding="utf-8")
        self.status.set(f"Лог сохранён: {path}")

    def open_output(self) -> None:
        path = self.var_output.get().strip() or self.default_output_for_current_mode(self.var_xlsx.get().strip(), self.var_regnum.get().strip())
        if not path:
            messagebox.showwarning("Нет пути", "Сначала укажите выходной файл.")
            return
        p = Path(path).expanduser()
        if not p.exists():
            messagebox.showwarning("Файл не найден", f"Файл пока не существует:\n{p}")
            return
        open_path(p)

    def open_output_folder(self) -> None:
        path = self.var_output.get().strip() or self.var_xlsx.get().strip()
        if not path:
            messagebox.showwarning("Нет пути", "Сначала выберите Excel-файл или выходной файл.")
            return
        p = Path(path).expanduser()
        folder = p if p.is_dir() else p.parent
        if not folder.exists():
            messagebox.showwarning("Папка не найдена", f"Папка не существует:\n{folder}")
            return
        open_path(folder)

    def save_current_settings(self) -> None:
        data = {
            # Сохраняем только непутьевые настройки, чтобы GUI не показывал
            # личные абсолютные пути пользователя при следующем запуске.
            "regnum": self.var_regnum.get(),
            "verbose": "1" if self.var_verbose.get() else "0",
            "debug": "1" if self.var_debug.get() else "0",
            "no_cache": "1" if self.var_no_cache.get() else "0",
            "dump_m": "1" if self.var_dump_m.get() else "0",
            "sqlite_all_banks": "1" if self.var_sqlite_all_banks.get() else "0",
        }
        save_settings(data)

    def destroy(self) -> None:
        if self.proc is not None:
            if not messagebox.askyesno("Выход", "Скрипт ещё выполняется. Остановить его и закрыть окно?"):
                return
            try:
                self.proc.terminate()
            except Exception:
                pass
        self.save_current_settings()
        super().destroy()


# -----------------------------------------------------------------------------
# Точка входа
# -----------------------------------------------------------------------------

def main() -> int:
    mp.freeze_support()
    root = PowerQueryExcelGUI()
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
