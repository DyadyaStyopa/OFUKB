#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Графическая оболочка для wtf2_xmlsafe_cellonly.py.

Что делает GUI:
- позволяет выбрать исходный Excel-файл;
- позволяет указать regnum банка;
- позволяет выбрать выходной файл, кэш, лог и debug-папку;
- запускает основной скрипт в отдельном процессе;
- показывает весь вывод скрипта в окне;
- позволяет остановить выполнение и открыть готовый файл.

Зависимости GUI: только стандартная библиотека Python.
Основной backend-скрипт по-прежнему требует свои зависимости:
    pip install pandas requests beautifulsoup4 lxml

Рекомендуемое размещение:
    pq_excel_gui.py
    wtf2_xmlsafe_cellonly.py

в одной папке.
"""

from __future__ import annotations

import json
import os
import queue
import re
import shlex
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Dict, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

APP_TITLE = "Заполнение Excel из данных ЦБ"
SETTINGS_FILE = Path.home() / ".pq_excel_gui_settings.json"


# -----------------------------------------------------------------------------
# Вспомогательные функции
# -----------------------------------------------------------------------------

def app_dir() -> Path:
    """Папка, где лежит GUI-скрипт."""
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def find_backend_script() -> Optional[Path]:
    """Пытается найти основной скрипт рядом с GUI."""
    folder = app_dir()
    exact = folder / "wtf2_xmlsafe_cellonly.py"
    if exact.exists():
        return exact

    candidates = sorted(
        folder.glob("wtf2_xmlsafe_cellonly*.py"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for candidate in candidates:
        if candidate.resolve() != Path(__file__).resolve():
            return candidate
    return None


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
        # Настройки — необязательная удобная функция, поэтому не мешаем основной работе.
        pass


# -----------------------------------------------------------------------------
# Главное окно
# -----------------------------------------------------------------------------

class PowerQueryExcelGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1060x760")
        self.minsize(900, 640)

        self.proc: Optional[subprocess.Popen[str]] = None
        self.reader_thread: Optional[threading.Thread] = None
        self.output_queue: queue.Queue[str] = queue.Queue()
        self.start_time: Optional[float] = None

        self.settings = load_settings()

        backend = self.settings.get("backend_script")
        if not backend:
            found = find_backend_script()
            backend = str(found) if found else ""

        self.var_python = tk.StringVar(value=self.settings.get("python", sys.executable))
        self.var_backend = tk.StringVar(value=backend)
        self.var_xlsx = tk.StringVar(value=self.settings.get("xlsx", ""))
        self.var_regnum = tk.StringVar(value=self.settings.get("regnum", ""))
        self.var_output = tk.StringVar(value=self.settings.get("output", ""))
        self.var_cache_dir = tk.StringVar(value=self.settings.get("cache_dir", ""))
        self.var_log_file = tk.StringVar(value=self.settings.get("log_file", ""))
        self.var_debug_dir = tk.StringVar(value=self.settings.get("debug_dir", ""))

        self.var_verbose = tk.BooleanVar(value=self.settings.get("verbose", "1") == "1")
        self.var_debug = tk.BooleanVar(value=self.settings.get("debug", "0") == "1")
        self.var_no_cache = tk.BooleanVar(value=self.settings.get("no_cache", "0") == "1")
        self.var_dump_m = tk.BooleanVar(value=self.settings.get("dump_m", "0") == "1")
        self.var_list_only = tk.BooleanVar(value=False)

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
        self._add_path_row(top, 1, "Выходной файл:", self.var_output, self.choose_output, "Куда сохранить")

        reg_frame = ttk.Frame(top)
        reg_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        reg_frame.columnconfigure(1, weight=1)
        ttk.Label(reg_frame, text="regnum банка:").grid(row=0, column=0, sticky="w")
        reg_entry = ttk.Entry(reg_frame, textvariable=self.var_regnum, width=16)
        reg_entry.grid(row=0, column=1, sticky="w", padx=(8, 6))
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

        advanced = ttk.LabelFrame(top, text="Дополнительные пути", padding=10)
        advanced.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        advanced.columnconfigure(1, weight=1)
        self._add_path_row(advanced, 0, "Python:", self.var_python, self.choose_python, "Выбрать Python")
        self._add_path_row(advanced, 1, "Backend-скрипт:", self.var_backend, self.choose_backend, "Выбрать скрипт")
        self._add_path_row(advanced, 2, "Папка кэша:", self.var_cache_dir, self.choose_cache_dir, "Выбрать папку")
        self._add_path_row(advanced, 3, "Файл лога:", self.var_log_file, self.choose_log_file, "Выбрать файл")
        self._add_path_row(advanced, 4, "Debug-папка:", self.var_debug_dir, self.choose_debug_dir, "Выбрать папку")

        command_frame = ttk.LabelFrame(self, text="Команда", padding=10)
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
            self.var_python,
            self.var_backend,
            self.var_xlsx,
            self.var_regnum,
            self.var_output,
            self.var_cache_dir,
            self.var_log_file,
            self.var_debug_dir,
        ]
        for var in variables:
            var.trace_add("write", lambda *_: self.update_command_preview())
        for var in [self.var_verbose, self.var_debug, self.var_no_cache, self.var_dump_m, self.var_list_only]:
            var.trace_add("write", lambda *_: self.update_command_preview())

    # ------------------------------------------------------------------
    # Выбор файлов/папок
    # ------------------------------------------------------------------

    def choose_xlsx(self) -> None:
        initialdir = str(Path(self.var_xlsx.get()).expanduser().parent) if self.var_xlsx.get() else str(Path.home())
        path = filedialog.askopenfilename(
            title="Выберите исходный Excel-файл",
            initialdir=initialdir,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_xlsx.set(path)
            if not self.var_output.get():
                self.set_auto_output()

    def choose_output(self) -> None:
        initialdir = str(Path(self.var_xlsx.get()).expanduser().parent) if self.var_xlsx.get() else str(Path.home())
        path = filedialog.asksaveasfilename(
            title="Куда сохранить заполненную книгу",
            initialdir=initialdir,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_output.set(path)

    def choose_python(self) -> None:
        path = filedialog.askopenfilename(title="Выберите интерпретатор Python", initialdir=str(Path(sys.executable).parent))
        if path:
            self.var_python.set(path)

    def choose_backend(self) -> None:
        initialdir = str(app_dir())
        path = filedialog.askopenfilename(
            title="Выберите backend-скрипт wtf2_xmlsafe_cellonly.py",
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

    def build_command(self) -> List[str]:
        python_exe = self.var_python.get().strip() or sys.executable
        backend = self.var_backend.get().strip()
        xlsx = self.var_xlsx.get().strip()
        args = [python_exe, backend, xlsx]

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

    def validate_inputs(self) -> bool:
        python_exe = Path(self.var_python.get().strip() or sys.executable).expanduser()
        backend = Path(self.var_backend.get().strip()).expanduser()
        xlsx = Path(self.var_xlsx.get().strip()).expanduser()
        regnum = self.var_regnum.get().strip()

        if not python_exe.exists():
            messagebox.showerror("Ошибка", f"Не найден Python:\n{python_exe}")
            return False
        if not backend.exists():
            messagebox.showerror("Ошибка", f"Не найден backend-скрипт:\n{backend}\n\nПоложите GUI рядом с wtf2_xmlsafe_cellonly.py или выберите скрипт вручную.")
            return False
        if not xlsx.exists():
            messagebox.showerror("Ошибка", f"Не найден Excel-файл:\n{xlsx}")
            return False
        if xlsx.suffix.lower() != ".xlsx":
            if not messagebox.askyesno("Предупреждение", "Файл не имеет расширения .xlsx. Всё равно продолжить?"):
                return False
        if regnum and not re.fullmatch(r"\d+", regnum):
            messagebox.showerror("Ошибка", "regnum должен состоять только из цифр, например 1000 или 1481.")
            return False
        return True

    def update_command_preview(self) -> None:
        command = quote_command(self.build_command())
        self.command_text.configure(state="normal")
        self.command_text.delete("1.0", "end")
        self.command_text.insert("1.0", command)
        self.command_text.configure(state="disabled")

    def set_auto_output(self) -> None:
        self.var_output.set(default_output_path(self.var_xlsx.get().strip(), self.var_regnum.get().strip()))

    def copy_command(self) -> None:
        command = quote_command(self.build_command())
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
        args = self.build_command()
        self.append_log("Команда:\n" + quote_command(args) + "\n\n")

        env = os.environ.copy()
        env.setdefault("PYTHONIOENCODING", "utf-8")
        env.setdefault("PYTHONUTF8", "1")

        try:
            self.proc = subprocess.Popen(
                args,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
                cwd=str(Path(self.var_xlsx.get()).expanduser().parent),
            )
        except Exception as exc:
            self.proc = None
            messagebox.showerror("Не удалось запустить", str(exc))
            return

        self.start_time = time.time()
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.progress.start(12)
        self.status.set("Выполняется...")

        self.reader_thread = threading.Thread(target=self._reader_worker, daemon=True)
        self.reader_thread.start()

    def _reader_worker(self) -> None:
        assert self.proc is not None
        assert self.proc.stdout is not None
        for line in self.proc.stdout:
            self.output_queue.put(line)
        return_code = self.proc.wait()
        self.output_queue.put(f"\n[Процесс завершён с кодом {return_code}]\n")
        self.output_queue.put("__PROCESS_DONE__")

    def _poll_process_output(self) -> None:
        try:
            while True:
                item = self.output_queue.get_nowait()
                if item == "__PROCESS_DONE__":
                    self._on_process_done()
                else:
                    self.append_log(item)
        except queue.Empty:
            pass
        self.after(100, self._poll_process_output)

    def _on_process_done(self) -> None:
        code = self.proc.returncode if self.proc is not None else None
        elapsed = int(time.time() - self.start_time) if self.start_time else 0
        self.proc = None
        self.reader_thread = None
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

    def stop_script(self) -> None:
        if self.proc is None:
            return
        if not messagebox.askyesno("Остановить", "Остановить выполнение скрипта?"):
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
        path = self.var_output.get().strip()
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
            "python": self.var_python.get(),
            "backend_script": self.var_backend.get(),
            "xlsx": self.var_xlsx.get(),
            "regnum": self.var_regnum.get(),
            "output": self.var_output.get(),
            "cache_dir": self.var_cache_dir.get(),
            "log_file": self.var_log_file.get(),
            "debug_dir": self.var_debug_dir.get(),
            "verbose": "1" if self.var_verbose.get() else "0",
            "debug": "1" if self.var_debug.get() else "0",
            "no_cache": "1" if self.var_no_cache.get() else "0",
            "dump_m": "1" if self.var_dump_m.get() else "0",
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
    root = PowerQueryExcelGUI()
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
