from __future__ import annotations

import os
import queue
import re
import shutil
import subprocess
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import yaml


def _application_dir() -> Path:
    """Directory for user-editable files and generated reports.

    In a PyInstaller onedir build ``__file__`` points into ``_internal``.
    User files must instead live next to ProjectTransfer.exe.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _bundle_dir() -> Path:
    """Directory containing bundled read-only resources."""
    raw = getattr(sys, "_MEIPASS", None)
    return Path(raw).resolve() if raw else Path(__file__).resolve().parent


APP_DIR = _application_dir()
BUNDLE_DIR = _bundle_dir()


def _default_file(name: str) -> Path:
    """Prefer an editable file next to the EXE, then a bundled template."""
    editable = APP_DIR / name
    if editable.exists():
        return editable
    bundled = BUNDLE_DIR / name
    if bundled.exists():
        return bundled
    return editable


def _materialize_default(name: str) -> Path:
    """Copy a bundled template next to the EXE when the editable copy is absent."""
    editable = APP_DIR / name
    if editable.exists():
        return editable
    bundled = BUNDLE_DIR / name
    if bundled.is_file() and bundled != editable:
        try:
            APP_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(bundled, editable)
            return editable
        except OSError:
            return bundled
    return editable


def _playwright_environment() -> tuple[dict[str, str], Path]:
    """Return an environment that never points Playwright into PyInstaller _MEI."""
    env = os.environ.copy()
    configured = env.get("PLAYWRIGHT_BROWSERS_PATH", "").strip()
    if configured and configured != "0":
        browser_dir = Path(configured).expanduser()
    else:
        candidates = [APP_DIR / "ms-playwright"]
        local_app_data = env.get("LOCALAPPDATA", "").strip()
        if local_app_data:
            candidates.append(Path(local_app_data) / "ms-playwright")
        user_profile = env.get("USERPROFILE", "").strip()
        if user_profile:
            candidates.append(Path(user_profile) / "AppData" / "Local" / "ms-playwright")
        browser_dir = next((item for item in candidates if item.is_dir()), candidates[0])
        env["PLAYWRIGHT_BROWSERS_PATH"] = str(browser_dir.resolve())
    env.setdefault("PLAYWRIGHT_SKIP_BROWSER_GC", "1")
    # Force UTF-8 for the frozen console worker.  Without this, Windows may
    # encode the pipe as CP866 while the GUI decodes it as UTF-8.
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    return env, browser_dir


class TransferGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Перенос проектов myrosmol.ru")
        self.geometry("1280x850")
        self.minsize(1100, 720)
        self.proc: subprocess.Popen[bytes] | None = None
        self.q: queue.Queue[tuple[str, object]] = queue.Queue()
        self.last_report: Path | None = None

        # Ensure editable defaults exist beside the EXE whenever possible.
        self.default_config = _materialize_default("config_transfer.yaml")
        self.default_source = _materialize_default("исходник.xls")
        self.default_target = _materialize_default("итог.xls")

        self._build()
        self.after(100, self._drain)

    def _build(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self, padding=12)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        self.config_var = tk.StringVar(value=str(self.default_config))
        self.source_var = tk.StringVar(value=str(self.default_source))
        self.target_var = tk.StringVar(value=str(self.default_target))
        rows = [
            ("Конфигурация", self.config_var, "yaml"),
            ("Исходные аккаунты", self.source_var, "xls"),
            ("Аккаунты-получатели", self.target_var, "xls"),
        ]
        for row_index, (label, variable, kind) in enumerate(rows):
            ttk.Label(top, text=label).grid(
                row=row_index, column=0, sticky="w", padx=(0, 8), pady=3
            )
            ttk.Entry(top, textvariable=variable).grid(
                row=row_index, column=1, sticky="ew", pady=3
            )
            ttk.Button(
                top,
                text="Обзор",
                command=lambda v=variable, k=kind: self._browse(v, k),
            ).grid(row=row_index, column=2, padx=(8, 0), pady=3)

        opts = ttk.LabelFrame(self, text="Режим и ускорение", padding=10)
        opts.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        for column in range(8):
            opts.columnconfigure(column, weight=1 if column in (1, 3, 5, 7) else 0)

        self.mode = tk.StringVar(value="upload")
        self.workers = tk.IntVar(value=2)
        self.row_var = tk.StringVar()
        self.limit_var = tk.StringVar()
        self.auto_select = tk.BooleanVar(value=True)
        self.all_accounts = tk.BooleanVar(value=True)
        self.headless = tk.BooleanVar(value=False)
        self.block_heavy = tk.BooleanVar(value=True)
        self.shots = tk.StringVar(value="errors")
        self.error_policy = tk.StringVar(value="ask")
        self.dup_policy = tk.StringVar(value="ask")

        ttk.Label(opts, text="Действие").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            opts,
            textvariable=self.mode,
            state="readonly",
            values=("download", "upload", "all", "parse"),
            width=14,
        ).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Label(opts, text="Потоки").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(opts, from_=1, to=4, textvariable=self.workers, width=5).grid(
            row=0, column=3, sticky="w", padx=5
        )
        ttk.Label(opts, text="Строка XLS").grid(row=0, column=4, sticky="w")
        ttk.Entry(opts, textvariable=self.row_var, width=8).grid(
            row=0, column=5, sticky="ew", padx=5
        )
        ttk.Label(opts, text="Лимит").grid(row=0, column=6, sticky="w")
        ttk.Entry(opts, textvariable=self.limit_var, width=8).grid(
            row=0, column=7, sticky="ew", padx=5
        )
        ttk.Checkbutton(
            opts, text="Автовыбор проекта", variable=self.auto_select
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Checkbutton(opts, text="Все аккаунты", variable=self.all_accounts).grid(
            row=1, column=2, columnspan=2, sticky="w"
        )
        ttk.Checkbutton(opts, text="Скрытый браузер", variable=self.headless).grid(
            row=1, column=4, columnspan=2, sticky="w"
        )
        ttk.Checkbutton(
            opts,
            text="Блокировать изображения/шрифты",
            variable=self.block_heavy,
        ).grid(row=1, column=6, columnspan=2, sticky="w")
        ttk.Label(opts, text="При ошибке").grid(row=2, column=0, sticky="w")
        ttk.Combobox(
            opts,
            textvariable=self.error_policy,
            state="readonly",
            values=("ask", "retry", "skip", "stop"),
        ).grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Label(opts, text="При дубликате").grid(row=2, column=2, sticky="w")
        ttk.Combobox(
            opts,
            textvariable=self.dup_policy,
            state="readonly",
            values=("ask", "existing", "new"),
        ).grid(row=2, column=3, sticky="ew", padx=5)
        ttk.Label(opts, text="Скриншоты").grid(row=2, column=4, sticky="w")
        ttk.Combobox(
            opts,
            textvariable=self.shots,
            state="readonly",
            values=("errors", "all", "none"),
        ).grid(row=2, column=5, sticky="ew", padx=5)

        middle = ttk.Panedwindow(self, orient="vertical")
        middle.grid(row=2, column=0, sticky="nsew", padx=12)
        log_frame = ttk.LabelFrame(middle, text="Журнал выполнения", padding=6)
        result_frame = ttk.LabelFrame(middle, text="Статус", padding=6)
        middle.add(log_frame, weight=4)
        middle.add(result_frame, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log = tk.Text(log_frame, wrap="word", font=("Consolas", 10))
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        answer_frame = ttk.Frame(log_frame)
        answer_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        answer_frame.columnconfigure(0, weight=1)
        self.answer = tk.StringVar()
        answer_entry = ttk.Entry(answer_frame, textvariable=self.answer)
        answer_entry.grid(row=0, column=0, sticky="ew")
        answer_entry.bind("<Return>", lambda _event: self.send_answer())
        ttk.Button(
            answer_frame,
            text="Отправить ответ в программу",
            command=self.send_answer,
        ).grid(row=0, column=1, padx=(6, 0))

        result_frame.columnconfigure(1, weight=1)
        self.status = tk.StringVar(value="Готово к запуску")
        self.progress = ttk.Progressbar(result_frame, mode="indeterminate")
        ttk.Label(result_frame, text="Состояние:").grid(row=0, column=0, sticky="w")
        ttk.Label(result_frame, textvariable=self.status).grid(
            row=0, column=1, sticky="w"
        )
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)

        buttons = ttk.Frame(self, padding=12)
        buttons.grid(row=3, column=0, sticky="ew")
        self.start_btn = ttk.Button(buttons, text="Запустить", command=self.start)
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(
            buttons, text="Остановить", command=self.stop, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=6)
        ttk.Button(
            buttons,
            text="Открыть отчёты",
            command=lambda: self._open(APP_DIR / "logs_transfer"),
        ).pack(side="left", padx=6)
        ttk.Button(
            buttons,
            text="Открыть скачанные проекты",
            command=lambda: self._open(APP_DIR / "скачанные_проекты"),
        ).pack(side="left")
        ttk.Button(buttons, text="Самопроверка", command=self.selftest).pack(
            side="right"
        )

    def _browse(self, variable: tk.StringVar, kind: str) -> None:
        file_types = (
            [("YAML", "*.yaml *.yml")]
            if kind == "yaml"
            else [("Excel 97-2003", "*.xls"), ("Все файлы", "*.*")]
        )
        selected = filedialog.askopenfilename(
            initialdir=APP_DIR, filetypes=file_types
        )
        if selected:
            variable.set(selected)

    def _runtime_config(self) -> Path:
        config_path = Path(self.config_var.get()).expanduser().resolve()
        with config_path.open("r", encoding="utf-8") as stream:
            config = yaml.safe_load(stream) or {}

        config["source_accounts_file"] = str(
            Path(self.source_var.get()).expanduser().resolve()
        )
        config["target_accounts_file"] = str(
            Path(self.target_var.get()).expanduser().resolve()
        )
        config.setdefault("browser", {})["headless"] = bool(self.headless.get())
        config.setdefault("performance", {})["block_heavy_resources"] = bool(
            self.block_heavy.get()
        )
        config["performance"]["screenshots"] = self.shots.get()
        config.setdefault("behavior", {})["interactive_account_selection"] = not bool(
            self.all_accounts.get()
        )
        config["behavior"]["interactive_project_selection"] = not bool(
            self.auto_select.get()
        )

        # Relative path settings must resolve next to the EXE, not inside _internal.
        paths = config.setdefault("paths", {})
        for key, default_name in (
            ("downloaded_projects", "скачанные_проекты"),
            ("screenshots", "screenshots_transfer"),
            ("logs", "logs_transfer"),
        ):
            raw = Path(str(paths.get(key, default_name)))
            if not raw.is_absolute():
                raw = APP_DIR / raw
            paths[key] = str(raw.resolve())

        runtime = APP_DIR / "runtime_gui.yaml"
        runtime.write_text(
            yaml.safe_dump(config, allow_unicode=True, sort_keys=False),
            encoding="utf-8",
        )
        return runtime

    def _worker_command(self, runtime: Path, workers: int) -> list[str]:
        common = [
            "--config",
            str(runtime),
            "--mode",
            self.mode.get(),
            "--workers",
            str(workers),
            "--error-policy",
            self.error_policy.get(),
            "--duplicate-policy",
            self.dup_policy.get(),
        ]

        if getattr(sys, "frozen", False):
            worker_exe = APP_DIR / "ProjectTransferWorker.exe"
            if not worker_exe.is_file():
                raise FileNotFoundError(
                    "Не найден ProjectTransferWorker.exe рядом с ProjectTransfer.exe. "
                    "Запускайте программу из полной папки dist\\ProjectTransfer или "
                    "пересоберите версию 2.6."
                )
            command = [str(worker_exe), *common]
        else:
            python_exe = APP_DIR / ".venv" / "Scripts" / "python.exe"
            if not python_exe.is_file():
                python_exe = Path(sys.executable)
            worker_script = APP_DIR / "optimized_transfer.py"
            command = [str(python_exe), "-u", str(worker_script), *common]

        if self.auto_select.get():
            command.append("--auto-select")
        if self.all_accounts.get():
            command.append("--all-accounts")
        if self.row_var.get().strip():
            command.extend(["--row", self.row_var.get().strip()])
        if self.limit_var.get().strip():
            command.extend(["--limit", self.limit_var.get().strip()])
        return command

    def start(self) -> None:
        if self.proc and self.proc.poll() is None:
            return
        try:
            runtime = self._runtime_config()
            workers = max(1, min(4, int(self.workers.get())))
            if workers > 1 and (
                self.error_policy.get() == "ask"
                or self.dup_policy.get() == "ask"
                or (self.mode.get() == "download" and not self.auto_select.get())
            ):
                workers = 1
                self._append(
                    "Интерактивные решения включены: число потоков "
                    "автоматически снижено до 1.\n"
                )
            command = self._worker_command(runtime, workers)
        except Exception as exc:
            messagebox.showerror("Ошибка конфигурации", str(exc))
            return

        self.log.delete("1.0", "end")
        self._append("Команда: " + " ".join(command) + "\n\n")
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        worker_env, browser_dir = _playwright_environment()
        self._append(f"Playwright Chromium: {browser_dir}\n\n")
        try:
            self.proc = subprocess.Popen(
                command,
                cwd=APP_DIR,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=False,
                bufsize=0,
                creationflags=creation_flags,
                env=worker_env,
            )
        except Exception as exc:
            messagebox.showerror("Ошибка запуска", str(exc))
            return

        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.status.set("Выполняется")
        self.progress.start(12)
        threading.Thread(target=self._reader, daemon=True).start()

    @staticmethod
    def _decode_worker_output(raw: bytes) -> str:
        """Decode worker output from UTF-8, CP866 or CP1251.

        UTF-8 is authoritative for v2.7 workers.  Legacy Windows encodings are
        single-byte and therefore never raise decoding errors, so a small
        readability score is used to choose between CP866 and CP1251.
        """
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            pass

        russian_letters = set(
            "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"
            "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
        )
        common_fragments = (
            "проект", "отч", "выбер", "ошиб", "файл", "аккаун",
            "брауз", "создан", "скачан", "загруз", "перенос",
            "строк", "дата", "успеш",
        )

        def readability(text: str) -> int:
            russian = sum(char in russian_letters for char in text)
            foreign_cyrillic = sum(
                0x0400 <= ord(char) <= 0x04FF and char not in russian_letters
                for char in text
            )
            box_drawing = sum(0x2500 <= ord(char) <= 0x257F for char in text)
            replacements = text.count("\ufffd")
            lower = text.lower()
            familiar = sum(20 for fragment in common_fragments if fragment in lower)
            return (
                russian * 2
                + familiar
                - foreign_cyrillic * 8
                - box_drawing * 10
                - replacements * 30
            )

        candidates = [raw.decode("cp866"), raw.decode("cp1251")]
        return max(candidates, key=readability)

    def _reader(self) -> None:
        assert self.proc and self.proc.stdout
        while True:
            raw = self.proc.stdout.readline()
            if not raw:
                break
            self.q.put(("line", self._decode_worker_output(raw)))
        code = self.proc.wait()
        self.q.put(("done", code))

    def _drain(self) -> None:
        try:
            while True:
                kind, value = self.q.get_nowait()
                if kind == "line":
                    line = str(value)
                    self._append(line)
                    match = re.search(r"Отчёт XLS:\s*(.+)", line)
                    if match:
                        self.last_report = Path(match.group(1).strip())
                else:
                    self._finished(int(value))
        except queue.Empty:
            pass
        self.after(100, self._drain)

    def _append(self, text: str) -> None:
        self.log.insert("end", text)
        self.log.see("end")

    def send_answer(self) -> None:
        if not self.proc or self.proc.poll() is not None or not self.proc.stdin:
            return
        answer = self.answer.get()
        self.proc.stdin.write((answer + "\n").encode("utf-8"))
        self.proc.stdin.flush()
        self._append(f"\n> {answer}\n")
        self.answer.set("")

    def stop(self) -> None:
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
            self.status.set("Остановка...")

    def _finished(self, code: int) -> None:
        self.progress.stop()
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.status.set("Завершено успешно" if code == 0 else f"Завершено, код {code}")
        if self.last_report and self.last_report.exists():
            self._append(f"\nГотовый отчёт: {self.last_report}\n")

    def _open(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except AttributeError:
            subprocess.Popen(["xdg-open", str(path)])

    def selftest(self) -> None:
        checks: list[tuple[str, bool, str]] = []
        config_path = Path(self.config_var.get()).expanduser()
        source_path = Path(self.source_var.get()).expanduser()
        target_path = Path(self.target_var.get()).expanduser()
        checks.append(("Конфигурация", config_path.is_file(), str(config_path)))
        checks.append(("Исходный XLS", source_path.is_file(), str(source_path)))
        checks.append(("Итоговый XLS", target_path.is_file(), str(target_path)))

        if getattr(sys, "frozen", False):
            worker = APP_DIR / "ProjectTransferWorker.exe"
            checks.append(("Процесс переноса", worker.is_file(), str(worker)))
        else:
            worker = APP_DIR / "optimized_transfer.py"
            checks.append(("Ускоренное ядро", worker.is_file(), str(worker)))

        _, browser_dir = _playwright_environment()
        browser_exes = list(browser_dir.glob("**/chrome.exe")) if browser_dir.is_dir() else []
        checks.append((
            "Chromium Playwright",
            bool(browser_exes),
            str(browser_exes[0] if browser_exes else browser_dir),
        ))

        try:
            with config_path.open("r", encoding="utf-8") as stream:
                yaml.safe_load(stream)
            checks.append(("Разбор YAML", True, "OK"))
        except Exception as exc:
            checks.append(("Разбор YAML", False, str(exc)))

        lines = ["\n--- САМОПРОВЕРКА ---"]
        for label, passed, detail in checks:
            lines.append(f"{'OK' if passed else 'ERROR'}: {label} — {detail}")
        self._append("\n".join(lines) + "\n")
        passed_all = all(item[1] for item in checks)
        messagebox.showinfo(
            "Самопроверка",
            "Пройдена" if passed_all else "Обнаружены ошибки; см. журнал",
        )


if __name__ == "__main__":
    TransferGUI().mainloop()
