"""
doc_to_pdf_converter.py
=======================
Конвертер файлов .doc и .docx в PDF.
GUI переработан на customtkinter для современного внешнего вида.

Зависимости:
  pip install docx2pdf customtkinter

На Linux дополнительно нужен LibreOffice:
  sudo apt install libreoffice
"""

import sys
import logging
import platform
import subprocess
import threading
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

# ══════════════════════════════════════════════
# Настройка логирования: файл log.txt + консоль
# ══════════════════════════════════════════════
LOG_FILE = Path("log.txt")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════
# Конвертация одного файла (логика не изменена)
# ══════════════════════════════════════════════
def convert_file(src: Path, dst: Path) -> bool:
    """
    Конвертирует один .doc/.docx в PDF.
    Windows/macOS — через docx2pdf (Microsoft Word).
    Linux         — через LibreOffice.
    Возвращает True при успехе, False при ошибке.
    """
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        os_name = platform.system()

        if os_name in ("Windows", "Darwin"):
            try:
                from docx2pdf import convert  # type: ignore
            except ImportError:
                raise RuntimeError(
                    "Библиотека docx2pdf не установлена. "
                    "Выполните: pip install docx2pdf"
                )
            convert(str(src), str(dst))

        else:
            result = subprocess.run(
                [
                    "libreoffice", "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(dst.parent),
                    str(src),
                ],
                capture_output=True, text=True, timeout=120,
            )
            if result.returncode != 0:
                raise RuntimeError(f"LibreOffice error:\n{result.stderr}")
            generated = dst.parent / (src.stem + ".pdf")
            if generated != dst and generated.exists():
                generated.rename(dst)

        logger.info("OK: %s -> %s", src, dst)
        return True

    except Exception as exc:  # noqa: BLE001
        logger.error("ОШИБКА: %s — %s", src, exc)
        return False


# ══════════════════════════════════════════════
# Поиск .doc/.docx рекурсивно (не изменено)
# ══════════════════════════════════════════════
def find_docs(folder: Path) -> list[Path]:
    """Возвращает все .doc и .docx в папке и вложенных папках."""
    return [
        p for p in folder.rglob("*")
        if p.suffix.lower() in (".doc", ".docx") and p.is_file()
    ]


# ══════════════════════════════════════════════
# Главное окно приложения — customtkinter
# ══════════════════════════════════════════════
class ConverterApp(ctk.CTk):
    # Цвета для записей в лог-виджете
    LOG_OK   = "#4ade80"   # зелёный — успех
    LOG_ERR  = "#f87171"   # красный — ошибка
    LOG_INFO = "#94a3b8"   # серый   — информация

    def __init__(self):
        super().__init__()

        # Тема: системная (авто light/dark), синий акцент
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.title("DOC / DOCX  →  PDF Конвертер")
        self.resizable(False, False)

        # Состояние выбора
        self.selected_files:  list[Path] = []
        self.selected_folder: Path | None = None

        self._build_ui()

    # ─────────────────────────────────────────
    # Построение интерфейса
    # ─────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        # ── Шапка ──────────────────────────────
        header = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(24, 8))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="DOC → PDF",
            font=ctk.CTkFont(size=26, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text="Конвертируйте файлы Word в PDF быстро и просто",
            font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        # ── Карточка выбора источника ──────────
        src_card = ctk.CTkFrame(self, corner_radius=12)
        src_card.grid(row=1, column=0, sticky="ew", padx=24, pady=8)
        src_card.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            src_card,
            text="ИСТОЧНИК",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray45", "gray65"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 8))

        # Кнопка «Выбрать файлы» (основной стиль)
        ctk.CTkButton(
            src_card,
            text="Выбрать файлы",
            height=40,
            command=self._pick_files,
        ).grid(row=1, column=0, padx=(16, 6), pady=(0, 16), sticky="ew")

        # Кнопка «Выбрать папку» (нейтральный стиль)
        ctk.CTkButton(
            src_card,
            text="Выбрать папку",
            height=40,
            fg_color=("gray80", "gray30"),
            text_color=("gray15", "gray90"),
            hover_color=("gray70", "gray40"),
            command=self._pick_folder,
        ).grid(row=1, column=1, padx=(6, 16), pady=(0, 16), sticky="ew")

        # Статус выбора
        self.lbl_selection = ctk.CTkLabel(
            src_card,
            text="Ничего не выбрано",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
            wraplength=460,
            justify="left",
        )
        self.lbl_selection.grid(
            row=2, column=0, columnspan=2,
            sticky="w", padx=16, pady=(0, 14),
        )

        # ── Кнопка «Начать конвертацию» ────────
        self.btn_convert = ctk.CTkButton(
            self,
            text="▶   Начать конвертацию",
            height=46,
            font=ctk.CTkFont(size=14, weight="bold"),
            state="disabled",
            command=self._start_conversion,
        )
        self.btn_convert.grid(row=2, column=0, sticky="ew", padx=24, pady=8)

        # ── Прогресс-бар ───────────────────────
        self.progress = ctk.CTkProgressBar(self, height=8, corner_radius=4)
        self.progress.set(0)
        self.progress.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 8))

        # ── Заголовок лога ──────────────────────
        ctk.CTkLabel(
            self,
            text="ЖУРНАЛ КОНВЕРТАЦИИ",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray45", "gray65"),
        ).grid(row=4, column=0, sticky="w", padx=24, pady=(4, 2))

        # ── Текстовый лог ───────────────────────
        self.log_box = ctk.CTkTextbox(
            self,
            height=180,
            font=ctk.CTkFont(family="Courier", size=12),
            corner_radius=10,
            wrap="word",
            state="disabled",
            activate_scrollbars=True,
        )
        self.log_box.grid(row=5, column=0, sticky="ew", padx=24, pady=(0, 8))

        # ── Подвал ─────────────────────────────
        ctk.CTkLabel(
            self,
            text=f"Подробный лог: {LOG_FILE.resolve()}",
            font=ctk.CTkFont(size=10),
            text_color=("gray55", "gray55"),
        ).grid(row=6, column=0, pady=(0, 16))

        self.geometry("540x610")

    # ─────────────────────────────────────────
    # Обработчик «Выбрать файлы»
    # ─────────────────────────────────────────
    def _pick_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите файлы .doc / .docx",
            filetypes=[("Word документы", "*.doc *.docx"), ("Все файлы", "*.*")],
        )
        if not files:
            return

        self.selected_files  = [Path(f) for f in files]
        self.selected_folder = None
        count   = len(self.selected_files)
        preview = "\n  ".join(f.name for f in self.selected_files[:4])
        suffix  = f"\n  … ещё {count - 4}" if count > 4 else ""

        self.lbl_selection.configure(
            text=f"Выбрано файлов: {count}\n  {preview}{suffix}",
            text_color=("#16a34a", "#4ade80"),
        )
        self.btn_convert.configure(state="normal")

    # ─────────────────────────────────────────
    # Обработчик «Выбрать папку»
    # ─────────────────────────────────────────
    def _pick_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с документами")
        if not folder:
            return

        self.selected_folder = Path(folder)
        self.selected_files  = []
        count = len(find_docs(self.selected_folder))

        if count == 0:
            self.lbl_selection.configure(
                text="В папке нет файлов .doc / .docx",
                text_color=("#b45309", "#fbbf24"),
            )
            self.btn_convert.configure(state="disabled")
        else:
            self.lbl_selection.configure(
                text=f"Папка: {self.selected_folder.name}\n"
                     f"  Найдено файлов: {count}",
                text_color=("#16a34a", "#4ade80"),
            )
            self.btn_convert.configure(state="normal")

    # ─────────────────────────────────────────
    # Запуск конвертации
    # ─────────────────────────────────────────
    def _start_conversion(self):
        """Формирует список задач и запускает фоновый поток конвертации."""
        self._set_ui_busy(True)
        self._log_clear()

        tasks: list[tuple[Path, Path]] = []

        if self.selected_folder:
            # Режим папки: выходные PDF зеркалируют структуру в «<имя>_pdf»
            src_root = self.selected_folder
            dst_root = src_root.parent / (src_root.name + "_pdf")
            for src in find_docs(src_root):
                rel = src.relative_to(src_root)
                tasks.append((src, dst_root / rel.with_suffix(".pdf")))
        else:
            # Режим файлов: PDF рядом с оригиналом
            for src in self.selected_files:
                tasks.append((src, src.with_suffix(".pdf")))

        if not tasks:
            messagebox.showwarning("Нет файлов", "Нечего конвертировать.")
            self._set_ui_busy(False)
            return

        threading.Thread(
            target=self._run_conversion, args=(tasks,), daemon=True
        ).start()

    def _run_conversion(self, tasks: list[tuple[Path, Path]]):
        """Выполняет конвертацию в фоновом потоке."""
        total   = len(tasks)
        success = 0
        fail    = 0

        self._log(f"Начало конвертации — файлов: {total}\n", self.LOG_INFO)

        for i, (src, dst) in enumerate(tasks, 1):
            self._log(f"[{i}/{total}]  {src.name}  …  ", self.LOG_INFO, newline=False)

            if convert_file(src, dst):
                success += 1
                self._log("OK", self.LOG_OK)
            else:
                fail += 1
                self._log("ОШИБКА", self.LOG_ERR)

            # Обновляем прогресс-бар (значение 0.0–1.0)
            self.after(0, self.progress.set, i / total)

        # Итоговая строка
        self._log(
            f"\n{'─' * 44}\n"
            f"Готово!  Успешно: {success}   Ошибок: {fail}",
            self.LOG_INFO,
        )
        logger.info("Готово. Успешно: %d, ошибок: %d", success, fail)

        self.after(0, self._set_ui_busy, False)

        # Всплывающее уведомление
        if fail == 0:
            self.after(0, messagebox.showinfo,
                       "Готово",
                       f"Все {success} файл(ов) успешно конвертированы!")
        else:
            self.after(0, messagebox.showwarning,
                       "Завершено с ошибками",
                       f"Успешно: {success}\nОшибок: {fail}\n\n"
                       f"Подробности: {LOG_FILE.resolve()}")

    # ─────────────────────────────────────────
    # Управление состоянием UI
    # ─────────────────────────────────────────
    def _set_ui_busy(self, busy: bool):
        """Блокирует/разблокирует кнопку конвертации."""
        self.btn_convert.configure(state="disabled" if busy else "normal")
        if not busy:
            self.progress.set(0)

    # ─────────────────────────────────────────
    # Вспомогательные методы лог-виджета
    # ─────────────────────────────────────────
    def _log(self, text: str, color: str = "#94a3b8", newline: bool = True):
        """Потокобезопасная запись строки в лог-виджет."""
        msg = text + ("\n" if newline else "")
        self.after(0, self._append_log, msg, color)

    def _append_log(self, msg: str, color: str):
        """Вставляет текст с цветом. Вызывается только из главного потока."""
        self.log_box.configure(state="normal")
        tag = f"c_{color.lstrip('#')}"
        self.log_box.tag_config(tag, foreground=color)
        self.log_box.insert("end", msg, tag)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        """Очищает лог и сбрасывает прогресс-бар."""
        self.log_box.configure(state="normal")
        self.log_box.delete("0.0", "end")
        self.log_box.configure(state="disabled")
        self.progress.set(0)


# ══════════════════════════════════════════════
# Точка входа
# ══════════════════════════════════════════════
if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
