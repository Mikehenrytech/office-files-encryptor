# main.py
# PySide6 GUI to password-protect PDF and OOXML (docx/xlsx/pptx) documents.
# - Drag & drop + file picker
# - Multi-language (ES/EN) via dict (no Qt Linguist required)
# - Magic-bytes / structure validation (PDF header, OOXML zip + structure)
# - Parallel processing using QThreadPool + QRunnable
# - Output in same folder as original with _protected suffix OR overwrite original (optional)
#
# Security notes:
# - Overwrite mode replaces original file after writing to a temp file in same directory.
# - This is NOT a secure wipe of original bytes (NTFS journaling, SSD wear leveling, etc.).
#   If you need secure deletion, it must be handled at filesystem/device level.

import os
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from PySide6.QtCore import Qt, QUrl, QObject, Signal, QRunnable, QThreadPool
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QCheckBox,
    QProgressBar,
)

import pikepdf
import msoffcrypto


SUPPORTED_PDF = {".pdf"}
SUPPORTED_OOXML = {".docx", ".xlsx", ".pptx"}
ALL_SUPPORTED_EXT = SUPPORTED_PDF | SUPPORTED_OOXML

PDF_MAGIC_PREFIX = b"%PDF-"


@dataclass
class FileEntry:
    path: Path
    detected_kind: str  # "PDF" | "WORD" | "EXCEL" | "POWERPOINT" | "UNKNOWN"
    status: str         # "PENDING" | "RUNNING" | "DONE" | "FAILED" | "UNSUPPORTED"
    message: str = ""


class I18N:
    STRINGS = {
        "en": {
            "app_title": "Document Password Protector",
            "toolbar_language": "Language",
            "add_files": "Add files…",
            "remove_selected": "Remove selected",
            "clear_list": "Clear",
            "password": "Password (open):",
            "show_password": "Show",
            "protect": "Protect",
            "overwrite": "Overwrite originals (replace files)",
            "overwrite_warn_title": "Overwrite warning",
            "overwrite_warn_body": (
                "This will REPLACE the original files with password-protected versions.\n\n"
                "Important:\n"
                "- This is not a secure wipe of the original file bytes.\n"
                "- If something goes wrong, your original may be lost.\n\n"
                "Do you want to enable overwrite mode?"
            ),
            "status_pending": "PENDING",
            "status_running": "RUNNING",
            "status_done": "DONE",
            "status_failed": "FAILED",
            "status_unsupported": "UNSUPPORTED",
            "kind_pdf": "PDF",
            "kind_word": "WORD",
            "kind_excel": "EXCEL",
            "kind_powerpoint": "POWERPOINT",
            "kind_unknown": "UNKNOWN",
            "err_no_password": "Please enter a password.",
            "err_no_files": "Please add at least one file.",
            "dlg_select_files": "Select files",
            "msg_done_title": "Completed",
            "msg_done_body": "Finished processing. Done: {done}, Failed: {failed}, Unsupported: {unsupported}",
            "hint_dragdrop": "Tip: You can drag & drop files here.",
            "err_magic_mismatch": "File content does not match extension/type (magic/structure mismatch).",
            "err_zip_invalid": "Not a valid ZIP container for OOXML.",
            "err_ooxml_structure": "ZIP is valid but not a supported OOXML document structure.",
            "progress": "Progress:",
            "busy": "Processing…",
            "idle": "Idle",
            "saved_as": "Saved: {path}",
            "overwritten": "Overwritten: {path}",
        },
        "es": {
            "app_title": "Protector de Contraseña para Documentos",
            "toolbar_language": "Idioma",
            "add_files": "Añadir archivos…",
            "remove_selected": "Quitar seleccionados",
            "clear_list": "Limpiar",
            "password": "Contraseña (apertura):",
            "show_password": "Mostrar",
            "protect": "Proteger",
            "overwrite": "Sobrescribir originales (reemplazar archivos)",
            "overwrite_warn_title": "Aviso de sobrescritura",
            "overwrite_warn_body": (
                "Esto REEMPLAZARÁ los archivos originales por versiones protegidas con contraseña.\n\n"
                "Importante:\n"
                "- No es un borrado seguro de los bytes originales.\n"
                "- Si algo falla, podrías perder el original.\n\n"
                "¿Quieres activar el modo sobrescritura?"
            ),
            "status_pending": "PENDIENTE",
            "status_running": "EN PROCESO",
            "status_done": "OK",
            "status_failed": "FALLO",
            "status_unsupported": "NO SOPORTADO",
            "kind_pdf": "PDF",
            "kind_word": "WORD",
            "kind_excel": "EXCEL",
            "kind_powerpoint": "POWERPOINT",
            "kind_unknown": "DESCONOCIDO",
            "err_no_password": "Introduce una contraseña.",
            "err_no_files": "Añade al menos un archivo.",
            "dlg_select_files": "Seleccionar archivos",
            "msg_done_title": "Completado",
            "msg_done_body": "Proceso finalizado. OK: {done}, Fallos: {failed}, No soportados: {unsupported}",
            "hint_dragdrop": "Consejo: puedes arrastrar y soltar archivos aquí.",
            "err_magic_mismatch": "El contenido del archivo no coincide con la extensión/tipo (mismatch de magic/estructura).",
            "err_zip_invalid": "No es un contenedor ZIP válido para OOXML.",
            "err_ooxml_structure": "El ZIP es válido pero no corresponde a una estructura OOXML soportada.",
            "progress": "Progreso:",
            "busy": "Procesando…",
            "idle": "En espera",
            "saved_as": "Guardado: {path}",
            "overwritten": "Sobrescrito: {path}",
        },
    }

    def __init__(self, lang: str = "en"):
        self.lang = lang if lang in self.STRINGS else "en"

    def set_lang(self, lang: str) -> None:
        if lang in self.STRINGS:
            self.lang = lang

    def t(self, key: str) -> str:
        return self.STRINGS[self.lang].get(key, key)


# ---------------- Magic/Structure validation ----------------

def detect_real_type(path: Path) -> Tuple[str, Optional[str]]:
    """
    Returns: (kind, error_message_if_any)
    kind in {"PDF","WORD","EXCEL","POWERPOINT","UNKNOWN"}

    Validates magic bytes / container structure.
    - PDF: begins with %PDF-
    - OOXML: ZIP + [Content_Types].xml + expected directory (word/xl/ppt)
    """
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except Exception as e:
        return "UNKNOWN", f"{type(e).__name__}: {e}"

    if head.startswith(PDF_MAGIC_PREFIX):
        return "PDF", None

    # OOXML (docx/xlsx/pptx) are ZIP-based
    try:
        with zipfile.ZipFile(path, "r") as z:
            names = z.namelist()
            if "[Content_Types].xml" not in names:
                return "UNKNOWN", None

            has_word = any(n.startswith("word/") for n in names)
            has_xl = any(n.startswith("xl/") for n in names)
            has_ppt = any(n.startswith("ppt/") for n in names)

            if has_word and not (has_xl or has_ppt):
                return "WORD", None
            if has_xl and not (has_word or has_ppt):
                return "EXCEL", None
            if has_ppt and not (has_word or has_xl):
                return "POWERPOINT", None

            # Best-effort if multiple dirs exist
            if has_word:
                return "WORD", None
            if has_xl:
                return "EXCEL", None
            if has_ppt:
                return "POWERPOINT", None

            return "UNKNOWN", None
    except zipfile.BadZipFile:
        return "UNKNOWN", None
    except Exception as e:
        return "UNKNOWN", f"{type(e).__name__}: {e}"


def extension_matches_kind(ext: str, kind: str) -> bool:
    ext = ext.lower()
    if kind == "PDF":
        return ext == ".pdf"
    if kind == "WORD":
        return ext == ".docx"
    if kind == "EXCEL":
        return ext == ".xlsx"
    if kind == "POWERPOINT":
        return ext == ".pptx"
    return False


def is_supported_kind(kind: str) -> bool:
    return kind in {"PDF", "WORD", "EXCEL", "POWERPOINT"}


def safe_output_path_same_folder(src: Path) -> Path:
    """
    Create protected file in the SAME directory as original.
    Example:
        report.pdf -> report_protected.pdf
        report_protected.pdf exists -> report_protected_2.pdf
    """
    parent = src.parent
    base = src.stem
    ext = src.suffix
    candidate = parent / f"{base}_protected{ext}"
    if not candidate.exists():
        return candidate
    i = 2
    while True:
        candidate = parent / f"{base}_protected_{i}{ext}"
        if not candidate.exists():
            return candidate
        i += 1


# ---------------- Protection implementations ----------------

def protect_pdf(src: Path, dst: Path, password: str) -> None:
    with pikepdf.open(str(src)) as pdf:
        pdf.save(
            str(dst),
            encryption=pikepdf.Encryption(
                user=password,
                owner=password,
                R=6,  # AES-256
            ),
        )


def protect_ooxml(src: Path, dst: Path, password: str) -> None:
    with open(src, "rb") as f_in:
        office_file = msoffcrypto.OfficeFile(f_in)
        with open(dst, "wb") as f_out:
            office_file.encrypt(password=password, outfile=f_out)


def atomic_replace(src_temp: Path, dst_final: Path) -> None:
    """
    Replace dst_final with src_temp atomically where possible.
    On Windows, os.replace is atomic for same filesystem.
    """
    os.replace(str(src_temp), str(dst_final))


# ---------------- Worker (thread pool) ----------------

class WorkerSignals(QObject):
    started = Signal(int)  # index
    finished = Signal(int, str, str)  # index, status, message
    progress = Signal(int, int)  # done, total


class ProtectWorker(QRunnable):
    """
    Worker does:
    - Real type detection by magic/structure
    - Validate that extension matches (strict)
    - Encrypt and write protected file
    - Either:
        A) Save as <name>_protected.ext in same folder
        B) Overwrite original safely: write to temp file then replace original
    """
    def __init__(self, index: int, entry: FileEntry, password: str,
                 overwrite: bool, total: int, done_counter: List[int], signals: WorkerSignals):
        super().__init__()
        self.index = index
        self.entry = entry
        self.password = password
        self.overwrite = overwrite
        self.total = total
        self.done_counter = done_counter  # mutable single-item list
        self.signals = signals

    def run(self):
        self.signals.started.emit(self.index)

        try:
            real_kind, _ = detect_real_type(self.entry.path)

            if not is_supported_kind(real_kind):
                ext = self.entry.path.suffix.lower()
                if ext in SUPPORTED_OOXML:
                    try:
                        zipfile.ZipFile(self.entry.path, "r").close()
                        self.signals.finished.emit(self.index, "FAILED", "ZIP valid but not supported OOXML structure")
                    except zipfile.BadZipFile:
                        self.signals.finished.emit(self.index, "FAILED", "Not a valid ZIP container for OOXML")
                    return
                self.signals.finished.emit(self.index, "UNSUPPORTED", "Unknown/unsupported file type")
                return

            if not extension_matches_kind(self.entry.path.suffix, real_kind):
                self.signals.finished.emit(self.index, "FAILED", "Magic/structure mismatch with extension")
                return

            src = self.entry.path

            if self.overwrite:
                # Write encrypted output to temp file in same directory, then replace original.
                temp_out = src.with_name(src.name + ".tmp_protecting")
                # Ensure we don't collide with an existing temp file
                if temp_out.exists():
                    try:
                        temp_out.unlink()
                    except Exception:
                        # If can't delete, choose a unique temp name
                        temp_out = src.with_name(src.name + f".tmp_protecting_{os.getpid()}")

                if real_kind == "PDF":
                    protect_pdf(src, temp_out, self.password)
                else:
                    protect_ooxml(src, temp_out, self.password)

                atomic_replace(temp_out, src)
                self.signals.finished.emit(self.index, "DONE", f"OVERWRITTEN::{str(src)}")

            else:
                dst = safe_output_path_same_folder(src)
                if real_kind == "PDF":
                    protect_pdf(src, dst, self.password)
                else:
                    protect_ooxml(src, dst, self.password)

                self.signals.finished.emit(self.index, "DONE", f"SAVED::{str(dst)}")

        except Exception as e:
            msg = f"{type(e).__name__}: {str(e)[:220]}"
            self.signals.finished.emit(self.index, "FAILED", msg)

        finally:
            self.done_counter[0] += 1
            self.signals.progress.emit(self.done_counter[0], self.total)


# ---------------- UI ----------------

class DropListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setSelectionMode(QListWidget.ExtendedSelection)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            return super().dropEvent(event)
        self.parent().handle_dropped_urls(event.mimeData().urls())
        event.acceptProposedAction()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.i18n = I18N("es")  # default ES
        self.entries: List[FileEntry] = []

        self.pool = QThreadPool.globalInstance()
        self.signals = WorkerSignals()
        self.signals.started.connect(self.on_worker_started)
        self.signals.finished.connect(self.on_worker_finished)
        self.signals.progress.connect(self.on_worker_progress)

        self._running = False
        self._total = 0
        self._done = 0

        self.setWindowTitle(self.i18n.t("app_title"))
        self.resize(980, 610)

        # Toolbar
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        self.lang_label = QLabel(self.i18n.t("toolbar_language") + ": ")
        toolbar.addWidget(self.lang_label)

        self.lang_combo = QComboBox()
        self.lang_combo.addItem("Español", "es")
        self.lang_combo.addItem("English", "en")
        self.lang_combo.setCurrentIndex(0)
        self.lang_combo.currentIndexChanged.connect(self.on_language_changed)
        toolbar.addWidget(self.lang_combo)

        toolbar.addSeparator()

        add_action = QAction(self.i18n.t("add_files"), self)
        add_action.triggered.connect(self.on_add_files)
        toolbar.addAction(add_action)

        remove_action = QAction(self.i18n.t("remove_selected"), self)
        remove_action.triggered.connect(self.on_remove_selected)
        toolbar.addAction(remove_action)

        clear_action = QAction(self.i18n.t("clear_list"), self)
        clear_action.triggered.connect(self.on_clear)
        toolbar.addAction(clear_action)

        self._toolbar_actions = {"add": add_action, "remove": remove_action, "clear": clear_action}

        # Central
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        self.hint = QLabel(self.i18n.t("hint_dragdrop"))
        self.hint.setAlignment(Qt.AlignLeft)
        layout.addWidget(self.hint)

        self.list_widget = DropListWidget(self)
        layout.addWidget(self.list_widget, stretch=1)

        # Password row
        pw_row = QHBoxLayout()
        self.pw_label = QLabel(self.i18n.t("password"))
        self.pw_input = QLineEdit()
        self.pw_input.setEchoMode(QLineEdit.Password)
        self.show_pw = QCheckBox(self.i18n.t("show_password"))
        self.show_pw.stateChanged.connect(self.on_toggle_pw)
        pw_row.addWidget(self.pw_label)
        pw_row.addWidget(self.pw_input, stretch=1)
        pw_row.addWidget(self.show_pw)
        layout.addLayout(pw_row)

        # Overwrite option row
        ow_row = QHBoxLayout()
        self.overwrite_cb = QCheckBox(self.i18n.t("overwrite"))
        self.overwrite_cb.stateChanged.connect(self.on_overwrite_toggled)
        ow_row.addWidget(self.overwrite_cb)
        ow_row.addStretch(1)
        layout.addLayout(ow_row)

        # Progress row
        prog_row = QHBoxLayout()
        self.progress_label = QLabel(self.i18n.t("progress"))
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setValue(0)
        self.state_label = QLabel(self.i18n.t("idle"))
        prog_row.addWidget(self.progress_label)
        prog_row.addWidget(self.progress_bar, stretch=1)
        prog_row.addWidget(self.state_label)
        layout.addLayout(prog_row)

        # Actions row
        act_row = QHBoxLayout()
        act_row.addStretch(1)
        self.protect_btn = QPushButton(self.i18n.t("protect"))
        self.protect_btn.clicked.connect(self.on_protect)
        act_row.addWidget(self.protect_btn)
        layout.addLayout(act_row)

        self.refresh_ui_texts()
        self.refresh_list_items()

    # -------- i18n --------
    def on_language_changed(self):
        lang = self.lang_combo.currentData()
        self.i18n.set_lang(lang)
        self.refresh_ui_texts()
        self.refresh_list_items()

    def refresh_ui_texts(self):
        self.setWindowTitle(self.i18n.t("app_title"))
        self.lang_label.setText(self.i18n.t("toolbar_language") + ": ")
        self._toolbar_actions["add"].setText(self.i18n.t("add_files"))
        self._toolbar_actions["remove"].setText(self.i18n.t("remove_selected"))
        self._toolbar_actions["clear"].setText(self.i18n.t("clear_list"))
        self.pw_label.setText(self.i18n.t("password"))
        self.show_pw.setText(self.i18n.t("show_password"))
        self.protect_btn.setText(self.i18n.t("protect"))
        self.hint.setText(self.i18n.t("hint_dragdrop"))
        self.progress_label.setText(self.i18n.t("progress"))
        self.state_label.setText(self.i18n.t("busy") if self._running else self.i18n.t("idle"))
        self.overwrite_cb.setText(self.i18n.t("overwrite"))

    # -------- DnD / file dialog --------
    def handle_dropped_urls(self, urls: List[QUrl]):
        if self._running:
            return
        paths = []
        for url in urls:
            if url.isLocalFile():
                paths.append(Path(url.toLocalFile()))
        self.add_files(paths)

    def on_add_files(self):
        if self._running:
            return
        files, _ = QFileDialog.getOpenFileNames(
            self,
            self.i18n.t("dlg_select_files"),
            str(Path.home()),
            "Documents (*.pdf *.docx *.xlsx *.pptx);;All files (*.*)",
        )
        self.add_files([Path(f) for f in files])

    def add_files(self, paths: List[Path]):
        for p in paths:
            if not p.exists() or not p.is_file():
                continue

            rp = p.resolve()
            if any(e.path.resolve() == rp for e in self.entries):
                continue

            kind_guess = self.kind_from_ext(p.suffix.lower())
            status = "PENDING" if p.suffix.lower() in ALL_SUPPORTED_EXT else "UNSUPPORTED"
            msg = "" if status == "PENDING" else "Unsupported extension"
            self.entries.append(FileEntry(path=p, detected_kind=kind_guess, status=status, message=msg))

        self.refresh_list_items()

    def on_remove_selected(self):
        if self._running:
            return
        selected = self.list_widget.selectedItems()
        if not selected:
            return
        idxs = sorted([self.list_widget.row(it) for it in selected], reverse=True)
        for i in idxs:
            if 0 <= i < len(self.entries):
                self.entries.pop(i)
        self.refresh_list_items()

    def on_clear(self):
        if self._running:
            return
        self.entries.clear()
        self.list_widget.clear()

    # -------- UI helpers --------
    def kind_from_ext(self, ext: str) -> str:
        if ext == ".pdf":
            return "PDF"
        if ext == ".docx":
            return "WORD"
        if ext == ".xlsx":
            return "EXCEL"
        if ext == ".pptx":
            return "POWERPOINT"
        return "UNKNOWN"

    def refresh_list_items(self):
        self.list_widget.clear()
        for e in self.entries:
            self.list_widget.addItem(self.render_entry(e))

    def render_entry(self, e: FileEntry) -> QListWidgetItem:
        kind_key = {
            "PDF": "kind_pdf",
            "WORD": "kind_word",
            "EXCEL": "kind_excel",
            "POWERPOINT": "kind_powerpoint",
            "UNKNOWN": "kind_unknown",
        }.get(e.detected_kind, "kind_unknown")

        status_key = {
            "PENDING": "status_pending",
            "RUNNING": "status_running",
            "DONE": "status_done",
            "FAILED": "status_failed",
            "UNSUPPORTED": "status_unsupported",
        }.get(e.status, "status_failed")

        msg = f" - {e.message}" if e.message else ""
        text = f"[{self.i18n.t(status_key)}] [{self.i18n.t(kind_key)}] {str(e.path)}{msg}"
        item = QListWidgetItem(text)

        if e.status == "DONE":
            item.setForeground(Qt.darkGreen)
        elif e.status in ("FAILED", "UNSUPPORTED"):
            item.setForeground(Qt.darkRed)
        elif e.status == "RUNNING":
            item.setForeground(Qt.darkBlue)
        return item

    def on_toggle_pw(self, state: int):
        self.pw_input.setEchoMode(QLineEdit.Normal if state == Qt.Checked else QLineEdit.Password)

    def on_overwrite_toggled(self, state: int):
        if state == Qt.Checked:
            resp = QMessageBox.warning(
                self,
                self.i18n.t("overwrite_warn_title"),
                self.i18n.t("overwrite_warn_body"),
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if resp != QMessageBox.Yes:
                self.overwrite_cb.blockSignals(True)
                self.overwrite_cb.setChecked(False)
                self.overwrite_cb.blockSignals(False)

    # -------- Parallel processing --------
    def on_protect(self):
        if self._running:
            return

        password = self.pw_input.text()
        if not password:
            QMessageBox.warning(self, self.i18n.t("msg_done_title"), self.i18n.t("err_no_password"))
            return
        if not self.entries:
            QMessageBox.warning(self, self.i18n.t("msg_done_title"), self.i18n.t("err_no_files"))
            return

        overwrite = self.overwrite_cb.isChecked()

        pending_indexes = []
        for idx, e in enumerate(self.entries):
            if e.status == "UNSUPPORTED":
                continue
            if e.path.suffix.lower() not in ALL_SUPPORTED_EXT:
                e.status = "UNSUPPORTED"
                e.message = "Unsupported extension"
                continue
            e.status = "PENDING"
            e.message = ""
            pending_indexes.append(idx)

        if not pending_indexes:
            self.refresh_list_items()
            QMessageBox.information(self, self.i18n.t("msg_done_title"), self.i18n.t("err_no_files"))
            return

        self._running = True
        self.refresh_ui_texts()
        self.protect_btn.setEnabled(False)
        self.lang_combo.setEnabled(False)
        self.overwrite_cb.setEnabled(False)

        self._total = len(pending_indexes)
        self._done = 0
        self.progress_bar.setMaximum(self._total)
        self.progress_bar.setValue(0)

        done_counter = [0]

        for idx in pending_indexes:
            worker = ProtectWorker(
                index=idx,
                entry=self.entries[idx],
                password=password,
                overwrite=overwrite,
                total=self._total,
                done_counter=done_counter,
                signals=self.signals,
            )
            self.pool.start(worker)

        self.refresh_list_items()

    def on_worker_started(self, index: int):
        if 0 <= index < len(self.entries):
            self.entries[index].status = "RUNNING"
            self.entries[index].message = ""
            self.refresh_list_items()

    def on_worker_finished(self, index: int, status: str, message: str):
        if 0 <= index < len(self.entries):
            e = self.entries[index]
            e.status = status

            # Localize known messages
            if message == "Magic/structure mismatch with extension":
                e.message = self.i18n.t("err_magic_mismatch")
            elif message == "Not a valid ZIP container for OOXML":
                e.message = self.i18n.t("err_zip_invalid")
            elif message == "ZIP valid but not supported OOXML structure":
                e.message = self.i18n.t("err_ooxml_structure")
            elif message.startswith("SAVED::"):
                p = message.split("SAVED::", 1)[1]
                e.message = self.i18n.t("saved_as").format(path=p)
            elif message.startswith("OVERWRITTEN::"):
                p = message.split("OVERWRITTEN::", 1)[1]
                e.message = self.i18n.t("overwritten").format(path=p)
            else:
                e.message = message

            # Update detected kind based on real detection (best-effort)
            real_kind, _ = detect_real_type(e.path)
            if real_kind in {"PDF", "WORD", "EXCEL", "POWERPOINT"}:
                e.detected_kind = real_kind

            self.refresh_list_items()

    def on_worker_progress(self, done: int, total: int):
        self._done = done
        self.progress_bar.setValue(done)

        if done >= total:
            self._running = False
            self.refresh_ui_texts()
            self.protect_btn.setEnabled(True)
            self.lang_combo.setEnabled(True)
            self.overwrite_cb.setEnabled(True)

            done_count = sum(1 for e in self.entries if e.status == "DONE")
            failed_count = sum(1 for e in self.entries if e.status == "FAILED")
            unsupported_count = sum(1 for e in self.entries if e.status == "UNSUPPORTED")

            QMessageBox.information(
                self,
                self.i18n.t("msg_done_title"),
                self.i18n.t("msg_done_body").format(
                    done=done_count, failed=failed_count, unsupported=unsupported_count
                ),
            )


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
