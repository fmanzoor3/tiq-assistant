"""End-of-day "What did you do today?" dialog.

Flow: speak (or type) a summary of the day -> transcribe locally -> the local
LLM turns it into draft entries -> review them in an editable table -> confirm
to save. Meetings for the day are auto-pulled and shown alongside.

Voice is optional (faster-whisper + sounddevice); the text box always works.
The LLM call runs on a worker thread so the UI stays responsive.
"""

from datetime import date
from typing import Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QSpinBox,
    QGroupBox, QWidget, QAbstractItemView, QMessageBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject

from tiq_assistant.core.models import (
    TimesheetEntry, ActivityCode, EntryStatus, EntrySource,
)
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.services.entry_generation_service import (
    EntryGenerationService, load_llm_config,
)
from tiq_assistant.integrations.llm_client import LLMError
from tiq_assistant.desktop.icon import create_app_icon


class _GenWorker(QObject):
    """Runs the LLM generation off the UI thread."""
    done = pyqtSignal(object)   # GenerationResult
    failed = pyqtSignal(str)

    def __init__(self, transcript, target_date, remaining, context=None):
        super().__init__()
        self._transcript = transcript
        self._target_date = target_date
        self._remaining = remaining
        self._context = context or {}

    def run(self):
        try:
            svc = EntryGenerationService()
            result = svc.generate(
                self._transcript, self._target_date, self._remaining,
                existing_entries=self._context.get("existing_entries"),
                meetings=self._context.get("meetings"),
                recent_context=self._context.get("recent_context"),
            )
            self.done.emit(result)
        except LLMError as e:
            self.failed.emit(str(e))
        except Exception as e:  # noqa: BLE001
            self.failed.emit(f"Unexpected error: {e}")


class _TranscribeWorker(QObject):
    """Runs speech-to-text (incl. first-run model download) off the UI thread."""
    done = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, wav_path, model_ref, vocabulary=None):
        super().__init__()
        self._wav = wav_path
        self._model_ref = model_ref
        self._vocabulary = vocabulary

    def run(self):
        try:
            from tiq_assistant.integrations import speech
            text = speech.transcribe(
                self._wav, model_ref=self._model_ref, vocabulary=self._vocabulary
            )
            self.done.emit(text)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(str(e))


class VoiceEntryDialog(QDialog):
    COLORS = {
        'primary': '#0078D4', 'success': '#107C10', 'danger': '#D13438',
        'warning': '#FFB900', 'gray': '#E1E1E1', 'gray_light': '#F5F5F5',
        'text': '#323130', 'text_secondary': '#605E5C',
    }
    MAX_ENTRY_HOURS = 12

    def __init__(self, target_date: date, remaining_hours: int, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._store = get_store()
        self._target_date = target_date
        self._remaining = max(1, remaining_hours)
        self._projects = self._store.get_projects()
        self._settings = self._store.get_settings()

        self._recorder = None
        self._recording = False
        self._thread = None
        self._worker = None
        self._tx_thread = None
        self._tx_worker = None
        self._saved_count = 0

        self._setup_ui()

    # --------------------------------------------------------------- UI

    def _setup_ui(self) -> None:
        self.setWindowTitle("What did you do today?")
        self.setWindowIcon(create_app_icon())
        self.setModal(True)
        self.setMinimumSize(820, 680)
        self.resize(1000, 780)
        self.setStyleSheet(f"""
            QDialog {{ background-color: white; color: {self.COLORS['text']}; }}
            QLabel {{ color: {self.COLORS['text']}; }}
            QGroupBox {{ font-weight: bold; border: 1px solid {self.COLORS['gray']};
                border-radius: 4px; margin-top: 12px; padding-top: 10px; }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 5px; }}
            QTextEdit, QComboBox, QSpinBox {{
                color: {self.COLORS['text']}; background-color: white;
                border: 1px solid {self.COLORS['gray']}; border-radius: 4px; padding: 6px;
            }}
            QComboBox QAbstractItemView {{ background: white; color: {self.COLORS['text']};
                selection-background-color: {self.COLORS['primary']}; selection-color: white; }}
            QTableWidget {{ border: 1px solid {self.COLORS['gray']};
                gridline-color: {self.COLORS['gray']}; background: white; color: {self.COLORS['text']}; }}
            QHeaderView::section {{ background: {self.COLORS['gray_light']}; padding: 6px;
                border: none; border-right: 1px solid {self.COLORS['gray']};
                border-bottom: 1px solid {self.COLORS['gray']}; font-weight: bold; }}
            QPushButton {{ padding: 8px 14px; border: 1px solid {self.COLORS['gray']};
                border-radius: 4px; background: white; color: {self.COLORS['text']}; }}
            QPushButton:hover {{ background: {self.COLORS['gray_light']}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(18, 18, 18, 18)

        title = QLabel(f"<h2>What did you do today?</h2>")
        layout.addWidget(title)
        sub = QLabel(
            f"{self._target_date.strftime('%A, %d %B %Y')} — "
            f"{self._remaining}h left to fill. "
            f"Speak or type what you worked on; I'll draft the entries."
        )
        sub.setStyleSheet(f"color: {self.COLORS['text_secondary']};")
        sub.setWordWrap(True)
        layout.addWidget(sub)

        # Voice controls
        voice_row = QHBoxLayout()
        self._record_btn = QPushButton("🎤 Start recording")
        self._record_btn.clicked.connect(self._toggle_record)
        voice_row.addWidget(self._record_btn)
        self._voice_status = QLabel("")
        self._voice_status.setStyleSheet(f"color: {self.COLORS['text_secondary']};")
        voice_row.addWidget(self._voice_status)
        voice_row.addStretch()
        layout.addLayout(voice_row)

        # Transcript box
        layout.addWidget(QLabel("Your summary (editable):"))
        self._transcript = QTextEdit()
        self._transcript.setPlaceholderText(
            "e.g. I worked on 3 tasks today, all under the default project. "
            "Agentbot feedback for a few hours, then optimization testing..."
        )
        self._transcript.setMinimumHeight(120)
        layout.addWidget(self._transcript)

        # Generate button
        gen_row = QHBoxLayout()
        self._generate_btn = QPushButton("✨ Generate entries")
        self._generate_btn.setStyleSheet(
            f"background: {self.COLORS['primary']}; color: white; border: none; "
            f"padding: 10px 18px; font-weight: bold;")
        self._generate_btn.clicked.connect(self._generate)
        gen_row.addWidget(self._generate_btn)
        self._gen_status = QLabel("")
        self._gen_status.setStyleSheet(f"color: {self.COLORS['text_secondary']};")
        gen_row.addWidget(self._gen_status)
        gen_row.addStretch()
        layout.addLayout(gen_row)

        # Proposed entries table
        group = QGroupBox("Proposed entries (edit before saving)")
        gl = QVBoxLayout(group)
        self._table = QTableWidget()
        self._table.setColumnCount(4)
        self._table.setHorizontalHeaderLabels(["Project", "Hours", "Description", ""])
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self._table.setColumnWidth(0, 230)   # Project: wide enough for full name
        self._table.setColumnWidth(1, 80)    # Hours
        self._table.setColumnWidth(3, 48)    # Remove button
        self._table.verticalHeader().setVisible(False)
        self._table.verticalHeader().setDefaultSectionSize(40)  # taller rows
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setWordWrap(True)
        gl.addWidget(self._table)

        add_row = QHBoxLayout()
        add_line_btn = QPushButton("+ Add line")
        add_line_btn.clicked.connect(lambda: self._add_row(None, 1, ""))
        add_row.addWidget(add_line_btn)
        add_row.addStretch()
        self._total_label = QLabel("")
        add_row.addWidget(self._total_label)
        gl.addLayout(add_row)
        layout.addWidget(group, 1)

        # Bottom buttons
        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        btns.addWidget(cancel)
        self._save_btn = QPushButton("Save entries")
        self._save_btn.setStyleSheet(
            f"background: {self.COLORS['success']}; color: white; border: none; "
            f"padding: 10px 20px; font-weight: bold;")
        self._save_btn.clicked.connect(self._save)
        btns.addWidget(self._save_btn)
        layout.addLayout(btns)

        self._update_total()

    # ------------------------------------------------------------ voice

    def _toggle_record(self) -> None:
        from tiq_assistant.integrations import speech

        if not self._recording:
            available, reason = speech.is_available()
            if not available:
                QMessageBox.information(
                    self, "Voice not available",
                    f"Voice input isn't available on this machine, so please type "
                    f"your summary instead.\n\nDetails: {reason}"
                )
                return
            try:
                self._recorder = speech.AudioRecorder()
                self._recorder.start()
            except Exception as e:  # noqa: BLE001
                QMessageBox.warning(self, "Microphone error", f"Could not start recording: {e}")
                return
            self._recording = True
            self._record_btn.setText("⏹ Stop & transcribe")
            self._voice_status.setText("Recording… speak now.")
        else:
            self._recording = False
            self._record_btn.setText("🎤 Start recording")
            self._record_btn.setEnabled(False)
            try:
                wav = self._recorder.stop()
            except Exception as e:  # noqa: BLE001
                self._record_btn.setEnabled(True)
                QMessageBox.warning(self, "Recording error", str(e))
                return
            if wav is None:
                self._voice_status.setText("No audio captured.")
                self._record_btn.setEnabled(True)
                return

            # Transcribe on a worker thread so the UI stays responsive -- the
            # first run downloads the model, which can take a while.
            self._voice_status.setText(
                "Transcribing… (first run downloads the speech model — please wait)"
            )
            cfg = load_llm_config(self._store)
            self._tx_thread = QThread()
            self._tx_worker = _TranscribeWorker(
                wav, cfg.whisper_model or "base", vocabulary=self._build_vocabulary()
            )
            self._tx_worker.moveToThread(self._tx_thread)
            self._tx_thread.started.connect(self._tx_worker.run)
            self._tx_worker.done.connect(self._on_transcribed)
            self._tx_worker.failed.connect(self._on_transcribe_failed)
            self._tx_worker.done.connect(self._tx_thread.quit)
            self._tx_worker.failed.connect(self._tx_thread.quit)
            self._tx_thread.start()

    def _build_vocabulary(self) -> str:
        """Domain terms (custom + project names + keywords) to bias transcription."""
        terms: list[str] = []
        # User-supplied custom terms first.
        try:
            custom = load_llm_config(self._store).custom_terms
            terms.extend(t.strip() for t in custom.split(",") if t.strip())
        except Exception:
            pass
        for p in self._projects:
            terms.append(p.name)
            terms.extend(p.keywords or [])
        # De-dup while preserving order; keep it reasonably short.
        seen = set()
        uniq = []
        for t in terms:
            t = (t or "").strip()
            if t and t.lower() not in seen:
                seen.add(t.lower())
                uniq.append(t)
        return ", ".join(uniq[:40])

    def _on_transcribed(self, text: str) -> None:
        # Append (don't clobber) so multiple takes accumulate.
        existing = self._transcript.toPlainText().strip()
        self._transcript.setPlainText((existing + " " + text).strip() if existing else text)
        self._voice_status.setText("Transcribed. Review/edit, then Generate.")
        self._record_btn.setEnabled(True)

    def _on_transcribe_failed(self, message: str) -> None:
        self._voice_status.setText("")
        self._record_btn.setEnabled(True)
        QMessageBox.warning(self, "Transcription error", message)

    # --------------------------------------------------------- generate

    def _generate(self) -> None:
        transcript = self._transcript.toPlainText().strip()
        if not transcript:
            QMessageBox.information(self, "Nothing to generate",
                                    "Please speak or type what you did today first.")
            return

        cfg = load_llm_config(self._store)
        if not cfg.enabled:
            QMessageBox.information(
                self, "AI assistant disabled",
                "Enable the AI assistant in Settings (and set the endpoint) to "
                "generate entries from your summary."
            )
            return

        self._generate_btn.setEnabled(False)
        self._gen_status.setText("Thinking… contacting the local LLM.")

        # Gather day context (existing entries, meetings, recent days) so the
        # LLM fills only the gap and matches ongoing phrasing.
        try:
            context = EntryGenerationService(store=self._store).gather_day_context(
                self._target_date
            )
        except Exception:
            context = {}

        # Run generation off the UI thread.
        self._thread = QThread()
        self._worker = _GenWorker(transcript, self._target_date, self._remaining, context)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.done.connect(self._on_generated)
        self._worker.failed.connect(self._on_gen_failed)
        self._worker.done.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.start()

    def _on_generated(self, result) -> None:
        self._generate_btn.setEnabled(True)
        self._table.setRowCount(0)
        for e in result.entries:
            self._add_row(e.project_id, e.hours, e.description)
        self._gen_status.setText(f"Drafted {len(result.entries)} entries. Review and save.")
        self._update_total()

    def _on_gen_failed(self, message: str) -> None:
        self._generate_btn.setEnabled(True)
        self._gen_status.setText("")
        QMessageBox.warning(self, "Could not generate entries", message)

    # ------------------------------------------------------------- table

    def _add_row(self, project_id, hours: int, description: str) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)

        project_combo = QComboBox()
        project_combo.addItem("-- None --", None)
        sel = 0
        for i, p in enumerate(self._projects):
            project_combo.addItem(p.name, p.id)
            if project_id and p.id == project_id:
                sel = i + 1
        project_combo.setCurrentIndex(sel)
        self._table.setCellWidget(row, 0, project_combo)

        hours_spin = QSpinBox()
        hours_spin.setRange(1, self.MAX_ENTRY_HOURS)
        hours_spin.setValue(max(1, min(hours, self.MAX_ENTRY_HOURS)))
        hours_spin.valueChanged.connect(self._update_total)
        self._table.setCellWidget(row, 1, hours_spin)

        self._table.setItem(row, 2, QTableWidgetItem(description))

        remove = QPushButton("✕")
        remove.clicked.connect(lambda _c, w=remove: self._remove_row(w))
        self._table.setCellWidget(row, 3, remove)

        self._update_total()

    def _remove_row(self, button) -> None:
        for row in range(self._table.rowCount()):
            if self._table.cellWidget(row, 3) is button:
                self._table.removeRow(row)
                break
        self._update_total()

    def _update_total(self) -> None:
        total = sum(
            self._table.cellWidget(r, 1).value()
            for r in range(self._table.rowCount())
            if self._table.cellWidget(r, 1)
        )
        color = self.COLORS['success'] if total == self._remaining else self.COLORS['text_secondary']
        self._total_label.setText(
            f"<span style='color:{color}; font-weight:bold;'>Total {total} / {self._remaining}h</span>"
        )

    # -------------------------------------------------------------- save

    def _save(self) -> None:
        if self._table.rowCount() == 0:
            QMessageBox.information(self, "Nothing to save", "There are no entries to save.")
            return

        saved = 0
        for row in range(self._table.rowCount()):
            project_combo = self._table.cellWidget(row, 0)
            hours_spin = self._table.cellWidget(row, 1)
            desc_item = self._table.item(row, 2)

            description = (desc_item.text().strip() if desc_item else "")
            if not description:
                QMessageBox.warning(self, "Description required",
                                    f"Line {row + 1} needs a description.")
                return

            project_id = project_combo.currentData()
            project = self._store.get_project(project_id) if project_id else None

            entry = TimesheetEntry(
                consultant_id=self._settings.consultant_id,
                entry_date=self._target_date,
                hours=hours_spin.value(),
                ticket_number=project.ticket_number if project else None,
                project_name=project.name if project else None,
                activity_code=self._settings.default_activity_code,
                location=self._settings.default_location,
                description=description,
                status=EntryStatus.DRAFT,
                source=EntrySource.MANUAL,
            )
            self._store.save_entry(entry)
            if project:
                self._store.update_recent_project(project)
            saved += 1

        self._saved_count = saved
        QMessageBox.information(self, "Saved", f"Saved {saved} entries.")
        self.accept()

    def get_saved_count(self) -> int:
        return self._saved_count
