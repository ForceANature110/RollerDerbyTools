import json
import re
import sys
import time
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import QObject, QUrl, Slot, Qt
from PySide6.QtGui import QAction, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QInputDialog,
    QPlainTextEdit,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineWidgets import QWebEngineView


PLAYBACK_SPEED_STEPS = [0.25, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0]
DEFAULT_FRAME_STEP_SECONDS = 1.0 / 60.0
TEXT_INPUT_FIELDS = [
    "home_jammer",
    "away_jammer",
    "home_pivot",
    "home_blocker_1",
    "home_blocker_2",
    "home_blocker_3",
    "away_pivot",
    "away_blocker_1",
    "away_blocker_2",
    "away_blocker_3",
    "home_score_start",
    "away_score_start",
    "home_passes",
    "away_passes",
    "home_score_end",
    "away_score_end",
    "home_star_pass",
    "away_star_pass",
    "lead_jammer",
    "notes",
]
PENALTY_CODES = {
    "A": "High Block",
    "B": "Back Block",
    "C": "Illegal Contact",
    "D": "Direction",
    "E": "Leg Block",
    "F": "Forearms",
    "G": "Misconduct",
    "H": "Blocking with the Head",
    "I": "Illegal Procedure",
    "L": "Low Block",
    "M": "Multiplayer Block",
    "N": "Interference",
    "P": "Illegal Position",
    "X": "Cut",
}


YOUTUBE_HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    html, body { margin: 0; padding: 0; background: #191970; overflow: hidden; height: 100%; }
    #player { position: absolute; inset: 0; }
  </style>
  <script src="https://www.youtube.com/iframe_api"></script>
  <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
</head>
<body>
  <div id="player"></div>
  <script>
    let player = null;
    let backend = null;
    let pendingVideoId = null;

    new QWebChannel(qt.webChannelTransport, function(channel) {
      backend = channel.objects.backend;
      if (pendingVideoId) {
        loadVideo(pendingVideoId);
      }
      setInterval(function() {
        if (player && typeof player.getCurrentTime === 'function' && backend) {
          backend.updateCurrentTime(player.getCurrentTime(), player.getPlayerState());
        }
      }, 150);
    });

    function onYouTubeIframeAPIReady() {
      player = new YT.Player('player', {
        width: '100%',
        height: '100%',
        videoId: '',
        playerVars: {
          playsinline: 1,
          rel: 0,
          modestbranding: 1,
          origin: 'https://fans-annotator.local',
        },
        events: {
          'onReady': function() {
            if (pendingVideoId) {
              loadVideo(pendingVideoId);
            }
          },
          'onStateChange': function(event) {
            if (backend && player) {
              backend.updateCurrentTime(player.getCurrentTime(), event.data);
            }
          }
        }
      });
    }

    function loadVideo(videoId) {
      pendingVideoId = videoId;
      if (player && typeof player.loadVideoById === 'function') {
        player.loadVideoById(videoId, 0);
      }
    }

    function playVideo() {
      if (player) player.playVideo();
    }

    function pauseVideo() {
      if (player) player.pauseVideo();
    }

    function seekTo(seconds) {
      if (player) player.seekTo(seconds, true);
    }

    function setPlaybackRate(rate) {
      if (player && typeof player.setPlaybackRate === 'function') {
        try { player.setPlaybackRate(rate); } catch (e) {}
      }
    }
  </script>
</body>
</html>
"""


@dataclass
class JamRecord:
    period_number: int = 1
    jam_number: int = 1
    start_time: Optional[float] = None
    end_time: Optional[float] = None
    home_jammer: str = ""
    away_jammer: str = ""
    home_lineup: List[str] = field(default_factory=list)
    away_lineup: List[str] = field(default_factory=list)
    home_score_start: Optional[int] = None
    away_score_start: Optional[int] = None
    home_passes: List[int] = field(default_factory=list)
    away_passes: List[int] = field(default_factory=list)
    home_score_end: Optional[int] = None
    away_score_end: Optional[int] = None
    home_star_pass: bool = False
    away_star_pass: bool = False
    lead_jammer: str = "unknown"
    penalties: List[dict] = field(default_factory=list)
    notes: str = ""


@dataclass
class AnnotationFile:
    source: str
    jams: List[JamRecord] = field(default_factory=list)


class BackendBridge(QObject):
    def __init__(self, window: "YouTubeAnnotatorWindow"):
        super().__init__()
        self.window = window

    @Slot(float, int)
    def updateCurrentTime(self, seconds: float, player_state: int) -> None:
        self.window.on_player_time_update(seconds, player_state)


class YouTubeAnnotatorWindow(QMainWindow):
    LINEUP_FIELD_MAP = {
        "home_pivot": ("home_lineup", 0),
        "home_blocker_1": ("home_lineup", 1),
        "home_blocker_2": ("home_lineup", 2),
        "home_blocker_3": ("home_lineup", 3),
        "away_pivot": ("away_lineup", 0),
        "away_blocker_1": ("away_lineup", 1),
        "away_blocker_2": ("away_lineup", 2),
        "away_blocker_3": ("away_lineup", 3),
    }

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("FaN's Derby Annotation Viewer")
        self.resize(1600, 950)

        self.youtube_url = ""
        self.youtube_id = ""
        self.current_time = 0.0
        self.player_state = -1
        self.playback_rate = 1.0
        self.last_seek_target: Optional[float] = None
        self.ignore_player_updates_until = 0.0
        self.current_jam_index = 0
        self.data = AnnotationFile(source="", jams=[JamRecord(period_number=1, jam_number=1)])
        self.output_path: Optional[Path] = None
        self.help_visible = False
        self.edit_field_index = 0

        self.edit_field_widgets = {}
        self.edit_prev_labels = {}

        self._build_ui()
        self._build_shortcuts()
        self.refresh_ui()
        self.refresh_edit_panel()
        self.refresh_penalty_panel()
        self.refresh_help_panel()
        self.annotation_file_label.setText("No annotation file loaded")
        self.source_status_label.setText("Open an annotation JSON. The video source link is stored inside it.")

    # ---------- UI ----------
    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QGridLayout(root)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setHorizontalSpacing(10)
        root_layout.setVerticalSpacing(10)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)

        source_box = QGroupBox("Annotation File and Video Source")
        source_layout = QVBoxLayout(source_box)

        file_row = QWidget()
        file_row_layout = QHBoxLayout(file_row)
        file_row_layout.setContentsMargins(0, 0, 0, 0)
        self.annotation_file_label = QLabel("No annotation file loaded")
        self.open_annotation_button = QPushButton("Open Annotation JSON")
        self.open_annotation_button.clicked.connect(self.open_annotation_file)
        self.save_as_button_top = QPushButton("Save As JSON")
        self.save_as_button_top.clicked.connect(self.save_annotations_as)
        file_row_layout.addWidget(self.annotation_file_label, stretch=1)
        file_row_layout.addWidget(self.open_annotation_button)
        file_row_layout.addWidget(self.save_as_button_top)

        link_row = QWidget()
        link_row_layout = QHBoxLayout(link_row)
        link_row_layout.setContentsMargins(0, 0, 0, 0)
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("Video source link stored in this annotation file")
        self.link_apply_button = QPushButton("Apply / Load Video")
        self.link_apply_button.clicked.connect(self.apply_source_link)
        link_row_layout.addWidget(self.url_edit, stretch=1)
        link_row_layout.addWidget(self.link_apply_button)

        self.source_status_label = QLabel("Open an annotation JSON. The video source link is stored inside it.")

        source_layout.addWidget(file_row)
        source_layout.addWidget(link_row)
        source_layout.addWidget(self.source_status_label)
        left_layout.addWidget(source_box)

        info_box = QGroupBox("Current Jam")
        info_layout = QFormLayout(info_box)
        self.period_label = QLabel("1")
        self.jam_label = QLabel("1")
        self.time_label = QLabel("0.000s")
        self.start_label = QLabel("-")
        self.end_label = QLabel("-")
        self.home_jammer_label = QLabel("-")
        self.away_jammer_label = QLabel("-")
        self.home_lineup_label = QLabel("-")
        self.away_lineup_label = QLabel("-")
        self.score_start_label = QLabel("- / -")
        self.home_pass_label = QLabel("-")
        self.away_pass_label = QLabel("-")
        self.score_end_label = QLabel("- / -")
        self.star_label = QLabel("N / N")
        self.lead_label = QLabel("unknown")
        self.penalties_label = QLabel("0")
        info_layout.addRow("Period", self.period_label)
        info_layout.addRow("Jam", self.jam_label)
        info_layout.addRow("Current time", self.time_label)
        info_layout.addRow("Jam start", self.start_label)
        info_layout.addRow("Jam end", self.end_label)
        info_layout.addRow("Home jammer", self.home_jammer_label)
        info_layout.addRow("Away jammer", self.away_jammer_label)
        info_layout.addRow("Home lineup", self.home_lineup_label)
        info_layout.addRow("Away lineup", self.away_lineup_label)
        info_layout.addRow("Score start H/A", self.score_start_label)
        info_layout.addRow("Home passes", self.home_pass_label)
        info_layout.addRow("Away passes", self.away_pass_label)
        info_layout.addRow("Score end H/A", self.score_end_label)
        info_layout.addRow("Star pass H/A", self.star_label)
        info_layout.addRow("Lead", self.lead_label)
        info_layout.addRow("Penalties", self.penalties_label)
        left_layout.addWidget(info_box)

        control_box = QGroupBox("Controls")
        control_layout = QGridLayout(control_box)
        buttons = [
            ("Play / Pause [Space]", self.toggle_play_pause, 0, 0),
            ("Mark Start [S]", self.mark_start, 0, 1),
            ("Mark End [E]", self.mark_end, 0, 2),
            ("Prev Jam [P]", self.previous_jam, 1, 0),
            ("Next Jam [N]", self.next_jam, 1, 1),
            ("Next Period [5]", self.start_next_period, 1, 2),
            ("Edit [M]", self.toggle_edit_panel, 2, 0),
            ("Penalty [B]", self.toggle_penalty_panel, 2, 1),
            ("Help [H]", self.toggle_help_panel, 2, 2),
            ("Save [Ctrl+S]", self.save_annotations, 3, 0),
            ("Jam Start [\]", self.jump_to_current_jam_start, 3, 1),
        ]
        for text, callback, row, col in buttons:
            button = QPushButton(text)
            button.clicked.connect(callback)
            control_layout.addWidget(button, row, col)
        left_layout.addWidget(control_box)

        notes_box = QGroupBox("Notes")
        notes_layout = QVBoxLayout(notes_box)
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.textChanged.connect(self.notes_changed)
        notes_layout.addWidget(self.notes_edit)
        left_layout.addWidget(notes_box, stretch=1)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        self.player_view = QWebEngineView()
        self.player_view.setHtml(YOUTUBE_HTML, QUrl("https://fans-annotator.local/"))
        self.bridge = BackendBridge(self)
        self.channel = QWebChannel(self.player_view.page())
        self.channel.registerObject("backend", self.bridge)
        self.player_view.page().setWebChannel(self.channel)
        right_layout.addWidget(self.player_view, stretch=4)

        lower_panel = QWidget()
        lower_layout = QHBoxLayout(lower_panel)
        lower_layout.setContentsMargins(0, 0, 0, 0)
        lower_layout.setSpacing(10)

        self.edit_box = QGroupBox("Edit Jam [M]")
        edit_box_layout = QVBoxLayout(self.edit_box)
        self.edit_scroll = QScrollArea()
        self.edit_scroll.setWidgetResizable(True)
        edit_container = QWidget()
        self.edit_form = QFormLayout(edit_container)
        self.edit_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self._build_edit_form()
        self.edit_scroll.setWidget(edit_container)
        edit_box_layout.addWidget(self.edit_scroll)
        lower_layout.addWidget(self.edit_box, stretch=3)

        side_stack = QWidget()
        side_stack_layout = QVBoxLayout(side_stack)
        side_stack_layout.setContentsMargins(0, 0, 0, 0)
        side_stack_layout.setSpacing(10)

        self.penalty_box = QGroupBox("Penalty [B]")
        penalty_layout = QFormLayout(self.penalty_box)
        self.penalty_skater_edit = QLineEdit()
        self.penalty_code_combo = QComboBox()
        self.penalty_code_combo.addItem("")
        for code, desc in PENALTY_CODES.items():
            self.penalty_code_combo.addItem(f"{code} – {desc}", code)
        self.penalty_team_combo = QComboBox()
        self.penalty_team_combo.addItems(["home", "away"])
        self.penalty_save_button = QPushButton("Save penalty [Enter]")
        self.penalty_save_button.clicked.connect(self.save_penalty)
        self.penalty_cancel_button = QPushButton("Cancel [Esc]")
        self.penalty_cancel_button.clicked.connect(self.cancel_penalty)
        penalty_buttons = QWidget()
        penalty_buttons_layout = QHBoxLayout(penalty_buttons)
        penalty_buttons_layout.setContentsMargins(0, 0, 0, 0)
        penalty_buttons_layout.addWidget(self.penalty_save_button)
        penalty_buttons_layout.addWidget(self.penalty_cancel_button)
        penalty_layout.addRow("Skater", self.penalty_skater_edit)
        penalty_layout.addRow("Code", self.penalty_code_combo)
        penalty_layout.addRow("Team", self.penalty_team_combo)
        penalty_layout.addRow(penalty_buttons)
        side_stack_layout.addWidget(self.penalty_box)

        self.help_box = QGroupBox("Help [H]")
        help_layout = QVBoxLayout(self.help_box)
        self.help_text = QPlainTextEdit()
        self.help_text.setReadOnly(True)
        help_layout.addWidget(self.help_text)
        side_stack_layout.addWidget(self.help_box, stretch=1)

        lower_layout.addWidget(side_stack, stretch=2)
        right_layout.addWidget(lower_panel, stretch=3)

        root_layout.addWidget(left_panel, 0, 0)
        root_layout.addWidget(right_panel, 0, 1)
        root_layout.setColumnStretch(0, 1)
        root_layout.setColumnStretch(1, 3)

        self.setCentralWidget(root)

        menu = self.menuBar().addMenu("File")
        open_action = QAction("Open Annotation File…", self)
        open_action.triggered.connect(self.open_annotation_file)
        save_as_action = QAction("Save As…", self)
        save_as_action.triggered.connect(self.save_annotations_as)
        menu.addAction(open_action)
        menu.addAction(save_as_action)

        self.edit_box.hide()
        self.penalty_box.hide()
        self.help_box.hide()

    def _build_edit_form(self) -> None:
        field_labels = {
            "home_jammer": "Home jammer",
            "away_jammer": "Away jammer",
            "home_pivot": "Home pivot",
            "home_blocker_1": "Home blocker 1",
            "home_blocker_2": "Home blocker 2",
            "home_blocker_3": "Home blocker 3",
            "away_pivot": "Away pivot",
            "away_blocker_1": "Away blocker 1",
            "away_blocker_2": "Away blocker 2",
            "away_blocker_3": "Away blocker 3",
            "home_score_start": "Home score start",
            "away_score_start": "Away score start",
            "home_passes": "Home passes",
            "away_passes": "Away passes",
            "home_score_end": "Home score end",
            "away_score_end": "Away score end",
            "home_star_pass": "Home star pass",
            "away_star_pass": "Away star pass",
            "lead_jammer": "Lead jammer",
            "notes": "Notes",
        }

        for field_name in TEXT_INPUT_FIELDS:
            row = QWidget()
            row_layout = QVBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(2)

            if field_name in {"home_star_pass", "away_star_pass"}:
                widget = QCheckBox()
                widget.stateChanged.connect(lambda _=None, f=field_name: self.on_edit_field_changed(f))
            elif field_name == "lead_jammer":
                widget = QComboBox()
                widget.addItems(["unknown", "home", "away", "none"])
                widget.currentTextChanged.connect(lambda _=None, f=field_name: self.on_edit_field_changed(f))
            elif field_name == "notes":
                widget = QPlainTextEdit()
                widget.setFixedHeight(90)
                widget.textChanged.connect(lambda f=field_name: self.on_edit_field_changed(f))
            else:
                widget = QLineEdit()
                widget.textChanged.connect(lambda _=None, f=field_name: self.on_edit_field_changed(f))

            prev_label = QLabel("Previously used: -")
            prev_label.setStyleSheet("color: #9ec7ff; font-size: 11px;")
            row_layout.addWidget(widget)
            row_layout.addWidget(prev_label)
            self.edit_form.addRow(field_labels[field_name], row)
            self.edit_field_widgets[field_name] = widget
            self.edit_prev_labels[field_name] = prev_label

    def _build_shortcuts(self) -> None:
        shortcuts = {
            "Space": self.toggle_play_pause,
            "S": self.mark_start,
            "E": self.mark_end,
            "N": self.next_jam,
            "P": self.previous_jam,
            "5": self.start_next_period,
            "Ctrl+S": self.save_annotations,
            "\\": self.jump_to_current_jam_start,
            "J": lambda: self.seek_relative(-5.0),
            "L": lambda: self.seek_relative(5.0),
            "A": lambda: self.seek_relative(-1.0),
            "F": lambda: self.seek_relative(1.0),
            "BracketLeft": self.jump_to_previous_saved_jam_start,
            "BracketRight": self.jump_to_next_saved_jam_start,
            "BraceLeft": self.jump_to_previous_saved_jam_end,
            "BraceRight": self.jump_to_next_saved_jam_end,
            "Comma": lambda: self.step_frame(-1),
            "Period": lambda: self.step_frame(1),
            "Minus": lambda: self.adjust_playback_speed(-1),
            "Equal": lambda: self.adjust_playback_speed(1),
            "0": self.reset_playback_speed,
            "H": self.toggle_help_panel,
            "M": self.toggle_edit_panel,
            "B": self.toggle_penalty_panel,
            "1": lambda: self.set_lead("home"),
            "2": lambda: self.set_lead("away"),
            "3": lambda: self.set_lead("none"),
            "4": lambda: self.set_lead("unknown"),
            "Z": lambda: self.undo_last_pass("home"),
            "X": lambda: self.undo_last_pass("away"),
            "D": self.delete_current_jam_if_empty,
            "Escape": self.handle_escape,
            "W": self.handle_w_key,
            "Tab": lambda: self.move_edit_field(1),
            "Return": lambda: self.move_edit_field(1),
            "Enter": lambda: self.move_edit_field(1),
            "Up": lambda: self.move_edit_field(-1),
            "Down": lambda: self.move_edit_field(1),
            "QuoteLeft": self.toggle_current_star_pass,
        }
        self._shortcuts = []
        for seq, callback in shortcuts.items():
            shortcut = QShortcut(QKeySequence(seq), self)
            shortcut.setContext(Qt.ApplicationShortcut)
            shortcut.activated.connect(callback)
            self._shortcuts.append(shortcut)

        pass_shortcuts = {
            "Shift+Q": ("home", 0),
            "Shift+W": ("home", 1),
            "Shift+E": ("home", 2),
            "Shift+R": ("home", 3),
            "Shift+T": ("home", 4),
            "Shift+A": ("away", 0),
            "Shift+S": ("away", 1),
            "Shift+D": ("away", 2),
            "Shift+F": ("away", 3),
            "Shift+G": ("away", 4),
        }
        for seq, (team, pts) in pass_shortcuts.items():
            shortcut = QShortcut(QKeySequence(seq), self)
            shortcut.setContext(Qt.ApplicationShortcut)
            shortcut.activated.connect(lambda t=team, p=pts: self.append_pass(t, p))
            self._shortcuts.append(shortcut)

    def _current_edit_field_name(self) -> str:
        return TEXT_INPUT_FIELDS[self.edit_field_index]

    def _focus_edit_field(self, index: int, ensure_visible: bool = True) -> None:
        self.edit_field_index = max(0, min(len(TEXT_INPUT_FIELDS) - 1, index))
        field_name = self._current_edit_field_name()
        widget = self.edit_field_widgets[field_name]
        widget.setFocus()
        if ensure_visible:
            self.edit_scroll.ensureWidgetVisible(widget)

    def move_edit_field(self, delta: int) -> None:
        if not self.edit_box.isVisible():
            return
        self._focus_edit_field((self.edit_field_index + delta) % len(TEXT_INPUT_FIELDS))

    def toggle_current_star_pass(self) -> None:
        if not self.edit_box.isVisible():
            return
        field_name = self._current_edit_field_name()
        if field_name in {"home_star_pass", "away_star_pass"}:
            widget = self.edit_field_widgets[field_name]
            if isinstance(widget, QCheckBox):
                widget.setChecked(not widget.isChecked())

    def handle_w_key(self) -> None:
        if self.edit_box.isVisible():
            self.save_annotations()
            self.edit_box.hide()
            self.player_view.setFocus()
        else:
            self.save_annotations()

    def handle_escape(self) -> None:
        if self.penalty_box.isVisible():
            self.cancel_penalty()
            self.player_view.setFocus()
        elif self.help_box.isVisible():
            self.help_box.hide()
            self.player_view.setFocus()
        elif self.edit_box.isVisible():
            self.edit_box.hide()
            self.player_view.setFocus()

    # ---------- Data helpers ----------
    def current_jam(self) -> JamRecord:
        while self.current_jam_index >= len(self.data.jams):
            self.data.jams.append(
                JamRecord(
                    period_number=self._default_period_for_new_jam(),
                    jam_number=self._default_jam_number_for_new_jam(),
                )
            )
        jam = self.data.jams[self.current_jam_index]
        self._apply_score_defaults(jam)
        return jam

    def _default_period_for_new_jam(self) -> int:
        if not self.data.jams:
            return 1
        return self.data.jams[-1].period_number

    def _default_jam_number_for_new_jam(self) -> int:
        if not self.data.jams:
            return 1
        return self.data.jams[-1].jam_number + 1

    def _apply_score_defaults(self, jam: JamRecord) -> None:
        try:
            jam_index = self.data.jams.index(jam)
        except ValueError:
            return
        if jam_index <= 0:
            self._recalculate_jam_scores(jam)
            return
        prev_jam = self.data.jams[jam_index - 1]
        if jam.home_score_start is None and prev_jam.home_score_end is not None:
            jam.home_score_start = prev_jam.home_score_end
        if jam.away_score_start is None and prev_jam.away_score_end is not None:
            jam.away_score_start = prev_jam.away_score_end
        self._recalculate_jam_scores(jam)

    def _recalculate_jam_scores(self, jam: JamRecord) -> None:
        jam.home_score_end = None if jam.home_score_start is None else jam.home_score_start + sum(jam.home_passes)
        jam.away_score_end = None if jam.away_score_start is None else jam.away_score_start + sum(jam.away_passes)

    def _propagate_scores_from_jam(self, changed_jam_index: int) -> None:
        if changed_jam_index < 0 or changed_jam_index >= len(self.data.jams):
            return
        self._apply_score_defaults(self.data.jams[changed_jam_index])
        for idx in range(changed_jam_index + 1, len(self.data.jams)):
            prev_jam = self.data.jams[idx - 1]
            jam = self.data.jams[idx]
            jam.home_score_start = prev_jam.home_score_end
            jam.away_score_start = prev_jam.away_score_end
            self._recalculate_jam_scores(jam)

    def _passes_to_text(self, passes: List[int]) -> str:
        return "-" if not passes else " + ".join(str(p) for p in passes)

    def _lineup_value_at(self, lineup: List[str], index: int) -> str:
        return lineup[index] if index < len(lineup) else ""

    def _set_lineup_value_at(self, lineup: List[str], index: int, value: str) -> List[str]:
        while len(lineup) <= index:
            lineup.append("")
        lineup[index] = value.strip()
        while lineup and lineup[-1] == "":
            lineup.pop()
        return lineup

    def _get_position_for_skater_in_jam(self, jam: JamRecord, team: str, skater: str):
        skater = skater.strip()
        if not skater:
            return None
        jammer_value = jam.home_jammer if team == "home" else jam.away_jammer
        lineup = jam.home_lineup if team == "home" else jam.away_lineup
        if jammer_value and jammer_value.strip() == skater:
            return ("jammer", None)
        if len(lineup) > 0 and lineup[0].strip() == skater:
            return ("pivot", 0)
        for idx in range(1, len(lineup)):
            if lineup[idx].strip() == skater:
                return ("blocker", idx)
        return None

    def _prefill_next_jam_from_recent_penalties(self, previous_index: int, next_index: int) -> None:
        if previous_index < 0 or next_index < 0:
            return
        if previous_index >= len(self.data.jams) or next_index >= len(self.data.jams):
            return
        previous_jam = self.data.jams[previous_index]
        next_jam = self.data.jams[next_index]
        if previous_jam.end_time is None:
            return
        recent_cutoff = previous_jam.end_time - 30.0
        for penalty in previous_jam.penalties:
            try:
                penalty_time = float(penalty.get("time", -1))
            except Exception:
                continue
            if penalty_time < recent_cutoff:
                continue
            team = str(penalty.get("team", "")).strip().lower()
            skater = str(penalty.get("skater", "")).strip()
            if team not in {"home", "away"} or not skater:
                continue
            position_info = self._get_position_for_skater_in_jam(previous_jam, team, skater)
            if position_info is None:
                continue
            role, slot_index = position_info
            if role == "jammer":
                if team == "home" and not next_jam.home_jammer:
                    next_jam.home_jammer = skater
                elif team == "away" and not next_jam.away_jammer:
                    next_jam.away_jammer = skater
            elif role == "pivot":
                lineup_attr = "home_lineup" if team == "home" else "away_lineup"
                lineup = list(getattr(next_jam, lineup_attr))
                if self._lineup_value_at(lineup, 0) == "":
                    setattr(next_jam, lineup_attr, self._set_lineup_value_at(lineup, 0, skater))
            elif role == "blocker" and slot_index is not None:
                lineup_attr = "home_lineup" if team == "home" else "away_lineup"
                lineup = list(getattr(next_jam, lineup_attr))
                if self._lineup_value_at(lineup, slot_index) == "":
                    setattr(next_jam, lineup_attr, self._set_lineup_value_at(lineup, slot_index, skater))

    def _collect_previous_values(self, field_name: str) -> List[str]:
        values: List[str] = []
        seen = set()
        for idx, jam in enumerate(self.data.jams):
            if idx == self.current_jam_index:
                continue
            value = ""
            if field_name in self.LINEUP_FIELD_MAP:
                lineup_attr, lineup_index = self.LINEUP_FIELD_MAP[field_name]
                lineup = getattr(jam, lineup_attr, [])
                value = self._lineup_value_at(lineup, lineup_index)
            else:
                raw = getattr(jam, field_name, "")
                if isinstance(raw, bool):
                    value = "Y" if raw else "N"
                elif raw is None:
                    value = ""
                elif isinstance(raw, list):
                    value = ", ".join(str(v) for v in raw)
                else:
                    value = str(raw).strip()
            value = value.strip()
            if value and value not in seen:
                seen.add(value)
                values.append(value)
        return values[:12]

    # ---------- UI refresh ----------
    def refresh_ui(self) -> None:
        jam = self.current_jam()
        self.period_label.setText(str(jam.period_number))
        self.jam_label.setText(str(jam.jam_number))
        self.time_label.setText(f"{self.current_time:.3f}s")
        self.start_label.setText("-" if jam.start_time is None else f"{jam.start_time:.3f}s")
        self.end_label.setText("-" if jam.end_time is None else f"{jam.end_time:.3f}s")
        self.home_jammer_label.setText(jam.home_jammer or "-")
        self.away_jammer_label.setText(jam.away_jammer or "-")
        self.home_lineup_label.setText(", ".join(jam.home_lineup) if jam.home_lineup else "-")
        self.away_lineup_label.setText(", ".join(jam.away_lineup) if jam.away_lineup else "-")
        hs = "-" if jam.home_score_start is None else str(jam.home_score_start)
        as_ = "-" if jam.away_score_start is None else str(jam.away_score_start)
        he = "-" if jam.home_score_end is None else str(jam.home_score_end)
        ae = "-" if jam.away_score_end is None else str(jam.away_score_end)
        self.score_start_label.setText(f"{hs} / {as_}")
        self.home_pass_label.setText(self._passes_to_text(jam.home_passes))
        self.away_pass_label.setText(self._passes_to_text(jam.away_passes))
        self.score_end_label.setText(f"{he} / {ae}")
        self.star_label.setText(f"{'Y' if jam.home_star_pass else 'N'} / {'Y' if jam.away_star_pass else 'N'}")
        self.lead_label.setText(jam.lead_jammer)
        self.penalties_label.setText(str(len(jam.penalties)))
        if self.notes_edit.toPlainText() != jam.notes:
            self.notes_edit.blockSignals(True)
            self.notes_edit.setPlainText(jam.notes)
            self.notes_edit.blockSignals(False)
        self.refresh_edit_panel()

    def refresh_edit_panel(self) -> None:
        jam = self.current_jam()
        values = {
            "home_jammer": jam.home_jammer,
            "away_jammer": jam.away_jammer,
            "home_pivot": self._lineup_value_at(jam.home_lineup, 0),
            "home_blocker_1": self._lineup_value_at(jam.home_lineup, 1),
            "home_blocker_2": self._lineup_value_at(jam.home_lineup, 2),
            "home_blocker_3": self._lineup_value_at(jam.home_lineup, 3),
            "away_pivot": self._lineup_value_at(jam.away_lineup, 0),
            "away_blocker_1": self._lineup_value_at(jam.away_lineup, 1),
            "away_blocker_2": self._lineup_value_at(jam.away_lineup, 2),
            "away_blocker_3": self._lineup_value_at(jam.away_lineup, 3),
            "home_score_start": "" if jam.home_score_start is None else str(jam.home_score_start),
            "away_score_start": "" if jam.away_score_start is None else str(jam.away_score_start),
            "home_passes": ", ".join(str(p) for p in jam.home_passes),
            "away_passes": ", ".join(str(p) for p in jam.away_passes),
            "home_score_end": "" if jam.home_score_end is None else str(jam.home_score_end),
            "away_score_end": "" if jam.away_score_end is None else str(jam.away_score_end),
            "home_star_pass": jam.home_star_pass,
            "away_star_pass": jam.away_star_pass,
            "lead_jammer": jam.lead_jammer,
            "notes": jam.notes,
        }
        for field_name, widget in self.edit_field_widgets.items():
            widget.blockSignals(True)
            value = values[field_name]
            if isinstance(widget, QLineEdit):
                if widget.text() != value:
                    widget.setText(value)
            elif isinstance(widget, QCheckBox):
                widget.setChecked(bool(value))
            elif isinstance(widget, QComboBox):
                idx = widget.findText(value)
                if idx >= 0:
                    widget.setCurrentIndex(idx)
            elif isinstance(widget, QPlainTextEdit):
                if widget.toPlainText() != value:
                    widget.setPlainText(value)
            widget.blockSignals(False)
            prev_values = self._collect_previous_values(field_name)
            self.edit_prev_labels[field_name].setText("Previously used: " + (" | ".join(prev_values) if prev_values else "-"))

    def refresh_penalty_panel(self) -> None:
        self.penalty_skater_edit.clear()
        self.penalty_code_combo.setCurrentIndex(0)
        self.penalty_team_combo.setCurrentText("home")

    def refresh_help_panel(self) -> None:
        help_text = """
Annotation files
  Open Annotation JSON ... Open an annotation file first
  Save As JSON .......... Save annotations to a JSON file
  The video source link is stored inside the same JSON file

Video source
  Paste a YouTube URL or video ID in the source box
  Click Apply / Load Video to load or replace the video
  If a JSON has no source link, the app will prompt for one on open

Playback
  Space ........ Play / Pause video
  a / f ........ Step backward / forward 1 second
  j / l ........ Jump backward / forward 5 seconds
  , / . ........ Approximate one-frame step backward / forward and pause playback

Jam navigation
  n ............ Go to next jam
  p ............ Go to previous jam
  [ ............ Jump to previous saved jam start
  ] ............ Jump to next saved jam start
  \ ........... Jump to current jam start
  { ............ Jump to previous saved jam end
  } ............ Jump to next saved jam end

Marking jam boundaries
  s ............ Mark jam start at current time
  e ............ Mark jam end at current time
  d ............ Delete current jam if it is empty

Game structure
  5 ............ Start next period and reset jam count

Scoring passes
  Shift+Q/W/E/R/T ... Add home pass worth 0/1/2/3/4
  Shift+A/S/D/F/G ... Add away pass worth 0/1/2/3/4
  z ................. Undo last home pass
  x ................. Undo last away pass

Star passes
  In Edit panel, move to Home star pass or Away star pass
  ` ................. Toggle current star pass checkbox

Lead jammer
  1 ............ Home lead
  2 ............ Away lead
  3 ............ No lead
  4 ............ Unknown / clear lead

Editing jam details
  m ............ Show / hide edit panel
  Tab / Enter .. Move to next edit field
  Up / Down .... Move between edit fields
  w ............ Save annotations and close edit panel
  Esc .......... Close edit / penalty / help panel

Penalties
  b ............ Show / hide penalty panel
  Enter ........ Save penalty

Playback speed
  - ............ Slow down playback
  = or + ....... Speed up playback
  0 ............ Reset to normal speed

General
  Ctrl+S ....... Save annotations
  h ............ Show / hide help
""".strip()
        self.help_text.setPlainText(help_text)

    # ---------- Player ----------
    def on_player_time_update(self, seconds: float, player_state: int) -> None:
        now = time.monotonic()
        seconds = max(0.0, seconds)
        self.player_state = player_state

        if now < self.ignore_player_updates_until and self.last_seek_target is not None:
            # Ignore stale player callbacks immediately after a seek.
            if abs(seconds - self.last_seek_target) > 1.0:
                return
            self.ignore_player_updates_until = 0.0

        self.current_time = seconds
        self.time_label.setText(f"{self.current_time:.3f}s")

    def run_js(self, js: str) -> None:
        self.player_view.page().runJavaScript(js)

    def toggle_play_pause(self) -> None:
        if self.player_state == 1:
            self.run_js("pauseVideo();")
        else:
            self.run_js("playVideo();")

    def seek_relative(self, delta: float) -> None:
        now = time.monotonic()
        base_time = self.current_time
        if now < self.ignore_player_updates_until and self.last_seek_target is not None:
            base_time = self.last_seek_target
        self.seek_to(max(0.0, base_time + delta))

    def step_frame(self, direction: int) -> None:
        if self.player_state == 1:
            self.run_js("pauseVideo();")
        self.seek_to(max(0.0, self.current_time + (DEFAULT_FRAME_STEP_SECONDS * direction)))

    def seek_to(self, seconds: float) -> None:
        seconds = max(0.0, seconds)
        self.last_seek_target = seconds
        self.ignore_player_updates_until = time.monotonic() + 0.4
        self.current_time = seconds
        self.time_label.setText(f"{self.current_time:.3f}s")
        self.run_js(f"seekTo({seconds});")

    def adjust_playback_speed(self, direction: int) -> None:
        try:
            current_index = PLAYBACK_SPEED_STEPS.index(self.playback_rate)
        except ValueError:
            current_index = PLAYBACK_SPEED_STEPS.index(1.0)
        new_index = max(0, min(len(PLAYBACK_SPEED_STEPS) - 1, current_index + direction))
        self.playback_rate = PLAYBACK_SPEED_STEPS[new_index]
        self.run_js(f"setPlaybackRate({self.playback_rate});")

    def reset_playback_speed(self) -> None:
        self.playback_rate = 1.0
        self.run_js("setPlaybackRate(1.0);")

    # ---------- Source ----------
    def extract_video_id(self, text: str) -> Optional[str]:
        text = text.strip()
        if re.fullmatch(r"[A-Za-z0-9_-]{11}", text):
            return text
        patterns = [
            r"v=([A-Za-z0-9_-]{11})",
            r"youtu\.be/([A-Za-z0-9_-]{11})",
            r"embed/([A-Za-z0-9_-]{11})",
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        return None

    def load_video_from_source_text(self, text: str, prompt_on_failure: bool = True) -> bool:
        video_id = self.extract_video_id(text)
        if not video_id:
            if prompt_on_failure:
                QMessageBox.warning(self, "Invalid video link", "Please enter a valid YouTube URL or video ID.")
            return False
        self.youtube_id = video_id
        self.youtube_url = text.strip()
        self.data.source = self.youtube_url or video_id
        self.url_edit.blockSignals(True)
        self.url_edit.setText(self.youtube_url)
        self.url_edit.blockSignals(False)
        self.run_js(f"loadVideo('{video_id}');")
        self.source_status_label.setText("Loaded video source from annotation file.")
        return True

    def prompt_for_source_link_if_missing(self) -> None:
        while True:
            text, ok = QInputDialog.getText(
                self,
                "Add Video Source",
                "This annotation file has no video source link yet. Paste a YouTube URL or video ID:",
                text=self.url_edit.text().strip(),
            )
            if not ok:
                self.source_status_label.setText("No video source link is set for this annotation file.")
                return
            if self.load_video_from_source_text(text, prompt_on_failure=False):
                self.save_if_path()
                return
            QMessageBox.warning(self, "Invalid video link", "Please enter a valid YouTube URL or video ID.")

    def apply_source_link(self) -> None:
        text = self.url_edit.text().strip()
        if self.load_video_from_source_text(text, prompt_on_failure=True):
            self.save_if_path()

    # ---------- Annotation actions ----------
    def mark_start(self) -> None:
        jam = self.current_jam()
        jam.start_time = round(self.current_time, 3)
        self.refresh_ui()
        self.save_if_path()

    def mark_end(self) -> None:
        jam = self.current_jam()
        jam.end_time = round(self.current_time, 3)
        self.refresh_ui()
        self.save_if_path()

    def append_pass(self, team: str, points: int) -> None:
        jam = self.current_jam()
        if team == "home":
            jam.home_passes.append(points)
        else:
            jam.away_passes.append(points)
        self._recalculate_jam_scores(jam)
        self._propagate_scores_from_jam(self.current_jam_index)
        self.refresh_ui()
        self.save_if_path()

    def undo_last_pass(self, team: str) -> None:
        jam = self.current_jam()
        passes = jam.home_passes if team == "home" else jam.away_passes
        if passes:
            passes.pop()
            self._recalculate_jam_scores(jam)
            self._propagate_scores_from_jam(self.current_jam_index)
            self.refresh_ui()
            self.save_if_path()

    def set_lead(self, lead: str) -> None:
        jam = self.current_jam()
        jam.lead_jammer = lead
        self.refresh_ui()
        self.save_if_path()

    def next_jam(self) -> None:
        previous_index = self.current_jam_index
        self.current_jam_index += 1
        jam = self.current_jam()
        self._apply_score_defaults(jam)
        self._prefill_next_jam_from_recent_penalties(previous_index, self.current_jam_index)
        self.refresh_ui()

    def previous_jam(self) -> None:
        self.current_jam_index = max(0, self.current_jam_index - 1)
        self.refresh_ui()

    def start_next_period(self) -> None:
        current = self.current_jam()
        next_index = self.current_jam_index + 1
        if next_index >= len(self.data.jams):
            self.data.jams.append(JamRecord(period_number=current.period_number + 1, jam_number=1))
        else:
            self.data.jams[next_index].period_number = current.period_number + 1
            self.data.jams[next_index].jam_number = 1
        self.current_jam_index = next_index
        jam = self.current_jam()
        self._apply_score_defaults(jam)
        self.refresh_ui()
        self.save_if_path()

    def jump_to_current_jam_start(self) -> None:
        jam = self.current_jam()
        if jam.start_time is not None:
            self.seek_to(jam.start_time)

    def jump_to_previous_saved_jam_start(self) -> None:
        for idx in range(self.current_jam_index - 1, -1, -1):
            jam = self.data.jams[idx]
            if jam.start_time is not None:
                self.current_jam_index = idx
                self.seek_to(jam.start_time)
                self.refresh_ui()
                return

    def jump_to_next_saved_jam_start(self) -> None:
        for idx in range(self.current_jam_index + 1, len(self.data.jams)):
            jam = self.data.jams[idx]
            if jam.start_time is not None:
                self.current_jam_index = idx
                self.seek_to(jam.start_time)
                self.refresh_ui()
                return

    def jump_to_previous_saved_jam_end(self) -> None:
        for idx in range(self.current_jam_index - 1, -1, -1):
            jam = self.data.jams[idx]
            if jam.end_time is not None:
                self.current_jam_index = idx
                self.seek_to(jam.end_time)
                self.refresh_ui()
                return

    def jump_to_next_saved_jam_end(self) -> None:
        for idx in range(self.current_jam_index + 1, len(self.data.jams)):
            jam = self.data.jams[idx]
            if jam.end_time is not None:
                self.current_jam_index = idx
                self.seek_to(jam.end_time)
                self.refresh_ui()
                return

    def delete_current_jam_if_empty(self) -> None:
        if not self.data.jams:
            return
        jam = self.current_jam()
        is_empty = (
            jam.start_time is None
            and jam.end_time is None
            and not jam.home_jammer
            and not jam.away_jammer
            and not jam.home_lineup
            and not jam.away_lineup
            and not jam.home_passes
            and not jam.away_passes
            and jam.lead_jammer == "unknown"
            and not jam.penalties
            and jam.home_score_start is None
            and jam.away_score_start is None
            and jam.home_score_end is None
            and jam.away_score_end is None
            and not jam.home_star_pass
            and not jam.away_star_pass
            and not jam.notes
        )
        if is_empty and len(self.data.jams) > 1:
            self.data.jams.pop(self.current_jam_index)
            self.current_jam_index = min(self.current_jam_index, len(self.data.jams) - 1)
            self.refresh_ui()
            self.save_if_path()

    # ---------- Edit panel ----------
    def toggle_edit_panel(self) -> None:
        showing = not self.edit_box.isVisible()
        self.edit_box.setVisible(showing)
        if showing:
            self._focus_edit_field(self.edit_field_index, ensure_visible=True)
        else:
            self.player_view.setFocus()

    def on_edit_field_changed(self, field_name: str) -> None:
        jam = self.current_jam()
        widget = self.edit_field_widgets[field_name]

        if field_name in self.LINEUP_FIELD_MAP:
            lineup_attr, index = self.LINEUP_FIELD_MAP[field_name]
            lineup = list(getattr(jam, lineup_attr))
            setattr(jam, lineup_attr, self._set_lineup_value_at(lineup, index, widget.text()))
        elif field_name in {"home_jammer", "away_jammer"}:
            setattr(jam, field_name, widget.text().strip())
        elif field_name in {"home_score_start", "away_score_start", "home_score_end", "away_score_end"}:
            text = widget.text().strip()
            value = None if text == "" else int(text) if text.lstrip("-").isdigit() else getattr(jam, field_name)
            setattr(jam, field_name, value)
            if field_name in {"home_score_start", "away_score_start"}:
                self._recalculate_jam_scores(jam)
                self._propagate_scores_from_jam(self.current_jam_index)
        elif field_name in {"home_passes", "away_passes"}:
            text = widget.text().strip()
            tokens = [t.strip() for t in text.replace("+", ",").replace(";", ",").split(",") if t.strip()]
            parsed = []
            for token in tokens:
                try:
                    parsed.append(int(token))
                except ValueError:
                    pass
            setattr(jam, field_name, parsed)
            self._recalculate_jam_scores(jam)
            self._propagate_scores_from_jam(self.current_jam_index)
        elif field_name in {"home_star_pass", "away_star_pass"}:
            setattr(jam, field_name, widget.isChecked())
        elif field_name == "lead_jammer":
            jam.lead_jammer = widget.currentText()
        elif field_name == "notes":
            jam.notes = widget.toPlainText()

        self.refresh_ui()
        self.save_if_path()

    # ---------- Penalties ----------
    def toggle_penalty_panel(self) -> None:
        showing = not self.penalty_box.isVisible()
        self.penalty_box.setVisible(showing)
        if showing:
            self.penalty_skater_edit.setFocus()
        else:
            self.player_view.setFocus()

    def save_penalty(self) -> None:
        jam = self.current_jam()
        skater = self.penalty_skater_edit.text().strip()
        code = self.penalty_code_combo.currentData()
        team = self.penalty_team_combo.currentText()
        if not skater or not code:
            return
        jam.penalties.append({
            "team": team,
            "skater": skater,
            "code": code,
            "time": round(self.current_time, 3),
        })
        self.refresh_penalty_panel()
        self.refresh_ui()
        self.save_if_path()

    def cancel_penalty(self) -> None:
        self.refresh_penalty_panel()
        self.penalty_box.hide()
        self.player_view.setFocus()

    # ---------- Notes ----------
    def notes_changed(self) -> None:
        jam = self.current_jam()
        jam.notes = self.notes_edit.toPlainText()
        self.save_if_path()

    # ---------- Help ----------
    def toggle_help_panel(self) -> None:
        self.help_box.setVisible(not self.help_box.isVisible())

    # ---------- Save/load ----------
    def save_payload(self) -> dict:
        self.data.source = self.url_edit.text().strip() or self.data.source
        return {
            "source": self.data.source,
            "youtube_id": self.youtube_id,
            "jams": [asdict(j) for j in self.data.jams],
        }

    def save_if_path(self) -> None:
        if self.output_path is not None:
            self.save_annotations()

    def save_annotations(self) -> None:
        if self.output_path is None:
            self.save_annotations_as()
            return
        self.output_path.write_text(json.dumps(self.save_payload(), indent=2), encoding="utf-8")
        self.annotation_file_label.setText(self.output_path.name)
        self.statusBar().showMessage(f"Saved to {self.output_path}", 3000)

    def save_annotations_as(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Save annotations", "youtube_annotations.json", "JSON files (*.json)")
        if not path:
            return
        self.output_path = Path(path)
        self.annotation_file_label.setText(self.output_path.name)
        self.save_annotations()

    def open_annotation_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Open annotations", "", "JSON files (*.json)")
        if not path:
            return
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
        self.output_path = Path(path)
        self.annotation_file_label.setText(self.output_path.name)
        self.data = AnnotationFile(
            source=payload.get("source", ""),
            jams=[JamRecord(**jam) for jam in payload.get("jams", [])] or [JamRecord()],
        )
        self.current_jam_index = 0

        source_text = payload.get("source", "") or payload.get("youtube_id", "")
        self.youtube_id = payload.get("youtube_id", "")
        self.youtube_url = source_text
        self.url_edit.setText(source_text)

        if source_text:
            self.load_video_from_source_text(source_text, prompt_on_failure=True)
        else:
            self.prompt_for_source_link_if_missing()

        self.refresh_ui()


def main() -> None:
    app = QApplication(sys.argv)
    window = YouTubeAnnotatorWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
