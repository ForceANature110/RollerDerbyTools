import argparse
import json
import os
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import List, Optional

import cv2
import numpy as np

try:
    import vlc  # python-vlc
except ImportError:
    vlc = None


WINDOW_NAME = "FaN's Derby Jam Annotator"
SEEK_SECONDS_SMALL = 1.0
SEEK_SECONDS_LARGE = 5.0
DEFAULT_FPS_FALLBACK = 30.0
PLAYBACK_SPEED_STEPS = [0.25, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0]
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
    video_file: str
    jams: List[JamRecord] = field(default_factory=list)


class JamAnnotator:
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

    def __init__(self, video_path: str, output_path: str) -> None:
        self.video_path = video_path
        self.output_path = output_path

        self.cap = cv2.VideoCapture(video_path)
        if not self.cap.isOpened():
            raise RuntimeError(f"Could not open video: {video_path}")

        fps = self.cap.get(cv2.CAP_PROP_FPS)
        self.fps = fps if fps and fps > 0 else DEFAULT_FPS_FALLBACK
        self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
        self.duration_seconds = self.total_frames / self.fps if self.total_frames > 0 else 0.0

        self.playing = False
        self.playback_speed = 1.0
        self.audio_instance = None
        self.audio_player = None
        self.audio_enabled = False
        self._init_audio()
        self.current_frame_index = 0
        self.current_frame = None

        self.data = self._load_or_create_annotations()
        self.current_jam_index = 0
        self.edit_mode = False
        self.edit_field_index = 0
        self.edit_buffer = ""
        self.show_help = False
        self.penalty_mode = False
        self.penalty_skater = ""
        self.penalty_code = ""
        self.penalty_team = "home"
        self.penalty_focus = "skater"
        if not self.data.jams:
            self.data.jams.append(JamRecord(period_number=1, jam_number=1))

    def _init_audio(self) -> None:
        if vlc is None:
            print("Audio disabled: install python-vlc and VLC media player to enable sound.")
            return
        try:
            self.audio_instance = vlc.Instance("--no-video")
            media = self.audio_instance.media_new(self.video_path)
            self.audio_player = self.audio_instance.media_player_new()
            self.audio_player.set_media(media)
            self.audio_enabled = True
            print("Audio enabled via VLC")
        except Exception as exc:
            print(f"Audio disabled: could not initialize VLC audio ({exc})")
            self.audio_instance = None
            self.audio_player = None
            self.audio_enabled = False

    def _sync_audio_play_state(self) -> None:
        if not self.audio_enabled or self.audio_player is None:
            return
        try:
            if self.playing:
                self.audio_player.play()
                self.audio_player.set_time(int(self.current_time_seconds() * 1000))
                try:
                    self.audio_player.set_rate(max(0.25, min(4.0, float(self.playback_speed))))
                except Exception:
                    pass
            else:
                self.audio_player.pause()
        except Exception:
            pass

    def _sync_audio_seek(self) -> None:
        if not self.audio_enabled or self.audio_player is None:
            return
        try:
            state = self.audio_player.get_state()
            if state in (vlc.State.NothingSpecial, vlc.State.Stopped, vlc.State.Ended):
                self.audio_player.play()
            self.audio_player.set_time(int(self.current_time_seconds() * 1000))
            try:
                self.audio_player.set_rate(max(0.25, min(4.0, float(self.playback_speed))))
            except Exception:
                pass
            if not self.playing:
                self.audio_player.pause()
        except Exception:
            pass

    def _load_or_create_annotations(self) -> AnnotationFile:
        if os.path.exists(self.output_path):
            with open(self.output_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            jams = [JamRecord(**jam) for jam in raw.get("jams", [])]
            return AnnotationFile(video_file=raw.get("video_file", self.video_path), jams=jams)
        return AnnotationFile(video_file=self.video_path)

    def save(self) -> None:
        payload = {
            "video_file": self.video_path,
            "jams": [asdict(jam) for jam in self.data.jams],
        }
        with open(self.output_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        print(f"Saved annotations to {self.output_path}")

    def _default_period_for_new_jam(self) -> int:
        if not self.data.jams:
            return 1
        return self.data.jams[-1].period_number

    def _default_jam_number_for_new_jam(self) -> int:
        if not self.data.jams:
            return 1
        return self.data.jams[-1].jam_number + 1

    def ensure_current_jam(self) -> JamRecord:
        while self.current_jam_index >= len(self.data.jams):
            self.data.jams.append(JamRecord(period_number=self._default_period_for_new_jam(), jam_number=self._default_jam_number_for_new_jam()))
        jam = self.data.jams[self.current_jam_index]
        self._apply_score_defaults(jam)
        return jam

    def _apply_score_defaults(self, jam: JamRecord) -> None:
        if jam.jam_number <= 1:
            return
        prev_index = jam.jam_number - 2
        if prev_index < 0 or prev_index >= len(self.data.jams):
            return
        prev_jam = self.data.jams[prev_index]

        if jam.home_score_start is None and prev_jam.home_score_end is not None:
            jam.home_score_start = prev_jam.home_score_end
        if jam.away_score_start is None and prev_jam.away_score_end is not None:
            jam.away_score_start = prev_jam.away_score_end

        self._recalculate_jam_scores(jam)

    def _recalculate_jam_scores(self, jam: JamRecord) -> None:
        if jam.home_score_start is not None:
            jam.home_score_end = jam.home_score_start + sum(jam.home_passes)
        elif jam.home_score_end is None:
            jam.home_score_end = None

        if jam.away_score_start is not None:
            jam.away_score_end = jam.away_score_start + sum(jam.away_passes)
        elif jam.away_score_end is None:
            jam.away_score_end = None

    def _passes_to_text(self, passes: List[int]) -> str:
        return "-" if not passes else " + ".join(str(p) for p in passes)

    def _append_pass(self, team: str, points: int) -> None:
        jam = self.ensure_current_jam()
        if team == "home":
            jam.home_passes.append(points)
        else:
            jam.away_passes.append(points)
        self._recalculate_jam_scores(jam)
        self._propagate_scores_from_jam(self.current_jam_index)
        self.save()
        print(f"Jam {jam.jam_number}: added {points}-point {team} pass")

    def _undo_last_pass(self, team: str) -> None:
        jam = self.ensure_current_jam()
        passes = jam.home_passes if team == "home" else jam.away_passes
        if passes:
            removed = passes.pop()
            self._recalculate_jam_scores(jam)
            self._propagate_scores_from_jam(self.current_jam_index)
            self.save()
            print(f"Jam {jam.jam_number}: removed last {team} pass ({removed})")

    def current_time_seconds(self) -> float:
        return self.current_frame_index / self.fps

    def seek_to_time(self, time_seconds: float) -> None:
        clamped = max(0.0, min(self.duration_seconds, time_seconds))
        frame_index = int(round(clamped * self.fps))
        self.seek_to_frame(frame_index)

    def seek_to_frame(self, frame_index: int) -> None:
        clamped = max(0, min(self.total_frames - 1, frame_index if self.total_frames > 0 else 0))
        self.current_frame_index = clamped
        self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame_index)
        ok, frame = self.cap.read()
        if ok:
            self.current_frame = frame
        else:
            self.current_frame = self._blank_frame("Unable to read frame")
        self._sync_audio_seek()

    def _blank_frame(self, text: str):
        frame = 255 * (cv2.UMat(720, 1280, cv2.CV_8UC3).get())
        cv2.putText(frame, text, (50, 80), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 255), 2, cv2.LINE_AA)
        return frame

    def step_frames(self, delta: int) -> None:
        self.seek_to_frame(self.current_frame_index + delta)

    def step_seconds(self, delta_seconds: float) -> None:
        self.seek_to_time(self.current_time_seconds() + delta_seconds)

    def mark_start(self) -> None:
        jam = self.ensure_current_jam()
        jam.start_time = round(self.current_time_seconds(), 3)
        print(f"Jam {jam.jam_number}: marked start at {jam.start_time:.3f}s")
        self.save()

    def mark_end(self) -> None:
        jam = self.ensure_current_jam()
        jam.end_time = round(self.current_time_seconds(), 3)
        print(f"Jam {jam.jam_number}: marked end at {jam.end_time:.3f}s")
        self.save()

    def next_jam(self) -> None:
        previous_index = self.current_jam_index
        self.current_jam_index += 1
        jam = self.ensure_current_jam()
        self._apply_score_defaults(jam)
        self._prefill_next_jam_from_recent_penalties(previous_index, self.current_jam_index)
        print(f"Moved to period {jam.period_number}, jam {jam.jam_number}")

    def jump_to_current_jam_start(self) -> None:
        jam = self.ensure_current_jam()
        if jam.start_time is not None:
            self.playing = False
            self.seek_to_time(jam.start_time)
            print(f"Jumped to jam {jam.jam_number} start at {jam.start_time:.3f}s")

    def jump_to_next_saved_jam_start(self) -> None:
        for idx in range(self.current_jam_index + 1, len(self.data.jams)):
            jam = self.data.jams[idx]
            if jam.start_time is not None:
                self.current_jam_index = idx
                self.playing = False
                self.seek_to_time(jam.start_time)
                print(f"Jumped to jam {jam.jam_number} start at {jam.start_time:.3f}s")
                return
        print("No later saved jam start found")

    def jump_to_previous_saved_jam_end(self) -> None:
        for idx in range(self.current_jam_index - 1, -1, -1):
            jam = self.data.jams[idx]
            if jam.end_time is not None:
                self.current_jam_index = idx
                self.playing = False
                self.seek_to_time(jam.end_time)
                print(f"Jumped to jam {jam.jam_number} end at {jam.end_time:.3f}s")
                return
        print("No earlier saved jam end found")

    def jump_to_next_saved_jam_end(self) -> None:
        for idx in range(self.current_jam_index + 1, len(self.data.jams)):
            jam = self.data.jams[idx]
            if jam.end_time is not None:
                self.current_jam_index = idx
                self.playing = False
                self.seek_to_time(jam.end_time)
                print(f"Jumped to jam {jam.jam_number} end at {jam.end_time:.3f}s")
                return
        print("No later saved jam end found")

    def jump_to_previous_saved_jam_start(self) -> None:
        for idx in range(self.current_jam_index - 1, -1, -1):
            jam = self.data.jams[idx]
            if jam.start_time is not None:
                self.current_jam_index = idx
                self.playing = False
                self.seek_to_time(jam.start_time)
                print(f"Jumped to jam {jam.jam_number} start at {jam.start_time:.3f}s")
                return
        print("No earlier saved jam start found")

    def previous_jam(self) -> None:
        self.current_jam_index = max(0, self.current_jam_index - 1)
        current = self.ensure_current_jam()
        print(f"Moved to period {current.period_number}, jam {current.jam_number}")

    def start_next_period(self) -> None:
        current = self.ensure_current_jam()
        next_index = self.current_jam_index + 1

        if next_index >= len(self.data.jams):
            self.data.jams.append(JamRecord(period_number=current.period_number + 1, jam_number=1))
        else:
            self.data.jams[next_index].period_number = current.period_number + 1
            self.data.jams[next_index].jam_number = 1
            for idx in range(next_index + 1, len(self.data.jams)):
                if self.data.jams[idx].period_number == current.period_number + 1:
                    self.data.jams[idx].jam_number = self.data.jams[idx - 1].jam_number + 1

        self.current_jam_index = next_index
        jam = self.ensure_current_jam()
        self._apply_score_defaults(jam)
        print(f"Started period {jam.period_number}, jam {jam.jam_number}")
        self.save()

    def _propagate_scores_from_jam(self, changed_jam_index: int) -> None:
        if changed_jam_index < 0 or changed_jam_index >= len(self.data.jams):
            return

        self._apply_score_defaults(self.data.jams[changed_jam_index])

        for idx in range(changed_jam_index + 1, len(self.data.jams)):
            prev_jam = self.data.jams[idx - 1]
            jam = self.data.jams[idx]

            old_home_start = jam.home_score_start
            old_away_start = jam.away_score_start

            jam.home_score_start = prev_jam.home_score_end
            jam.away_score_start = prev_jam.away_score_end
            self._recalculate_jam_scores(jam)

            if jam.home_score_start == old_home_start and jam.away_score_start == old_away_start:
                # Still continue, because a prior jam change can affect all later end scores.
                pass

    def delete_current_jam_if_empty(self) -> None:
        if not self.data.jams:
            return
        jam = self.ensure_current_jam()
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
            removed = self.data.jams.pop(self.current_jam_index)
            for i, existing in enumerate(self.data.jams, start=1):
                existing.jam_number = i
            self.current_jam_index = min(self.current_jam_index, len(self.data.jams) - 1)
            print(f"Deleted empty jam {removed.jam_number}")
            self.save()

    def undo_last_boundary(self) -> None:
        jam = self.ensure_current_jam()
        if jam.end_time is not None:
            print(f"Jam {jam.jam_number}: cleared end")
            jam.end_time = None
        elif jam.start_time is not None:
            print(f"Jam {jam.jam_number}: cleared start")
            jam.start_time = None
        else:
            print(f"Jam {jam.jam_number}: nothing to undo")
        self.save()

    def begin_edit_mode(self) -> None:
        self.edit_mode = True
        self.edit_field_index = 0
        self.edit_buffer = self._field_to_buffer(TEXT_INPUT_FIELDS[self.edit_field_index])

    def end_edit_mode(self, save_changes: bool) -> None:
        if save_changes:
            self._commit_current_edit_field()
            self.save()
        self.edit_mode = False
        self.edit_field_index = 0
        self.edit_buffer = ""

    def begin_penalty_mode(self) -> None:
        self.penalty_mode = True
        self.penalty_skater = ""
        self.penalty_code = ""
        self.penalty_team = "home"
        self.penalty_focus = "skater"

    def end_penalty_mode(self, save_penalty: bool) -> None:
        if save_penalty:
            jam = self.ensure_current_jam()
            skater = self.penalty_skater.strip()
            code = self.penalty_code.strip().upper()
            if skater and code:
                jam.penalties.append({
                    "team": self.penalty_team,
                    "skater": skater,
                    "code": code,
                    "time": round(self.current_time_seconds(), 3),
                })
                self.save()
                print(f"Jam {jam.jam_number}: added penalty {code} for {self.penalty_team} skater {skater}")
        self.penalty_mode = False
        self.penalty_skater = ""
        self.penalty_code = ""
        self.penalty_team = "home"
        self.penalty_focus = "skater"

    def _handle_penalty_key(self, key: int) -> bool:
        valid_penalty_codes = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "P", "X"}
        if key == 255:
            return True
        if key == 27:
            self.end_penalty_mode(save_penalty=False)
            return True
        if key in (13, 10):
            self.end_penalty_mode(save_penalty=True)
            return True
        if key == 9:
            if self.penalty_focus == "skater":
                self.penalty_focus = "code"
            elif self.penalty_focus == "code":
                self.penalty_focus = "team"
            else:
                self.penalty_focus = "skater"
            return True
        if key in (8, 127):
            if self.penalty_focus == "skater":
                self.penalty_skater = self.penalty_skater[:-1]
            elif self.penalty_focus == "code":
                self.penalty_code = self.penalty_code[:-1]
            else:
                self.penalty_team = "away" if self.penalty_team == "home" else "home"
            return True
        if 32 <= key <= 126:
            ch = chr(key)
            if self.penalty_focus == "skater" and ch.isdigit():
                self.penalty_skater += ch
                return True
            if self.penalty_focus == "code" and ch.isalpha():
                up = ch.upper()
                if up in valid_penalty_codes:
                    self.penalty_code = up
                return True
            if self.penalty_focus == "team" and ch.lower() in {"h", "a"}:
                self.penalty_team = "home" if ch.lower() == "h" else "away"
                return True
        return False

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
                    lineup = self._set_lineup_value_at(lineup, 0, skater)
                    setattr(next_jam, lineup_attr, lineup)
            elif role == "blocker" and slot_index is not None:
                lineup_attr = "home_lineup" if team == "home" else "away_lineup"
                lineup = list(getattr(next_jam, lineup_attr))
                if self._lineup_value_at(lineup, slot_index) == "":
                    lineup = self._set_lineup_value_at(lineup, slot_index, skater)
                    setattr(next_jam, lineup_attr, lineup)

    def _lineup_value_at(self, lineup: List[str], index: int) -> str:
        return lineup[index] if index < len(lineup) else ""

    def _set_lineup_value_at(self, lineup: List[str], index: int, value: str) -> List[str]:
        while len(lineup) <= index:
            lineup.append("")
        lineup[index] = value.strip()
        while lineup and lineup[-1] == "":
            lineup.pop()
        return lineup

    def _field_to_buffer(self, field_name: str) -> str:
        jam = self.ensure_current_jam()
        if field_name in self.LINEUP_FIELD_MAP:
            lineup_attr, index = self.LINEUP_FIELD_MAP[field_name]
            lineup = getattr(jam, lineup_attr)
            return self._lineup_value_at(lineup, index)

        value = getattr(jam, field_name)
        if isinstance(value, bool):
            return "y" if value else "n"
        if value is None:
            return ""
        return str(value)

    def _assign_field_from_buffer(self, field_name: str, buffer: str) -> None:
        jam = self.ensure_current_jam()
        text = buffer.strip()

        if field_name in {"home_jammer", "away_jammer", "notes"}:
            setattr(jam, field_name, text)
            return

        if field_name in {"home_passes", "away_passes"}:
            if text == "":
                setattr(jam, field_name, [])
            else:
                tokens = [token.strip() for token in text.replace("+", ",").replace(";", ",").split(",")]
                parsed = []
                for token in tokens:
                    if token:
                        try:
                            parsed.append(int(token))
                        except ValueError:
                            pass
                setattr(jam, field_name, parsed)
            self._recalculate_jam_scores(jam)
            self._propagate_scores_from_jam(self.current_jam_index)
            return

        if field_name in self.LINEUP_FIELD_MAP:
            lineup_attr, index = self.LINEUP_FIELD_MAP[field_name]
            lineup = list(getattr(jam, lineup_attr))
            updated = self._set_lineup_value_at(lineup, index, text)
            setattr(jam, lineup_attr, updated)
            return

        if field_name in {"home_score_start", "away_score_start", "home_score_end", "away_score_end"}:
            if text == "" or text.lower() in {"none", "null", "clear", "x"}:
                setattr(jam, field_name, None)
                if field_name in {"home_score_start", "away_score_start"}:
                    self._recalculate_jam_scores(jam)
                    self._propagate_scores_from_jam(self.current_jam_index)
                return
            try:
                setattr(jam, field_name, int(text))
                if field_name in {"home_score_start", "away_score_start"}:
                    self._recalculate_jam_scores(jam)
                    self._propagate_scores_from_jam(self.current_jam_index)
            except ValueError:
                pass
            return

        if field_name in {"home_star_pass", "away_star_pass"}:
            if text == "":
                return
            normalized = text.lower()
            setattr(jam, field_name, normalized in {"y", "yes", "1", "true", "t"})
            return

        if field_name == "lead_jammer":
            normalized = text.strip().lower()
            if normalized in {"home", "away", "none", "unknown"}:
                jam.lead_jammer = normalized
            elif normalized in {"h", "1"}:
                jam.lead_jammer = "home"
            elif normalized in {"a", "2"}:
                jam.lead_jammer = "away"
            elif normalized in {"n", "3"}:
                jam.lead_jammer = "none"
            elif normalized in {"u", "4", "", "clear"}:
                jam.lead_jammer = "unknown"
            return

    def _commit_current_edit_field(self) -> None:
        field_name = TEXT_INPUT_FIELDS[self.edit_field_index]
        self._assign_field_from_buffer(field_name, self.edit_buffer)

        jam = self.ensure_current_jam()
        if field_name in {"home_score_start", "away_score_start", "home_passes", "away_passes"}:
            self._recalculate_jam_scores(jam)
            self._propagate_scores_from_jam(self.current_jam_index)

    def _move_edit_field(self, delta: int) -> None:
        self._commit_current_edit_field()
        self.edit_field_index = (self.edit_field_index + delta) % len(TEXT_INPUT_FIELDS)
        self.edit_buffer = self._field_to_buffer(TEXT_INPUT_FIELDS[self.edit_field_index])

    def _handle_edit_key(self, key: int) -> bool:
        if key == 255:
            return True
        if key == 27:
            self.end_edit_mode(save_changes=False)
            return True
        if key in (13, 10):
            self._move_edit_field(1)
            return True
        if key == 9:
            self._move_edit_field(1)
            return True
        if key in (8, 127):
            self.edit_buffer = self.edit_buffer[:-1]
            return True
        if key in (ord("`"), ord("~")):
            current_field = TEXT_INPUT_FIELDS[self.edit_field_index]
            if current_field in {"home_star_pass", "away_star_pass"}:
                self.edit_buffer = "n" if self.edit_buffer.lower() in {"y", "yes", "1", "true", "t"} else "y"
            return True
        if key in (82,):
            self._move_edit_field(-1)
            return True
        if key in (84,):
            self._move_edit_field(1)
            return True
        if 32 <= key <= 126:
            self.edit_buffer += chr(key)
            return True
        return False

    def draw_overlay(self, frame):
        jam = self.ensure_current_jam()
        overlay = frame.copy()
        h, w = overlay.shape[:2]
        panel_width = min(560, w - 20)
        panel_height = min(300, h - 20)
        x0, y0 = 10, 10
        x1, y1 = x0 + panel_width, y0 + panel_height

        cv2.rectangle(overlay, (x0, y0), (x1, y1), (0, 0, 0), -1)
        cv2.addWeighted(overlay, 0.45, frame, 0.55, 0, frame)

        lines = [
            f"Video: {Path(self.video_path).name}",
            f"Time: {self.current_time_seconds():.3f}s   Frame: {self.current_frame_index}/{max(self.total_frames - 1, 0)}   Speed: {self.playback_speed:.2f}x",
            f"Period: {jam.period_number}   Jam: {jam.jam_number}{'   [EDIT MODE]' if self.edit_mode else ''}",
            f"Start: {self._fmt_time(jam.start_time)}   End: {self._fmt_time(jam.end_time)}",
            f"Home jammer: {jam.home_jammer or '-'}   Away jammer: {jam.away_jammer or '-'}",
            f"Home lineup: {self._fmt_lineup(jam.home_lineup)}",
            f"Away lineup: {self._fmt_lineup(jam.away_lineup)}",
            f"Score start H/A: {self._fmt_int(jam.home_score_start)}/{self._fmt_int(jam.away_score_start)}",
            f"Passes home: {self._passes_to_text(jam.home_passes)}",
            f"Passes away: {self._passes_to_text(jam.away_passes)}",
            f"Score end   H/A: {self._fmt_int(jam.home_score_end)}/{self._fmt_int(jam.away_score_end)}",
            f"Star pass H/A: {'Y' if jam.home_star_pass else 'N'}/{'Y' if jam.away_star_pass else 'N'}",
            f"Lead: {jam.lead_jammer.upper() if jam.lead_jammer else 'UNKNOWN'}",
            f"Penalties: {len(jam.penalties)}",
            f"Notes: {jam.notes or '-'}",
        ]

        help_lines = [
            "Playback:",
            "  Space  – Play / Pause video",
            "  a / f  – Step backward / forward 1 second",
            "  j / l  – Jump backward / forward 5 seconds",
            "  , / .  – Step one frame backward / forward",
            "",
            "Jam navigation:",
            "  n      – Go to next jam",
            "  p      – Go to previous jam",
            "  [      – Jump to previous saved jam start",
            "  ]      – Jump to next saved jam start",
            "  \\      – Jump to current jam start",
            "  {      – Jump to previous saved jam end",
            "  }      – Jump to next saved jam end",
            "",
            "Marking jam boundaries:",
            "  s      – Mark jam start at current time",
            "  e      – Mark jam end at current time",
            "  u      – Undo last boundary (end first, then start)",
            "",
            "Game structure:",
            "  5      – Start next period (resets jam numbering)",
            "",
            "Scoring passes:",
            "  Shift+Q/W/E/R/T – Add home pass worth 0/1/2/3/4 points",
            "  Shift+A/S/D/F/G – Add away pass worth 0/1/2/3/4 points",
            "  z                – Undo last home pass",
            "  x                – Undo last away pass",
            "",
            "Star passes:",
            "  In Edit mode (`m`):",
            "  Navigate to 'Home star pass' or 'Away star pass'",
            "  ` (backtick) – Toggle star pass Y/N",
            "",
            "Lead jammer:",
            "  1 – Home lead",
            "  2 – Away lead",
            "  3 – No lead",
            "  4 – Unknown / clear lead",
            "",
            "Editing jam details:",
            "  m      – Open edit panel for lineups, jammers, scores, notes",
            "  w      – Save edits when in edit mode",
            "",
            "Penalties:",
            "  b      – Open penalty dialog",
            "  Tab    – Cycle penalty fields (skater / code / team)",
            "  Enter  – Save penalty",
            "  Esc    – Cancel penalty entry",
            "",
            "Playback speed:",
            "  -      – Slow down playback",
            "  = or + – Speed up playback",
            "  0      – Reset to normal speed",
            "",
            "General:",
            "  w      – Save annotations",
            "  q      – Quit program",
            "  h      – Toggle this help screen",
        ]

        if self.show_help:
            self._draw_help_panel(frame, help_lines)
        else:
            lines.append("Press H for help")
            y = 35
            for line in lines:
                cv2.putText(frame, line, (20, y), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 1, cv2.LINE_AA)
                y += 30

        if self.edit_mode:
            self.draw_edit_panel(frame)
        if self.penalty_mode:
            self.draw_penalty_panel(frame)

    def _draw_help_panel(self, frame, help_lines: List[str]) -> None:
        h, w = frame.shape[:2]
        panel = frame.copy()

        x0, y0 = 12, 12
        x1 = max(320, int(w * 0.78))
        y1 = h - 12

        cv2.rectangle(panel, (x0, y0), (x1, y1), (8, 8, 24), -1)
        cv2.rectangle(panel, (x0, y0), (x1, y1), (180, 200, 255), 1)
        cv2.addWeighted(panel, 0.88, frame, 0.12, 0, frame)

        cv2.putText(frame, "Help", (x0 + 16, y0 + 28), cv2.FONT_HERSHEY_SIMPLEX, 0.72, (255, 255, 255), 2, cv2.LINE_AA)
        cv2.putText(frame, "Press H to hide", (x1 - 150, y0 + 28), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (220, 220, 220), 1, cv2.LINE_AA)

        column_gap = 30
        inner_x0 = x0 + 18
        inner_y0 = y0 + 54
        inner_x1 = x1 - 18
        inner_y1 = y1 - 18
        col_width = (inner_x1 - inner_x0 - column_gap) // 2

        font = cv2.FONT_HERSHEY_SIMPLEX
        font_scale = 0.48
        line_h = 18

        columns = [[], []]
        current_col = 0
        current_lines = 0
        max_lines_per_col = max(8, (inner_y1 - inner_y0) // line_h)

        for line in help_lines:
            if current_lines >= max_lines_per_col and current_col == 0:
                current_col = 1
                current_lines = 0
            if current_lines >= max_lines_per_col:
                break
            columns[current_col].append(line)
            current_lines += 1

        for col_idx, col_lines in enumerate(columns):
            base_x = inner_x0 + col_idx * (col_width + column_gap)
            y = inner_y0
            for line in col_lines:
                color = (255, 255, 255) if line.endswith(":") else (225, 225, 225)
                weight = 2 if line.endswith(":") else 1
                cv2.putText(frame, line, (base_x, y), font, font_scale, color, weight, cv2.LINE_AA)
                y += line_h

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

    def draw_penalty_panel(self, frame) -> None:
        h, w = frame.shape[:2]
        panel = frame.copy()
        width = min(900, w - 80)
        height = 290
        x0 = (w - width) // 2
        y0 = max(20, h - height - 20)
        x1 = x0 + width
        y1 = y0 + height

        cv2.rectangle(panel, (x0, y0), (x1, y1), (20, 20, 20), -1)
        cv2.rectangle(panel, (x0, y0), (x1, y1), (255, 255, 255), 1)
        cv2.addWeighted(panel, 0.9, frame, 0.1, 0, frame)

        cv2.putText(frame, "Add Penalty", (x0 + 20, y0 + 35), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2, cv2.LINE_AA)
        cv2.putText(frame, "Tab cycles fields: skater -> code -> team. Enter saves. Esc cancels.", (x0 + 20, y0 + 65), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (220, 220, 220), 1, cv2.LINE_AA)

        skater_color = (180, 220, 255) if self.penalty_focus == "skater" else (255, 255, 255)
        code_color = (180, 220, 255) if self.penalty_focus == "code" else (255, 255, 255)
        team_color = (180, 220, 255) if self.penalty_focus == "team" else (255, 255, 255)

        cv2.putText(frame, f"Skater: {self.penalty_skater or '-'}", (x0 + 20, y0 + 110), cv2.FONT_HERSHEY_SIMPLEX, 0.7, skater_color, 1, cv2.LINE_AA)
        cv2.putText(frame, f"Code: {self.penalty_code or '-'}", (x0 + 260, y0 + 110), cv2.FONT_HERSHEY_SIMPLEX, 0.7, code_color, 1, cv2.LINE_AA)
        cv2.putText(frame, f"Team: {self.penalty_team.upper()}", (x0 + 460, y0 + 110), cv2.FONT_HERSHEY_SIMPLEX, 0.7, team_color, 1, cv2.LINE_AA)

        cv2.putText(frame, "Official penalty codes:", (x0 + 20, y0 + 150), cv2.FONT_HERSHEY_SIMPLEX, 0.56, (220, 220, 220), 1, cv2.LINE_AA)
        code_lines = [
            "A High Block   B Back Block   C Illegal Contact   D Direction",
            "E Leg Block    F Forearms    G Misconduct        H Blocking with the Head",
            "I Illegal Procedure   L Low Block   M Multiplayer Block   N Interference",
            "P Illegal Position   X Cut",
        ]
        cy = y0 + 180
        for line in code_lines:
            cv2.putText(frame, line, (x0 + 20, cy), cv2.FONT_HERSHEY_SIMPLEX, 0.48, (200, 200, 200), 1, cv2.LINE_AA)
            cy += 24

    def draw_edit_panel(self, frame) -> None:
        jam = self.ensure_current_jam()
        panel = frame.copy()
        h, w = frame.shape[:2]

        # Match the custom layout: video occupies top-right ~80% width.
        video_x0 = int(w * 0.2)
        video_w = w - video_x0
        video_h = int(video_w * 9 / 16) if video_w > 0 else 0
        if self.current_frame is not None:
            fh, fw = self.current_frame.shape[:2]
            if fw > 0:
                video_h = int(video_w * (fh / fw))
        video_h = min(video_h, h)

        # Put editor in the bottom area under the video, spanning most of the width.
        x0 = 20
        y0 = min(h - 40, video_h + 10)
        width = w - 40
        height = h - y0 - 20
        x1 = x0 + width
        y1 = y0 + height

        cv2.rectangle(panel, (x0, y0), (x1, y1), (20, 20, 20), -1)
        cv2.rectangle(panel, (x0, y0), (x1, y1), (255, 255, 255), 1)
        cv2.addWeighted(panel, 0.88, frame, 0.12, 0, frame)

        cv2.putText(frame, f"Edit Jam {jam.jam_number}", (x0 + 20, y0 + 30), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2, cv2.LINE_AA)
        cv2.putText(frame, "Tab/Enter next | Up/Down move | ` toggles booleans | Backspace deletes | W saves | Esc cancels", (x0 + 20, y0 + 55), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (220, 220, 220), 1, cv2.LINE_AA)

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
            "lead_jammer": "Lead jammer (home/away/none/unknown)",
            "notes": "Notes",
        }

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
            "home_score_start": self._fmt_int(jam.home_score_start),
            "away_score_start": self._fmt_int(jam.away_score_start),
            "home_passes": ", ".join(str(p) for p in jam.home_passes),
            "away_passes": ", ".join(str(p) for p in jam.away_passes),
            "home_score_end": self._fmt_int(jam.home_score_end),
            "away_score_end": self._fmt_int(jam.away_score_end),
            "home_star_pass": "Y" if jam.home_star_pass else "N",
            "away_star_pass": "Y" if jam.away_star_pass else "N",
            "lead_jammer": jam.lead_jammer,
            "notes": jam.notes,
        }

        row_y = y0 + 85
        row_h = 26
        label_x = x0 + 25
        value_x = x0 + 260
        suggestions_x = min(x0 + 620, x1 - 260)

        visible_rows = max(1, (height - 100) // row_h)
        scroll_start = max(0, min(self.edit_field_index - visible_rows // 2, len(TEXT_INPUT_FIELDS) - visible_rows))
        scroll_end = min(len(TEXT_INPUT_FIELDS), scroll_start + visible_rows)

        for idx in range(scroll_start, scroll_end):
            field_name = TEXT_INPUT_FIELDS[idx]
            is_active = idx == self.edit_field_index
            label = field_labels[field_name]
            display_value = self.edit_buffer if is_active else (values[field_name] if values[field_name] not in {None, ''} else '-')
            if is_active:
                cv2.rectangle(frame, (x0 + 15, row_y - 18), (x1 - 15, row_y + 8), (70, 70, 70), -1)
                cv2.rectangle(frame, (x0 + 15, row_y - 18), (x1 - 15, row_y + 8), (255, 255, 255), 1)
            cv2.putText(frame, label, (label_x, row_y), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1, cv2.LINE_AA)
            value_text = f"> {display_value}" if is_active else str(display_value)
            cv2.putText(frame, value_text, (value_x, row_y), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1, cv2.LINE_AA)
            if is_active:
                text_size = cv2.getTextSize(value_text, cv2.FONT_HERSHEY_SIMPLEX, 0.6, 1)[0]
                cursor_x = min(suggestions_x - 20, value_x + text_size[0] + 3)
                cv2.line(frame, (cursor_x, row_y - 16), (cursor_x, row_y + 4), (255, 255, 255), 1)

                previous_values = self._collect_previous_values(field_name)
                cv2.putText(frame, "Previously used:", (suggestions_x, row_y), cv2.FONT_HERSHEY_SIMPLEX, 0.52, (180, 220, 255), 1, cv2.LINE_AA)
                prev_text = " | ".join(previous_values) if previous_values else "-"
                cv2.putText(frame, prev_text, (suggestions_x, row_y + 20), cv2.FONT_HERSHEY_SIMPLEX, 0.48, (220, 220, 220), 1, cv2.LINE_AA)
            row_y += row_h

        if len(TEXT_INPUT_FIELDS) > visible_rows:
            bar_x = x1 - 10
            bar_top = y0 + 80
            bar_bottom = y1 - 15
            cv2.line(frame, (bar_x, bar_top), (bar_x, bar_bottom), (100, 100, 100), 2)
            thumb_h = max(20, int((visible_rows / len(TEXT_INPUT_FIELDS)) * (bar_bottom - bar_top)))
            max_offset = max(1, len(TEXT_INPUT_FIELDS) - visible_rows)
            thumb_y = bar_top + int((scroll_start / max_offset) * max(1, (bar_bottom - bar_top - thumb_h)))
            cv2.rectangle(frame, (bar_x - 4, thumb_y), (bar_x + 4, thumb_y + thumb_h), (220, 220, 220), -1)

    @staticmethod
    def _fmt_time(value: Optional[float]) -> str:
        return "-" if value is None else f"{value:.3f}s"

    @staticmethod
    def _fmt_int(value: Optional[int]) -> str:
        return "-" if value is None else str(value)

    @staticmethod
    def _fmt_lineup(value: List[str]) -> str:
        return "-" if not value else ", ".join(value)

    def adjust_playback_speed(self, direction: int) -> None:
        try:
            current_index = PLAYBACK_SPEED_STEPS.index(self.playback_speed)
        except ValueError:
            current_index = PLAYBACK_SPEED_STEPS.index(1.0)
        new_index = max(0, min(len(PLAYBACK_SPEED_STEPS) - 1, current_index + direction))
        self.playback_speed = PLAYBACK_SPEED_STEPS[new_index]
        self._sync_audio_seek()
        print(f"Playback speed: {self.playback_speed:.2f}x")

    def reset_playback_speed(self) -> None:
        self.playback_speed = 1.0
        self._sync_audio_seek()
        print("Playback speed: 1.00x")

    def read_next_frame_if_playing(self) -> None:
        if not self.playing:
            return

        frames_to_advance = max(1, int(round(self.playback_speed))) if self.playback_speed >= 1.0 else 1
        frame = None
        ok = False
        for _ in range(frames_to_advance):
            ok, frame = self.cap.read()
            if not ok:
                break

        if ok and frame is not None:
            self.current_frame_index = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES)) - 1
            self.current_frame = frame
        else:
            self.playing = False

    def run(self) -> None:
        self.seek_to_frame(0)
        cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)

        while True:
            # Detect if window was closed via the window manager (X button)
            if cv2.getWindowProperty(WINDOW_NAME, cv2.WND_PROP_VISIBLE) < 1:
                break

            self.read_next_frame_if_playing()

            if self.current_frame is None:
                self.current_frame = self._blank_frame("No frame available")

            frame = self.current_frame
            fh, fw = frame.shape[:2]

            # Create midnight blue canvas
            canvas_h = fh
            canvas_w = fw
            canvas = np.full((canvas_h, canvas_w, 3), (112, 25, 25), dtype=np.uint8)  # BGR midnight blue

            # Target video area (top-right 80% width)
            target_w = int(canvas_w * 0.8)
            scale = target_w / fw
            target_h = int(fh * scale)

            if target_h > canvas_h:
                scale = canvas_h / fh
                target_h = canvas_h
                target_w = int(fw * scale)

            resized = cv2.resize(frame, (target_w, target_h))

            x_offset = canvas_w - target_w
            y_offset = 0

            canvas[y_offset:y_offset+target_h, x_offset:x_offset+target_w] = resized

            display = canvas

            self.draw_overlay(display)
            cv2.imshow(WINDOW_NAME, display)

            if self.playing:
                delay_ms = max(1, int((1000 / self.fps) / min(self.playback_speed, 1.0)))
            else:
                delay_ms = 30
            key = cv2.waitKey(delay_ms) & 0xFF

            if self.edit_mode:
                if key == ord("w"):
                    self.end_edit_mode(save_changes=True)
                else:
                    self._handle_edit_key(key)
                continue

            if self.penalty_mode:
                self._handle_penalty_key(key)
                continue

            if key == 255:
                continue

            if key == ord(" "):
                self.playing = not self.playing
                self._sync_audio_play_state()
            elif key == ord("Q"):
                self._append_pass("home", 0)
            elif key == ord("W"):
                self._append_pass("home", 1)
            elif key == ord("E"):
                self._append_pass("home", 2)
            elif key == ord("R"):
                self._append_pass("home", 3)
            elif key == ord("T"):
                self._append_pass("home", 4)
            elif key == ord("A"):
                self._append_pass("away", 0)
            elif key == ord("S"):
                self._append_pass("away", 1)
            elif key == ord("D"):
                self._append_pass("away", 2)
            elif key == ord("F"):
                self._append_pass("away", 3)
            elif key == ord("G"):
                self._append_pass("away", 4)
            elif key == ord("z"):
                self._undo_last_pass("home")
            elif key == ord("x"):
                self._undo_last_pass("away")
            elif key == ord("1"):
                jam = self.ensure_current_jam()
                jam.lead_jammer = "home"
                self.save()
            elif key == ord("2"):
                jam = self.ensure_current_jam()
                jam.lead_jammer = "away"
                self.save()
            elif key == ord("3"):
                jam = self.ensure_current_jam()
                jam.lead_jammer = "none"
                self.save()
            elif key == ord("4"):
                jam = self.ensure_current_jam()
                jam.lead_jammer = "unknown"
                self.save()
            elif key == ord("5"):
                self.start_next_period()
            elif key == ord("-"):
                self.adjust_playback_speed(-1)
            elif key in (ord("="), ord("+")):
                self.adjust_playback_speed(1)
            elif key == ord("0"):
                self.reset_playback_speed()
            elif key == ord("h"):
                self.show_help = not self.show_help
            elif key == ord("q"):
                self.save()
                break
            elif key == ord("s"):
                self.mark_start()
            elif key == ord("e"):
                self.mark_end()
            elif key == ord("n"):
                self.next_jam()
            elif key == ord("p"):
                self.previous_jam()
            elif key == ord("u"):
                self.undo_last_boundary()
            elif key == ord("w"):
                self.save()
            elif key == ord("m"):
                self.begin_edit_mode()
            elif key == ord("b"):
                self.begin_penalty_mode()
            elif key == ord("d"):
                self.delete_current_jam_if_empty()
            elif key in (81, ord("a")):  # common left arrow code on some platforms
                self.playing = False
                self.step_seconds(-SEEK_SECONDS_SMALL)
            elif key in (83, ord("f")):  # common right arrow code on some platforms
                self.playing = False
                self.step_seconds(SEEK_SECONDS_SMALL)
            elif key == ord("j"):
                self.playing = False
                self.step_seconds(-SEEK_SECONDS_LARGE)
            elif key == ord("l"):
                self.playing = False
                self.step_seconds(SEEK_SECONDS_LARGE)
            elif key == ord("["):
                self.jump_to_previous_saved_jam_start()
            elif key == ord("]"):
                self.jump_to_next_saved_jam_start()
            elif key == 92:
                self.jump_to_current_jam_start()
            elif key == ord("{"):
                self.jump_to_previous_saved_jam_end()
            elif key == ord("}"):
                self.jump_to_next_saved_jam_end()
            elif key == ord(","):
                self.playing = False
                self.step_frames(-1)
            elif key == ord("."):
                self.playing = False
                self.step_frames(1)

        self.cap.release()
        if self.audio_enabled and self.audio_player is not None:
            try:
                self.audio_player.stop()
            except Exception:
                pass
        cv2.destroyAllWindows()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Annotate roller derby jams in a video.")
    parser.add_argument("video", nargs="?", default=None, help="Path to the input video file")
    parser.add_argument(
        "--output",
        default=None,
        help="Path to output JSON annotation file (default: same name as video with .jams.json)",
    )
    return parser.parse_args()


def default_output_path(video_path: str) -> str:
    video = Path(video_path)
    return str(video.with_suffix(".jams.json"))


def prompt_for_video_file() -> str:
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        root.update()
        selected = filedialog.askopenfilename(
            title="Select roller derby video",
            filetypes=[
                ("Video files", "*.mp4 *.webm *.mkv *.avi *.mov *.m4v"),
                ("All files", "*.*"),
            ],
        )
        root.destroy()
        if selected:
            return selected
    except Exception:
        pass

    while True:
        entered = input("Enter path to video file: ").strip().strip('"')
        if entered and os.path.exists(entered):
            return entered
        print("Video file not found. Please try again.")


def main() -> None:
    args = parse_args()
    video_path = args.video or prompt_for_video_file()
    output_path = args.output or default_output_path(video_path)
    annotator = JamAnnotator(video_path, output_path)
    annotator.run()


if __name__ == "__main__":
    main()
