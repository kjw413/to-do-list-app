# -*- coding: utf-8 -*-
"""
업무 To-Do List — 데스크톱 할 일 관리 프로그램
==================================================
- 우선순위(높음/중간/낮음), 마감일/남은기간(영업일)·긴급도, 카테고리 분류, 완료 체크
- 캘린더(휴가/출장/교육 일정) + 공휴일을 비근무일로 보고 남은기간을 재계산
- 표준 라이브러리(Tkinter)만 사용 → 별도 설치 없이 실행 가능
- 데이터는 %APPDATA%\\TodoList\\tasks.json (일정은 schedules.json) 에 자동 저장

실행:        py todo_app.py
자체점검:    py todo_app.py --selftest   (창을 띄우지 않고 UI 구성만 검증)
"""

import os
import sys
import json
import uuid
import shutil
try:
    import winreg  # Windows 자동 실행 등록용 (표준 라이브러리)
except ImportError:
    winreg = None
from dataclasses import dataclass, field, asdict, fields as dataclass_fields
from datetime import date, datetime, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont


# ============================================================
# 상수 / 설정
# ============================================================
APP_TITLE = "업무 To-Do List"
DATA_VERSION = 1

PRIORITIES = ["높음", "중간", "낮음"]
PRIORITY_WEIGHT = {"높음": 3, "중간": 2, "낮음": 1}

# 긴급도(영업일 기준) 행 배경색
COLOR_OVERDUE = "#ffe0e0"   # 1영업일 전/지연 (위험)
COLOR_TODAY   = "#ffe9b8"   # 2~4영업일 전 (임박)
COLOR_SOON    = "#fff7d6"   # 5~7영업일 전 (주의)
COLOR_DONE_FG = "#9aa0a6"   # 완료 항목 글자색

SORT_MODES = ["긴급도순", "마감일 빠른순", "우선순위순", "카테고리순", "등록순"]
FILTER_ALL = "전체"

# ----- 캘린더(일정) -----
SCHEDULE_KINDS = ["휴가", "출장", "교육", "기타"]
SCHEDULE_COLORS = {           # 달력 셀/목록에 표시할 일정 종류별 색
    "휴가": "#cfe3ff",
    "출장": "#ffe2c2",
    "교육": "#e0d3ff",
    "기타": "#dcdcdc",
}

# 달력 셀 배경색
COLOR_CAL_SAT     = "#eef3ff"   # 토요일
COLOR_CAL_SUN     = "#ffeef0"   # 일요일
COLOR_CAL_HOLIDAY = "#ffdede"   # 공휴일
COLOR_CAL_TODAY   = "#fff3c4"   # 오늘
COLOR_CAL_OTHER   = "#f3f3f3"   # 다른 달 날짜

# 대한민국 공휴일(관공서 공휴일) — 비근무일 계산에서 자동 제외한다.
# ※ 설날/추석 등 음력 명절과 대체공휴일은 해마다 바뀌므로 연도별로 확인·갱신해야 한다.
PUBLIC_HOLIDAYS_RAW = {
    2025: ["01-01", "01-27", "01-28", "01-29", "01-30", "03-01", "03-03",
           "05-05", "05-06", "06-06", "08-15", "10-03", "10-05", "10-06",
           "10-07", "10-08", "10-09", "12-25"],
    2026: ["01-01", "02-16", "02-17", "02-18", "03-01", "03-02", "05-05",
           "05-24", "05-25", "06-06", "08-15", "08-17", "09-24", "09-25",
           "09-26", "09-28", "10-03", "10-05", "10-09", "12-25"],
    2027: ["01-01", "02-06", "02-07", "02-08", "02-09", "03-01", "05-05",
           "05-13", "06-06", "08-15", "08-16", "09-14", "09-15", "09-16",
           "10-03", "10-04", "10-09", "10-11", "12-25", "12-27"],
}


# ============================================================
# 저장 위치
# ============================================================
def data_dir() -> Path:
    """데이터/설정을 저장할 폴더 (%APPDATA%\\TodoList). 없으면 생성."""
    base = os.environ.get("APPDATA") or str(Path.home())
    d = Path(base) / "TodoList"
    d.mkdir(parents=True, exist_ok=True)
    return d


def app_dir() -> Path:
    """프로그램 본체가 위치한 폴더. (exe로 빌드되면 exe가 있는 폴더)"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


DATA_FILE = data_dir() / "tasks.json"
SETTINGS_FILE = data_dir() / "settings.json"
SCHEDULE_FILE = data_dir() / "schedules.json"
HOLIDAY_FILE = app_dir() / "DB_holiday.xlsx"   # 공휴일 DB (A:날짜 B:요일 C:공휴일명)


# ============================================================
# 데이터 모델
# ============================================================
@dataclass
class Task:
    title: str
    category: str = ""
    priority: str = "중간"
    due: str = ""                 # "YYYY-MM-DD" 또는 ""(없음)
    note: str = ""
    done: bool = False
    id: str = field(default_factory=lambda: uuid.uuid4().hex[:8])
    created_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))
    done_at: str = ""


@dataclass
class ScheduleEvent:
    """캘린더 일정(휴가/출장/교육 등). non_working=True면 비근무일로 계산에서 제외."""
    title: str
    kind: str = "휴가"            # SCHEDULE_KINDS 중 하나
    start: str = ""               # "YYYY-MM-DD"
    end: str = ""                 # "YYYY-MM-DD" (비었으면 start와 동일 = 하루)
    non_working: bool = True      # to-do 남은기간 계산에서 비근무일로 간주할지
    note: str = ""
    id: str = field(default_factory=lambda: uuid.uuid4().hex[:8])
    created_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    def date_range(self):
        """이 일정이 포함하는 모든 날짜(date) 리스트. 형식 오류면 빈 리스트."""
        d0 = parse_due(self.start)
        if d0 is None:
            return []
        d1 = parse_due(self.end) or d0
        if d1 < d0:
            d0, d1 = d1, d0
        out, cur = [], d0
        while cur <= d1:
            out.append(cur)
            cur += timedelta(days=1)
        return out


# ------- 날짜/긴급도 유틸 -------
def parse_due(s: str):
    """'YYYY-MM-DD' 문자열을 date로. 비었거나 형식 오류면 None."""
    if not s:
        return None
    try:
        return datetime.strptime(s.strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def normalize_due_input(s: str, default_year: int = None):
    """마감일 입력값을 'YYYY-MM-DD'로 보정. 실패하면 None, 빈 값은 ''."""
    raw = (s or "").strip()
    if not raw:
        return ""

    d = parse_due(raw)
    if d is not None:
        return d.isoformat()

    if not raw.isdigit():
        return None

    if len(raw) == 4:
        year = default_year or date.today().year
        candidate = f"{year}-{raw[:2]}-{raw[2:]}"
    elif len(raw) == 8:
        candidate = f"{raw[:4]}-{raw[4:6]}-{raw[6:]}"
    else:
        return None

    d = parse_due(candidate)
    return d.isoformat() if d is not None else None


def days_left(s: str):
    """오늘 기준 남은 일수. 마감일 없으면 None. (음수 = 지연)"""
    d = parse_due(s)
    if d is None:
        return None
    return (d - date.today()).days


def _coerce_date(v):
    """엑셀 셀 값(datetime/date/문자열)을 date로 변환. 실패 시 None."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def _builtin_holidays() -> dict:
    """엑셀을 못 읽을 때 쓰는 내장 공휴일(이름은 '공휴일'로 통일)."""
    out = {}
    for year, days in PUBLIC_HOLIDAYS_RAW.items():
        for mmdd in days:
            d = parse_due(f"{year}-{mmdd}")
            if d:
                out[d] = "공휴일"
    return out


def load_public_holidays(path: Path = None) -> dict:
    """공휴일 {date: 이름} 사전을 만든다.

    우선 DB_holiday.xlsx(A:날짜, C:공휴일명)를 읽고,
    파일이 없거나 openpyxl 미설치/읽기 실패 시 내장 목록으로 폴백한다.
    """
    path = Path(path) if path else HOLIDAY_FILE
    if path.exists():
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            out = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                d = _coerce_date(row[0])
                if d is None:
                    continue
                name = row[2] if len(row) > 2 and row[2] else "공휴일"
                out[d] = str(name).strip()
            wb.close()
            if out:
                return out
        except Exception:
            pass
    return _builtin_holidays()


PUBLIC_HOLIDAYS = load_public_holidays()


def reload_public_holidays():
    """DB_holiday.xlsx를 다시 읽어 PUBLIC_HOLIDAYS를 갱신한다."""
    global PUBLIC_HOLIDAYS
    PUBLIC_HOLIDAYS = load_public_holidays()
    return PUBLIC_HOLIDAYS


class WorkCalendar:
    """비근무일(주말 + 공휴일 + 등록된 휴가/출장/교육 일정)을 판정한다.

    영업일 계산 함수들이 이 전역 인스턴스(WORK_CALENDAR)를 참조하므로,
    일정이 추가/변경되면 set_events()로 비근무일 집합을 갱신해야 한다.
    """

    def __init__(self):
        self._event_days: set = set()   # 일정에서 비롯된 비근무일

    def set_events(self, events):
        """ScheduleEvent 목록으로부터 비근무일 집합을 다시 만든다."""
        days = set()
        for ev in events:
            if getattr(ev, "non_working", True):
                days.update(ev.date_range())
        self._event_days = days

    def is_non_working(self, d: date) -> bool:
        return (d.weekday() >= 5) or (d in PUBLIC_HOLIDAYS) or (d in self._event_days)

    def is_working_day(self, d: date) -> bool:
        return not self.is_non_working(d)


WORK_CALENDAR = WorkCalendar()


def business_days_left(s: str):
    """오늘 기준 마감일까지 남은 영업일수.

    주말·공휴일과 등록된 휴가/출장/교육 일정을 비근무일로 보고 제외한다
    (WORK_CALENDAR 참조).
    """
    d = parse_due(s)
    if d is None:
        return None
    today = date.today()
    if d == today:
        return 0

    step = 1 if d > today else -1
    n = 0
    cur = today
    while cur != d:
        cur += timedelta(days=step)
        if WORK_CALENDAR.is_working_day(cur):
            n += step

    if d < today and n == 0:
        return -1
    return n


def dday_text(task: "Task") -> str:
    """남은기간 표시 문자열(영업일 기준)."""
    if task.done:
        return "완료"
    n = business_days_left(task.due)
    if n is None:
        return "-"
    if n < 0:
        return f"지연 {abs(n)}일"
    return f"{n}일"


def urgency_score(task: "Task") -> int:
    """긴급도 점수(높을수록 급함). 마감 임박/지연 + 우선순위를 결합."""
    p = PRIORITY_WEIGHT.get(task.priority, 2)
    n = business_days_left(task.due)
    if n is None:
        time_score = 30           # 마감일 없음 → 중간 정도
    elif n < 0:
        overdue_days = abs(days_left(task.due) or n)
        time_score = 220 + min(overdue_days, 60)   # 지연 = 최상위(많이 지날수록 ↑)
    elif n <= 1:
        time_score = 180
    elif n <= 4:
        time_score = 120
    elif n <= 7:
        time_score = 80
    else:
        time_score = 40
    return time_score + p * 5


def sort_key(mode: str):
    """정렬 모드별 key 함수. 공통적으로 '미완료 먼저, 완료는 뒤'."""
    def big_if_none(n):
        return n if n is not None else 99999

    if mode == "마감일 빠른순":
        return lambda t: (t.done, days_left(t.due) is None, big_if_none(days_left(t.due)), -PRIORITY_WEIGHT[t.priority])
    if mode == "우선순위순":
        return lambda t: (t.done, -PRIORITY_WEIGHT[t.priority], big_if_none(days_left(t.due)))
    if mode == "카테고리순":
        return lambda t: (t.done, (t.category or "힣"), -urgency_score(t))
    if mode == "등록순":
        return lambda t: (t.done, t.created_at)
    # 기본: 긴급도순
    return lambda t: (t.done, -urgency_score(t), t.created_at)


# ============================================================
# 저장소 (JSON, 원자적 저장 + 손상 복구)
# ============================================================
class Store:
    def __init__(self, path: Path = DATA_FILE):
        self.path = Path(path)
        self.tasks: list[Task] = []
        self.load()

    def load(self):
        if not self.path.exists():
            self.tasks = []
            return
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            items = raw.get("tasks", []) if isinstance(raw, dict) else raw
            valid_keys = {f.name for f in dataclass_fields(Task)}
            self.tasks = [Task(**{k: v for k, v in it.items() if k in valid_keys}) for it in items]
        except Exception:
            # 손상된 파일은 백업하고 빈 목록으로 시작 (데이터 유실 방지)
            try:
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                backup = self.path.parent / f"{self.path.stem}.corrupt-{stamp}.json"
                shutil.copy2(self.path, backup)
            except Exception:
                pass
            self.tasks = []

    def save(self):
        payload = {"version": DATA_VERSION, "tasks": [asdict(t) for t in self.tasks]}
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.path)   # 원자적 교체 (저장 중 종료돼도 안전)

    # 편의 메서드
    def add(self, task: Task):
        self.tasks.append(task)
        self.save()

    def remove(self, task_id: str):
        self.tasks = [t for t in self.tasks if t.id != task_id]
        self.save()

    def by_id(self, task_id: str):
        for t in self.tasks:
            if t.id == task_id:
                return t
        return None

    def categories(self):
        return sorted({t.category for t in self.tasks if t.category})


# ============================================================
# 일정(캘린더) 저장소
# ============================================================
class ScheduleStore:
    """휴가/출장/교육 등 일정(ScheduleEvent) 저장소. tasks.json과 같은 방식."""

    def __init__(self, path: Path = SCHEDULE_FILE):
        self.path = Path(path)
        self.events: list[ScheduleEvent] = []
        self.load()

    def load(self):
        if not self.path.exists():
            self.events = []
            return
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            items = raw.get("events", []) if isinstance(raw, dict) else raw
            valid_keys = {f.name for f in dataclass_fields(ScheduleEvent)}
            self.events = [ScheduleEvent(**{k: v for k, v in it.items() if k in valid_keys})
                           for it in items]
        except Exception:
            try:
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                backup = self.path.parent / f"{self.path.stem}.corrupt-{stamp}.json"
                shutil.copy2(self.path, backup)
            except Exception:
                pass
            self.events = []

    def save(self):
        payload = {"version": DATA_VERSION, "events": [asdict(e) for e in self.events]}
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.path)

    def add(self, ev: ScheduleEvent):
        self.events.append(ev)
        self.save()

    def remove(self, ev_id: str):
        self.events = [e for e in self.events if e.id != ev_id]
        self.save()

    def by_id(self, ev_id: str):
        for e in self.events:
            if e.id == ev_id:
                return e
        return None

    def events_on(self, d: date):
        """특정 날짜에 걸치는 일정 목록."""
        return [e for e in self.events if d in e.date_range()]

    def events_in_month(self, year: int, month: int):
        """해당 월에 하루라도 걸치는 일정을, 시작일 순으로."""
        out = []
        for e in self.events:
            for d in e.date_range():
                if d.year == year and d.month == month:
                    out.append(e)
                    break
        out.sort(key=lambda e: (e.start or "9999", e.kind))
        return out


# ============================================================
# 설정(창 크기/정렬/필터) 저장·복원
# ============================================================
def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(d: dict):
    try:
        SETTINGS_FILE.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


# ============================================================
# Windows 시작 시 자동 실행 (레지스트리 HKCU\...\Run, 관리자 권한 불필요)
# ============================================================
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_NAME = "UpmuTodoList"


def _autostart_command() -> str:
    """로그인 시 실행할 명령. exe로 빌드되면 exe 경로, 소스면 pythonw+스크립트."""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    pyw = Path(sys.executable).with_name("pythonw.exe")
    runner = str(pyw) if pyw.exists() else sys.executable
    return f'"{runner}" "{Path(__file__).resolve()}"'


def is_autostart_enabled() -> bool:
    if winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, AUTOSTART_NAME)
            return True
    except OSError:          # 키/값 없음 등
        return False


def set_autostart(enable: bool):
    if winreg is None:
        raise RuntimeError("이 OS에서는 자동 실행 등록을 지원하지 않습니다.")
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
        if enable:
            winreg.SetValueEx(k, AUTOSTART_NAME, 0, winreg.REG_SZ, _autostart_command())
        else:
            try:
                winreg.DeleteValue(k, AUTOSTART_NAME)
            except FileNotFoundError:
                pass


# ============================================================
# 업무 추가/수정 다이얼로그
# ============================================================
class TaskDialog(tk.Toplevel):
    """업무 추가/수정용 모달 창. 저장 시 self.result 에 Task를 담는다."""

    def __init__(self, parent, task: Task = None, categories=None):
        super().__init__(parent)
        self.result = None
        self.editing = task is not None
        self.task = task or Task(title="")
        categories = categories or []

        self.title("업무 수정" if self.editing else "새 업무 추가")
        self.transient(parent)
        self.resizable(False, False)
        self.configure(padx=16, pady=14)

        pad = {"padx": 6, "pady": 6}

        # 업무 내용
        ttk.Label(self, text="업무 내용 *").grid(row=0, column=0, sticky="w", **pad)
        self.var_title = tk.StringVar(value=self.task.title)
        e_title = ttk.Entry(self, textvariable=self.var_title, width=44)
        e_title.grid(row=0, column=1, columnspan=3, sticky="we", **pad)

        # 카테고리
        ttk.Label(self, text="카테고리").grid(row=1, column=0, sticky="w", **pad)
        self.var_cat = tk.StringVar(value=self.task.category)
        self.cb_cat = ttk.Combobox(self, textvariable=self.var_cat, values=categories, width=20)
        self.cb_cat.grid(row=1, column=1, sticky="we", **pad)

        # 우선순위
        ttk.Label(self, text="우선순위").grid(row=1, column=2, sticky="e", **pad)
        self.var_pri = tk.StringVar(value=self.task.priority)
        cb_pri = ttk.Combobox(self, textvariable=self.var_pri, values=PRIORITIES,
                              width=8, state="readonly")
        cb_pri.grid(row=1, column=3, sticky="w", **pad)

        # 마감일 + 빠른 설정 버튼
        ttk.Label(self, text="마감일").grid(row=2, column=0, sticky="w", **pad)
        self.var_due = tk.StringVar(value=self.task.due)
        e_due = ttk.Entry(self, textvariable=self.var_due, width=14)
        e_due.grid(row=2, column=1, sticky="w", **pad)
        ttk.Label(self, text="(YYYY-MM-DD)", foreground="#888").grid(row=2, column=2, sticky="w")

        quick = ttk.Frame(self)
        quick.grid(row=2, column=3, sticky="w", **pad)
        ttk.Button(quick, text="오늘", width=5,
                   command=lambda: self._set_due(0, from_current=False)).pack(side="left", padx=1)
        ttk.Button(quick, text="내일", width=5,
                   command=lambda: self._set_due(1)).pack(side="left", padx=1)
        ttk.Button(quick, text="+1주", width=5,
                   command=lambda: self._set_due(7)).pack(side="left", padx=1)
        ttk.Button(quick, text="지움", width=5,
                   command=lambda: self.var_due.set("")).pack(side="left", padx=1)

        # 메모
        ttk.Label(self, text="메모").grid(row=3, column=0, sticky="nw", **pad)
        self.txt_note = tk.Text(self, width=44, height=4)
        self.txt_note.grid(row=3, column=1, columnspan=3, sticky="we", **pad)
        self.txt_note.insert("1.0", self.task.note)

        # 버튼
        btns = ttk.Frame(self)
        btns.grid(row=4, column=0, columnspan=4, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="저장", command=self._on_save).pack(side="right", padx=4)
        ttk.Button(btns, text="취소", command=self._on_cancel).pack(side="right")

        self.columnconfigure(1, weight=1)

        # 단축키/포커스
        self.bind("<Return>", lambda e: self._on_save())
        self.bind("<Escape>", lambda e: self._on_cancel())
        e_title.focus_set()
        e_title.icursor("end")

        # 모달 처리
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._center(parent)
        self.grab_set()

    def _center(self, parent):
        self.update_idletasks()
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            w, h = self.winfo_width(), self.winfo_height()
            self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 3}")
        except Exception:
            pass

    def _set_due(self, offset_days: int, from_current: bool = True):
        base = date.today()
        if from_current:
            due = normalize_due_input(self.var_due.get())
            d = parse_due(due) if due else None
            if d is not None:
                base = d
        self.var_due.set((base + timedelta(days=offset_days)).isoformat())

    def _on_save(self):
        title = self.var_title.get().strip()
        if not title:
            messagebox.showwarning("입력 필요", "업무 내용을 입력하세요.", parent=self)
            return
        due = normalize_due_input(self.var_due.get())
        if due is None:
            messagebox.showwarning("날짜 형식 오류",
                                   "마감일은 YYYY-MM-DD, MMDD, YYYYMMDD 형식으로 입력하세요.\n예: 2026-06-30, 0630, 20260630",
                                   parent=self)
            return
        self.var_due.set(due)
        # 기존 객체에 반영 (id/created_at 유지)
        self.task.title = title
        self.task.category = self.var_cat.get().strip()
        self.task.priority = self.var_pri.get() if self.var_pri.get() in PRIORITIES else "중간"
        self.task.due = due
        self.task.note = self.txt_note.get("1.0", "end").strip()
        self.result = self.task
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ============================================================
# 일정 추가/수정 다이얼로그
# ============================================================
class ScheduleDialog(tk.Toplevel):
    """휴가/출장/교육 등 일정 추가·수정 모달. 저장 시 self.result 에 ScheduleEvent."""

    def __init__(self, parent, event: ScheduleEvent = None, default_date: date = None):
        super().__init__(parent)
        self.result = None
        self.editing = event is not None
        self.event = event or ScheduleEvent(title="")
        if not self.editing and default_date is not None:
            self.event.start = default_date.isoformat()
            self.event.end = default_date.isoformat()

        self.title("일정 수정" if self.editing else "새 일정 추가")
        self.transient(parent)
        self.resizable(False, False)
        self.configure(padx=16, pady=14)
        pad = {"padx": 6, "pady": 6}

        # 종류
        ttk.Label(self, text="종류").grid(row=0, column=0, sticky="w", **pad)
        self.var_kind = tk.StringVar(value=self.event.kind)
        cb_kind = ttk.Combobox(self, textvariable=self.var_kind, values=SCHEDULE_KINDS,
                               width=10, state="readonly")
        cb_kind.grid(row=0, column=1, sticky="w", **pad)

        # 비근무일 처리
        self.var_nw = tk.BooleanVar(value=self.event.non_working)
        ttk.Checkbutton(self, text="비근무일로 처리(남은기간 계산에서 제외)",
                        variable=self.var_nw).grid(row=0, column=2, columnspan=2,
                                                    sticky="w", **pad)

        # 제목
        ttk.Label(self, text="제목 *").grid(row=1, column=0, sticky="w", **pad)
        self.var_title = tk.StringVar(value=self.event.title)
        e_title = ttk.Entry(self, textvariable=self.var_title, width=40)
        e_title.grid(row=1, column=1, columnspan=3, sticky="we", **pad)

        # 시작일 / 종료일
        ttk.Label(self, text="시작일").grid(row=2, column=0, sticky="w", **pad)
        self.var_start = tk.StringVar(value=self.event.start)
        ttk.Entry(self, textvariable=self.var_start, width=14).grid(row=2, column=1, sticky="w", **pad)
        ttk.Label(self, text="종료일").grid(row=2, column=2, sticky="e", **pad)
        self.var_end = tk.StringVar(value=self.event.end)
        ttk.Entry(self, textvariable=self.var_end, width=14).grid(row=2, column=3, sticky="w", **pad)
        ttk.Label(self, text="(YYYY-MM-DD · 종료일을 비우면 하루 일정)", foreground="#888")\
            .grid(row=3, column=1, columnspan=3, sticky="w", padx=6)

        # 메모
        ttk.Label(self, text="메모").grid(row=4, column=0, sticky="nw", **pad)
        self.txt_note = tk.Text(self, width=40, height=3)
        self.txt_note.grid(row=4, column=1, columnspan=3, sticky="we", **pad)
        self.txt_note.insert("1.0", self.event.note)

        # 버튼
        btns = ttk.Frame(self)
        btns.grid(row=5, column=0, columnspan=4, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="저장", command=self._on_save).pack(side="right", padx=4)
        ttk.Button(btns, text="취소", command=self._on_cancel).pack(side="right")

        self.columnconfigure(1, weight=1)
        self.bind("<Return>", lambda e: self._on_save())
        self.bind("<Escape>", lambda e: self._on_cancel())
        e_title.focus_set()
        e_title.icursor("end")

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._center(parent)
        self.grab_set()

    def _center(self, parent):
        self.update_idletasks()
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            w, h = self.winfo_width(), self.winfo_height()
            self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 3}")
        except Exception:
            pass

    def _on_save(self):
        title = self.var_title.get().strip()
        if not title:
            messagebox.showwarning("입력 필요", "일정 제목을 입력하세요.", parent=self)
            return
        start = self.var_start.get().strip()
        end = self.var_end.get().strip()
        if parse_due(start) is None:
            messagebox.showwarning("날짜 형식 오류",
                                   "시작일은 YYYY-MM-DD 형식이어야 합니다.\n예: 2026-07-20",
                                   parent=self)
            return
        if end and parse_due(end) is None:
            messagebox.showwarning("날짜 형식 오류",
                                   "종료일은 YYYY-MM-DD 형식이어야 합니다.", parent=self)
            return
        if end and parse_due(end) < parse_due(start):
            messagebox.showwarning("날짜 확인", "종료일이 시작일보다 빠릅니다.", parent=self)
            return

        self.event.kind = self.var_kind.get() if self.var_kind.get() in SCHEDULE_KINDS else "기타"
        self.event.non_working = self.var_nw.get()
        self.event.title = title
        self.event.start = start
        self.event.end = end
        self.event.note = self.txt_note.get("1.0", "end").strip()
        self.result = self.event
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ============================================================
# 월간 캘린더 창
# ============================================================
class CalendarWindow(tk.Toplevel):
    """월간 달력 그리드. 주말·공휴일·일정을 색으로 보여주고, 날짜 클릭으로 일정 추가.

    일정이 바뀌면 app.on_schedules_changed() 를 호출해 본문 목록의
    남은기간을 다시 계산하게 한다.
    """

    WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"]
    DEFAULT_WIDTH = 1200
    DEFAULT_HEIGHT = 800

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.sched: ScheduleStore = app.sched
        today = date.today()
        self.year, self.month = today.year, today.month

        self.title("업무 캘린더 · 일정 관리")
        self.transient(app)
        self.minsize(900, 680)
        self.geometry(f"{self.DEFAULT_WIDTH}x{self.DEFAULT_HEIGHT}")
        self.configure(padx=10, pady=8)

        self._build_header()
        self._build_grid()
        self._build_event_list()

        self.bind("<Escape>", lambda e: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._render()
        self._center(app)
        self.grab_set()

    def _center(self, parent=None):
        self.update_idletasks()
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            w = max(self.winfo_width(), self.DEFAULT_WIDTH)
            h = max(self.winfo_height(), self.DEFAULT_HEIGHT)
            x = max((sw - w) // 2, 0)
            y = max((sh - h) // 2, 0)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    # ----- 상단: 월 이동 -----
    def _build_header(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="◀ 이전달", width=8,
                   command=lambda: self._shift_month(-1)).pack(side="left")
        ttk.Button(bar, text="오늘", width=6, command=self._goto_today).pack(side="left", padx=4)
        ttk.Button(bar, text="다음달 ▶", width=8,
                   command=lambda: self._shift_month(1)).pack(side="left")
        self.lbl_month = ttk.Label(bar, text="", font=("맑은 고딕", 13, "bold"))
        self.lbl_month.pack(side="left", padx=16)
        ttk.Button(bar, text="＋ 일정 추가",
                   command=lambda: self._add_event()).pack(side="right")
        ttk.Label(bar, text="범례:  □휴가 □출장 □교육  · 빨강=공휴일/일요일",
                  foreground="#888").pack(side="right", padx=10)

    # ----- 가운데: 달력 그리드 -----
    def _build_grid(self):
        self.grid_frame = ttk.Frame(self)
        self.grid_frame.pack(fill="both", expand=True)
        for c in range(7):
            self.grid_frame.columnconfigure(c, weight=1, uniform="day")
        # 요일 헤더
        for c, wd in enumerate(self.WEEKDAYS):
            fg = "#c0392b" if c == 0 else ("#2257c4" if c == 6 else "#333")
            ttk.Label(self.grid_frame, text=wd, anchor="center",
                      foreground=fg, font=("맑은 고딕", 10, "bold"))\
                .grid(row=0, column=c, sticky="nsew", padx=1, pady=(0, 2))
        # 날짜 셀(6주 × 7일) — tk.Frame + Label 로 색 입힘
        self.cells = []
        for r in range(1, 7):
            self.grid_frame.rowconfigure(r, weight=1, uniform="week")
            row_cells = []
            for c in range(7):
                cell = tk.Frame(self.grid_frame, bd=1, relief="solid",
                                background="white", highlightthickness=0)
                cell.grid(row=r, column=c, sticky="nsew", padx=1, pady=1)
                daynum = tk.Label(cell, text="", anchor="nw", background="white",
                                  font=("맑은 고딕", 10, "bold"))
                daynum.pack(anchor="nw", fill="x", padx=3, pady=(2, 0))
                body = tk.Label(cell, text="", anchor="nw", justify="left",
                                background="white", font=("맑은 고딕", 8),
                                wraplength=96)
                body.pack(anchor="nw", fill="both", expand=True, padx=3)
                row_cells.append({"frame": cell, "daynum": daynum, "body": body, "date": None})
            self.cells.append(row_cells)

    # ----- 아래: 이달 일정 목록 -----
    def _build_event_list(self):
        wrap = ttk.LabelFrame(self, text="이번 달 일정", padding=6)
        wrap.pack(fill="x", pady=(8, 0))
        cols = ("kind", "title", "start", "end", "days", "nw")
        self.evtree = ttk.Treeview(wrap, columns=cols, show="headings", height=5,
                                   selectmode="browse")
        for c, (text, w, anc) in {
            "kind":  ("종류", 60, "center"),
            "title": ("제목", 240, "w"),
            "start": ("시작", 100, "center"),
            "end":   ("종료", 100, "center"),
            "days":  ("일수", 50, "center"),
            "nw":    ("비근무", 60, "center"),
        }.items():
            self.evtree.heading(c, text=text)
            self.evtree.column(c, width=w, anchor=anc, stretch=(c == "title"))
        self.evtree.pack(side="left", fill="x", expand=True)
        self.evtree.bind("<Double-1>", lambda e: self._edit_selected_event())
        self.evtree.bind("<Delete>", lambda e: self._delete_selected_event())

        btns = ttk.Frame(wrap)
        btns.pack(side="left", fill="y", padx=(8, 0))
        ttk.Button(btns, text="수정", width=8,
                   command=self._edit_selected_event).pack(pady=2)
        ttk.Button(btns, text="삭제", width=8,
                   command=self._delete_selected_event).pack(pady=2)

    # ----- 렌더링 -----
    def _render(self):
        import calendar as _cal
        self.lbl_month.config(text=f"{self.year}년 {self.month}월")
        today = date.today()

        # 그 달 1일이 포함된 주(일요일 시작)부터 6주를 채운다
        first = date(self.year, self.month, 1)
        start = first - timedelta(days=(first.weekday() + 1) % 7)  # 직전 일요일
        for r in range(6):
            for c in range(7):
                cur = start + timedelta(days=r * 7 + c)
                cell = self.cells[r][c]
                cell["date"] = cur
                in_month = (cur.month == self.month and cur.year == self.year)
                evs = self.sched.events_on(cur)
                holiday = cur in PUBLIC_HOLIDAYS

                # 배경색 결정
                if cur == today:
                    bg = COLOR_CAL_TODAY
                elif not in_month:
                    bg = COLOR_CAL_OTHER
                elif holiday or cur.weekday() == 6:
                    bg = COLOR_CAL_HOLIDAY
                elif cur.weekday() == 5:
                    bg = COLOR_CAL_SAT
                else:
                    bg = "white"
                # 일정(비근무)이 있으면 종류 색을 우선
                if in_month and evs:
                    bg = SCHEDULE_COLORS.get(evs[0].kind, bg)

                fg = "#bbbbbb" if not in_month else (
                    "#c0392b" if (holiday or cur.weekday() == 6) else
                    ("#2257c4" if cur.weekday() == 5 else "#222"))

                cell["frame"].configure(background=bg)
                cell["daynum"].configure(text=str(cur.day), background=bg, foreground=fg)

                # 셀 본문: 공휴일명 + 일정 + 마감 업무 수
                lines = []
                if in_month:
                    for hol_name in self._holiday_names(cur):
                        lines.append(f"※{hol_name}" if hol_name else "※공휴일")
                    for ev in evs:
                        lines.append(f"[{ev.kind}] {ev.title}")
                    due_n = sum(1 for t in self.app.store.tasks
                                if not t.done and parse_due(t.due) == cur)
                    if due_n:
                        lines.append(f"📌마감 {due_n}건")
                cell["body"].configure(text="\n".join(lines[:4]), background=bg)

                # 클릭으로 일정 추가
                for w in (cell["frame"], cell["daynum"], cell["body"]):
                    w.unbind("<Button-1>")
                    w.bind("<Button-1>", lambda e, d=cur: self._add_event(d))

        self._refresh_event_list()

    def _holiday_names(self, d: date):
        """그 날짜의 공휴일명(없으면 빈 리스트)."""
        name = PUBLIC_HOLIDAYS.get(d)
        return [name] if name else []

    def _refresh_event_list(self):
        self.evtree.delete(*self.evtree.get_children())
        for ev in self.sched.events_in_month(self.year, self.month):
            days = len(ev.date_range())
            self.evtree.insert("", "end", iid=ev.id,
                               values=(ev.kind, ev.title, ev.start, ev.end or ev.start,
                                       f"{days}일", "예" if ev.non_working else "아니오"))

    # ----- 월 이동 -----
    def _shift_month(self, delta):
        m = self.month + delta
        y = self.year
        while m < 1:
            m += 12
            y -= 1
        while m > 12:
            m -= 12
            y += 1
        self.year, self.month = y, m
        self._render()

    def _goto_today(self):
        t = date.today()
        self.year, self.month = t.year, t.month
        self._render()

    # ----- 일정 CRUD -----
    def _add_event(self, default_date: date = None):
        dlg = ScheduleDialog(self, event=None, default_date=default_date)
        self.wait_window(dlg)
        if dlg.result:
            self.sched.add(dlg.result)
            self._after_change()

    def _selected_event(self):
        sel = self.evtree.selection()
        return self.sched.by_id(sel[0]) if sel else None

    def _edit_selected_event(self):
        ev = self._selected_event()
        if not ev:
            return
        dlg = ScheduleDialog(self, event=ev)
        self.wait_window(dlg)
        if dlg.result:
            self.sched.save()
            self._after_change()

    def _delete_selected_event(self):
        ev = self._selected_event()
        if not ev:
            return
        if messagebox.askyesno("삭제 확인", f"'{ev.title}'\n이 일정을 삭제할까요?", parent=self):
            self.sched.remove(ev.id)
            self._after_change()

    def _after_change(self):
        """일정 변경 후: 비근무일 재계산 → 본문 갱신 → 달력 다시 그림."""
        self.app.on_schedules_changed()
        self._render()


# ============================================================
# 메인 윈도우
# ============================================================
class TodoApp(tk.Tk):
    REFRESH_MS = 10 * 60 * 1000   # 주기적 갱신 간격(10분)

    def __init__(self):
        super().__init__()
        self.store = Store()
        self.sched = ScheduleStore()
        WORK_CALENDAR.set_events(self.sched.events)   # 비근무일 초기화
        self.settings = load_settings()

        self.title(APP_TITLE)
        self.minsize(600, 360)
        geom = self.settings.get("geometry")
        self.geometry(geom if geom else "960x600")

        self._setup_style()

        # 상태 변수
        self.sort_mode = tk.StringVar(value=self.settings.get("sort", SORT_MODES[0]))
        self.filter_cat = tk.StringVar(value=self.settings.get("filter_cat", FILTER_ALL))
        self.hide_done = tk.BooleanVar(value=self.settings.get("hide_done", False))
        self.search_var = tk.StringVar(value="")
        self.summary_var = tk.BooleanVar(value=self.settings.get("show_startup_summary", True))
        self.summary_var.trace_add(
            "write", lambda *a: self._save_pref("show_startup_summary", self.summary_var.get()))
        self.autostart_var = tk.BooleanVar(value=is_autostart_enabled())
        self.ontop_var = tk.BooleanVar(value=self.settings.get("always_on_top", False))
        self.compact_var = tk.BooleanVar(value=self.settings.get("compact", False))

        self._build_menu()
        self._build_toolbar()
        self._build_tree()
        self._build_statusbar()
        self._bind_keys()
        self._apply_ontop(save=False)
        self._apply_compact(save=False)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.refresh()

        # 시작 시 '오늘 요약' 팝업 + 주기적 갱신(자정 넘어가도 남은기간 최신화)
        if self.summary_var.get():
            self.after(500, lambda: self._show_summary(on_startup=True))
        self._refresh_job = self.after(self.REFRESH_MS, self._periodic_refresh)

    # -------- 스타일/폰트 --------
    def _setup_style(self):
        try:
            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
                tkfont.nametofont(name).configure(family="맑은 고딕", size=10)
        except Exception:
            pass
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Treeview", rowheight=26)
        style.configure("Treeview.Heading", font=("맑은 고딕", 10, "bold"))
        # 완료 항목용 취소선 폰트
        self.strike_font = tkfont.Font(family="맑은 고딕", size=10, overstrike=1)

    # -------- 메뉴 --------
    def _build_menu(self):
        menubar = tk.Menu(self)

        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="데이터 폴더 열기", command=self._open_data_dir)
        m_file.add_command(label="백업(내보내기)…", command=self._export)
        m_file.add_command(label="가져오기(병합)…", command=self._import)
        m_file.add_separator()
        m_file.add_command(label="종료", command=self._on_close)
        menubar.add_cascade(label="파일", menu=m_file)

        m_set = tk.Menu(menubar, tearoff=0)
        m_set.add_checkbutton(label="Windows 시작 시 자동 실행",
                              variable=self.autostart_var, command=self._toggle_autostart)
        m_set.add_checkbutton(label="시작 시 '오늘 요약' 표시", variable=self.summary_var)
        m_set.add_separator()
        m_set.add_checkbutton(label="항상 위에 표시", variable=self.ontop_var,
                              command=self._apply_ontop)
        m_set.add_checkbutton(label="간단히 보기(스티커 모드)", variable=self.compact_var,
                              command=self._apply_compact)
        menubar.add_cascade(label="설정", menu=m_set)

        m_cal = tk.Menu(menubar, tearoff=0)
        m_cal.add_command(label="캘린더 / 일정 관리 열기", command=self._open_calendar)
        menubar.add_cascade(label="캘린더", menu=m_cal)

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="단축키 / 색상 안내", command=self._show_help)
        m_help.add_command(label="정보", command=self._show_about)
        menubar.add_cascade(label="도움말", menu=m_help)

        self.config(menu=menubar)

    # -------- 상단 툴바(빠른 추가 + 필터) --------
    def _build_toolbar(self):
        # 1행: 빠른 추가
        row1 = ttk.Frame(self, padding=(10, 8, 10, 2))
        row1.pack(fill="x")
        self.row1 = row1
        ttk.Label(row1, text="빠른 추가:").pack(side="left")
        self.quick_entry = ttk.Entry(row1)
        self.quick_entry.pack(side="left", fill="x", expand=True, padx=6)
        self.quick_entry.bind("<Return>", lambda e: self._quick_add())
        ttk.Button(row1, text="추가", command=self._quick_add).pack(side="left")
        ttk.Button(row1, text="상세 추가…", command=self._add_detail).pack(side="left", padx=(6, 0))
        ttk.Button(row1, text="📅 캘린더", command=self._open_calendar).pack(side="left", padx=(6, 0))

        # 2행: 정렬/필터/검색
        row2 = ttk.Frame(self, padding=(10, 2, 10, 8))
        row2.pack(fill="x")
        self.row2 = row2
        ttk.Label(row2, text="정렬:").pack(side="left")
        cb_sort = ttk.Combobox(row2, textvariable=self.sort_mode, values=SORT_MODES,
                               width=12, state="readonly")
        cb_sort.pack(side="left", padx=(4, 12))
        cb_sort.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        ttk.Label(row2, text="카테고리:").pack(side="left")
        self.cb_filter = ttk.Combobox(row2, textvariable=self.filter_cat,
                                      values=[FILTER_ALL], width=12, state="readonly")
        self.cb_filter.pack(side="left", padx=(4, 12))
        self.cb_filter.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        ttk.Checkbutton(row2, text="완료 숨기기", variable=self.hide_done,
                        command=self.refresh).pack(side="left", padx=(0, 12))

        ttk.Label(row2, text="검색:").pack(side="left")
        e_search = ttk.Entry(row2, textvariable=self.search_var, width=14)
        e_search.pack(side="left", padx=4)
        self.search_var.trace_add("write", lambda *a: self.refresh())

    # -------- 업무 목록(Treeview) --------
    def _build_tree(self):
        frame = ttk.Frame(self, padding=(10, 0, 10, 6))
        frame.pack(fill="both", expand=True)

        cols = ("done", "priority", "title", "category", "due", "dday")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        headings = {
            "done": ("완료", 44, "center", 40),
            "priority": ("우선순위", 64, "center", 56),
            "title": ("업무 내용", 160, "w", 80),
            "category": ("카테고리", 88, "center", 60),
            "due": ("마감일", 92, "center", 70),
            "dday": ("남은기간", 80, "center", 60),
        }
        for c, (text, width, anchor, minw) in headings.items():
            self.tree.heading(c, text=text, command=lambda cc=c: self._sort_by_column(cc))
            self.tree.column(c, width=width, anchor=anchor, minwidth=minw,
                             stretch=(c == "title"))

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # 행 색상 태그
        self.tree.tag_configure("overdue", background=COLOR_OVERDUE)
        self.tree.tag_configure("today", background=COLOR_TODAY)
        self.tree.tag_configure("soon", background=COLOR_SOON)
        self.tree.tag_configure("done", foreground=COLOR_DONE_FG, font=self.strike_font)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Return>", lambda e: self._edit_selected())
        self.tree.bind("<space>", lambda e: self._toggle_selected())
        self.tree.bind("<Delete>", lambda e: self._delete_selected())

    # -------- 하단 버튼/상태바 --------
    def _build_statusbar(self):
        bar = ttk.Frame(self, padding=(10, 0, 10, 8))
        bar.pack(fill="x")
        self.bottom_bar = bar
        ttk.Button(bar, text="완료/취소", command=self._toggle_selected).pack(side="left")
        ttk.Button(bar, text="수정", command=self._edit_selected).pack(side="left", padx=6)
        ttk.Button(bar, text="삭제", command=self._delete_selected).pack(side="left")
        ttk.Button(bar, text="오늘 요약", command=self._show_summary).pack(side="left", padx=(16, 0))

        self.status = ttk.Label(bar, text="", anchor="e")
        self.status.pack(side="right", fill="x", expand=True)

    def _bind_keys(self):
        self.bind("<Control-n>", lambda e: self.quick_entry.focus_set())
        self.bind("<Control-f>", lambda e: self.refresh())
        self.bind("<Control-d>", lambda e: self._show_summary())
        self.bind("<Control-l>", lambda e: self._open_calendar())

    # ========================================================
    # 동작
    # ========================================================
    def visible_tasks(self):
        items = list(self.store.tasks)
        if self.hide_done.get():
            items = [t for t in items if not t.done]
        cat = self.filter_cat.get()
        if cat and cat != FILTER_ALL:
            items = [t for t in items if t.category == cat]
        q = self.search_var.get().strip().lower()
        if q:
            items = [t for t in items
                     if q in t.title.lower() or q in t.category.lower() or q in t.note.lower()]
        items.sort(key=sort_key(self.sort_mode.get()))
        return items

    def _row_tag(self, task: Task) -> str:
        if task.done:
            return "done"
        n = business_days_left(task.due)
        if n is None:
            return ""
        if n <= 1:
            return "overdue"
        if n <= 4:
            return "today"
        if n <= 7:
            return "soon"
        return ""

    def refresh(self):
        # 카테고리 필터 목록 갱신
        cats = [FILTER_ALL] + self.store.categories()
        self.cb_filter["values"] = cats
        if self.filter_cat.get() not in cats:
            self.filter_cat.set(FILTER_ALL)

        # 목록 다시 그리기 (현재 선택 유지)
        sel = self.tree.selection()
        sel_id = sel[0] if sel else None
        self.tree.delete(*self.tree.get_children())
        for t in self.visible_tasks():
            check = "☑" if t.done else "□"
            due_disp = t.due if t.due else "-"
            self.tree.insert("", "end", iid=t.id,
                             values=(check, t.priority, t.title, t.category or "-",
                                     due_disp, dday_text(t)),
                             tags=(self._row_tag(t),))
        if sel_id and self.tree.exists(sel_id):
            self.tree.selection_set(sel_id)
            self.tree.see(sel_id)
        self._update_status()

    def _update_status(self):
        all_t = self.store.tasks
        total = len(all_t)
        open_n = sum(1 for t in all_t if not t.done)
        overdue = sum(1 for t in all_t if not t.done and (days_left(t.due) or 0) < 0
                      and days_left(t.due) is not None)
        today_n = sum(1 for t in all_t if not t.done and days_left(t.due) == 0)
        self.status.config(
            text=f"총 {total}건 · 미완료 {open_n}건 · 지연 {overdue}건 · 오늘마감 {today_n}건"
        )

    # ----- 선택 항목 헬퍼 -----
    def _selected_task(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.store.by_id(sel[0])

    # ----- 추가 -----
    def _quick_add(self):
        title = self.quick_entry.get().strip()
        if not title:
            self.bell()
            return
        cat = self.filter_cat.get()
        cat = cat if cat and cat != FILTER_ALL else ""
        self.store.add(Task(title=title, category=cat, priority="중간"))
        self.quick_entry.delete(0, "end")
        self.refresh()

    def _add_detail(self):
        dlg = TaskDialog(self, task=None, categories=self.store.categories())
        self.wait_window(dlg)
        if dlg.result:
            self.store.add(dlg.result)
            self.refresh()

    # ----- 수정 -----
    def _on_double_click(self, event):
        # 완료 열을 더블클릭하면 토글, 그 외에는 수정
        col = self.tree.identify_column(event.x)
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell" and col == "#1":
            self._toggle_selected()
        else:
            self._edit_selected()

    def _edit_selected(self):
        task = self._selected_task()
        if not task:
            return
        dlg = TaskDialog(self, task=task, categories=self.store.categories())
        self.wait_window(dlg)
        if dlg.result:
            self.store.save()
            self.refresh()

    # ----- 완료 토글 -----
    def _toggle_selected(self):
        task = self._selected_task()
        if not task:
            return
        task.done = not task.done
        task.done_at = datetime.now().isoformat(timespec="seconds") if task.done else ""
        self.store.save()
        self.refresh()

    # ----- 삭제 -----
    def _delete_selected(self):
        task = self._selected_task()
        if not task:
            return
        if messagebox.askyesno("삭제 확인", f"'{task.title}'\n이 업무를 삭제할까요?", parent=self):
            self.store.remove(task.id)
            self.refresh()

    # ----- 정렬(헤더 클릭) -----
    def _sort_by_column(self, col):
        mapping = {"due": "마감일 빠른순", "priority": "우선순위순",
                   "category": "카테고리순", "dday": "긴급도순"}
        if col in mapping:
            self.sort_mode.set(mapping[col])
            self.refresh()

    # ----- 파일 메뉴 -----
    def _open_data_dir(self):
        try:
            os.startfile(str(data_dir()))
        except Exception as e:
            messagebox.showinfo("데이터 폴더", f"{data_dir()}\n\n({e})", parent=self)

    def _export(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="백업 파일 저장",
            defaultextension=".json", filetypes=[("JSON", "*.json")],
            initialfile=f"todo-backup-{date.today().isoformat()}.json")
        if not path:
            return
        try:
            payload = {"version": DATA_VERSION, "tasks": [asdict(t) for t in self.store.tasks]}
            Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            messagebox.showinfo("백업 완료", f"{len(self.store.tasks)}건을 저장했습니다.", parent=self)
        except Exception as e:
            messagebox.showerror("백업 실패", str(e), parent=self)

    def _import(self):
        path = filedialog.askopenfilename(
            parent=self, title="가져올 파일 선택", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8"))
            items = raw.get("tasks", []) if isinstance(raw, dict) else raw
            valid_keys = {f.name for f in dataclass_fields(Task)}
            existing = {t.id for t in self.store.tasks}
            added = 0
            for it in items:
                t = Task(**{k: v for k, v in it.items() if k in valid_keys})
                if t.id in existing:        # id 충돌 시 새 id 부여
                    t.id = uuid.uuid4().hex[:8]
                self.store.tasks.append(t)
                added += 1
            self.store.save()
            self.refresh()
            messagebox.showinfo("가져오기 완료", f"{added}건을 추가했습니다.", parent=self)
        except Exception as e:
            messagebox.showerror("가져오기 실패", str(e), parent=self)

    # ----- 도움말 -----
    def _show_help(self):
        messagebox.showinfo(
            "단축키 / 색상 안내",
            "[단축키]\n"
            "  Enter(빠른추가칸)  : 업무 빠르게 추가\n"
            "  더블클릭            : 업무 수정 (완료열 더블클릭=완료 토글)\n"
            "  Space               : 선택 업무 완료/취소\n"
            "  Delete              : 선택 업무 삭제\n"
            "  Ctrl+N              : 빠른 추가칸으로 이동\n"
            "  Ctrl+D              : '오늘 요약' 다시 보기\n"
            "  Ctrl+L              : 캘린더 / 일정 관리 열기\n\n"
            "[행 색상 = 긴급도]\n"
            "  빨강 : 영업일 기준 1일 전~마감일 지남\n"
            "  주황 : 영업일 기준 2~4일 전\n"
            "  노랑 : 영업일 기준 5~7일 전\n"
            "  회색 취소선 : 완료\n\n"
            "[헤더 클릭] 마감일·우선순위·카테고리·남은기간 헤더를 누르면 그 기준으로 정렬됩니다.",
            parent=self)

    def _show_about(self):
        messagebox.showinfo(
            "정보",
            f"{APP_TITLE}\n\n"
            "Python + Tkinter 데스크톱 앱\n"
            f"데이터 저장 위치:\n{DATA_FILE}",
            parent=self)

    # ----- 설정 / 알림 -----
    def _save_pref(self, key, value):
        self.settings[key] = value
        save_settings(self.settings)

    def _toggle_autostart(self):
        want = self.autostart_var.get()
        try:
            set_autostart(want)
            messagebox.showinfo(
                "자동 실행",
                "Windows 시작 시 자동 실행을 켰습니다." if want
                else "Windows 시작 시 자동 실행을 껐습니다.",
                parent=self)
        except Exception as e:
            messagebox.showerror("자동 실행 설정 실패", str(e), parent=self)
            self.autostart_var.set(is_autostart_enabled())   # 실패 시 실제 상태로 복원

    def _apply_ontop(self, save=True):
        self.attributes("-topmost", self.ontop_var.get())
        if save:
            self._save_pref("always_on_top", self.ontop_var.get())

    def _apply_compact(self, save=True):
        """스티커 모드: 필터줄·하단 버튼·부가 열을 숨겨 폭을 좁힌다."""
        compact = self.compact_var.get()
        if compact:
            self.row2.pack_forget()
            self.bottom_bar.pack_forget()
            self.tree["displaycolumns"] = ("done", "title", "dday")
            self.minsize(280, 200)
        else:
            self.row2.pack(fill="x", after=self.row1)
            self.bottom_bar.pack(fill="x")
            self.tree["displaycolumns"] = "#all"
            self.minsize(600, 360)
        if save:
            self._save_pref("compact", compact)

    def _periodic_refresh(self):
        self.refresh()                       # 날짜 경과(자정 등) 반영
        self._refresh_job = self.after(self.REFRESH_MS, self._periodic_refresh)

    # ----- 캘린더 / 일정 -----
    def _open_calendar(self):
        reload_public_holidays()   # DB_holiday.xlsx 변경분 반영
        self.refresh()             # 공휴일 변동 시 남은기간 재계산
        win = getattr(self, "_cal_win", None)
        if win is not None and win.winfo_exists():
            win.lift()
            win.focus_set()
            return
        self._cal_win = CalendarWindow(self)

    def on_schedules_changed(self):
        """일정이 추가/수정/삭제됐을 때: 비근무일 재계산 후 목록 새로고침."""
        WORK_CALENDAR.set_events(self.sched.events)
        self.refresh()

    def _show_summary(self, on_startup=False):
        """오늘 처리할(지연/1영업일 이내 마감) 업무를 팝업으로 요약."""
        items = [t for t in self.store.tasks if not t.done]
        overdue = [t for t in items if business_days_left(t.due) is not None and business_days_left(t.due) < 0]
        duesoon = [t for t in items if business_days_left(t.due) in (0, 1)]
        overdue.sort(key=sort_key("긴급도순"))
        duesoon.sort(key=sort_key("긴급도순"))

        # 시작 팝업인데 급한 게 없으면 조용히 넘어감(잔소리 방지)
        if on_startup and not overdue and not duesoon:
            return

        dlg = tk.Toplevel(self)
        dlg.title("오늘의 업무 요약")
        dlg.transient(self)
        dlg.configure(padx=14, pady=12)
        dlg.resizable(False, False)

        head = (f"오늘  {date.today().isoformat()}      "
                f"지연 {len(overdue)}건 · 1영업일 이내 {len(duesoon)}건")
        ttk.Label(dlg, text=head, font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(0, 8))

        txt = tk.Text(dlg, width=54, height=14, wrap="word",
                      relief="flat", background="#fafafa", borderwidth=0)
        txt.pack(fill="both", expand=True)
        txt.tag_configure("red", foreground="#c0392b")
        txt.tag_configure("orange", foreground="#b9770e")
        txt.tag_configure("muted", foreground="#888888")
        txt.tag_configure("sec", font=("맑은 고딕", 10, "bold"), spacing1=6)

        def add_section(title, tasks, color):
            if not tasks:
                return
            txt.insert("end", f"{title}  ({len(tasks)}건)\n", "sec")
            for t in tasks[:10]:
                cat = f"[{t.category}] " if t.category else ""
                txt.insert("end", f"    · {cat}{t.title}")
                txt.insert("end", f"    {dday_text(t)} · {t.priority}\n", color)
            if len(tasks) > 10:
                txt.insert("end", f"    … 외 {len(tasks) - 10}건\n", "muted")
            txt.insert("end", "\n")

        if not overdue and not duesoon:
            txt.insert("end", "급한 업무가 없습니다. 좋은 하루 되세요!\n", "muted")
        add_section("지연된 업무", overdue, "red")
        add_section("1영업일 이내 마감", duesoon, "red")
        txt.configure(state="disabled")

        bottom = ttk.Frame(dlg)
        bottom.pack(fill="x", pady=(10, 0))
        ttk.Checkbutton(bottom, text="시작 시 이 요약 표시",
                        variable=self.summary_var).pack(side="left")
        ttk.Button(bottom, text="닫기", command=dlg.destroy).pack(side="right")

        dlg.bind("<Escape>", lambda e: dlg.destroy())
        dlg.update_idletasks()
        try:
            px = self.winfo_rootx()
            py = self.winfo_rooty()
            pw = self.winfo_width()
            w = dlg.winfo_width()
            dlg.geometry(f"+{px + (pw - w) // 2}+{py + 80}")
        except Exception:
            pass
        dlg.grab_set()

    # ----- 종료 -----
    def _on_close(self):
        try:
            self.after_cancel(self._refresh_job)
        except Exception:
            pass
        self.settings.update({
            "geometry": self.geometry(),
            "sort": self.sort_mode.get(),
            "filter_cat": self.filter_cat.get(),
            "hide_done": self.hide_done.get(),
            "show_startup_summary": self.summary_var.get(),
            "always_on_top": self.ontop_var.get(),
            "compact": self.compact_var.get(),
        })
        save_settings(self.settings)
        self.destroy()


# ============================================================
# 진입점
# ============================================================
def _enable_dpi_awareness():
    """Windows 고해상도 모니터에서 또렷하게 보이도록."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main():
    # 자체 점검: 창을 띄우지 않고 UI 구성만 검증
    if "--selftest" in sys.argv:
        assert normalize_due_input("0708", default_year=2026) == "2026-07-08"
        assert normalize_due_input("20260708") == "2026-07-08"
        app = TodoApp()
        app.withdraw()
        app.update_idletasks()
        app.refresh()
        app.destroy()
        print("selftest OK")
        return

    _enable_dpi_awareness()
    try:
        app = TodoApp()
        app.mainloop()
    except Exception as e:
        # --windowed(콘솔 없음)로 빌드된 경우를 대비해 오류를 기록/표시
        try:
            (data_dir() / "error.log").write_text(
                datetime.now().isoformat() + "\n" + repr(e), encoding="utf-8")
        except Exception:
            pass
        try:
            messagebox.showerror(APP_TITLE, f"오류가 발생했습니다:\n{e}")
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
