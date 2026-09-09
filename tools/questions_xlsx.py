"""questions.xlsx <-> question dicts.

Sheet "Questions", one row per question:
  ID | Pool (Normal/Hard) | Topic | Question | A | B | C | D | Correct (A-D) | Explanation | Source | Version
Rows with an empty Question are ignored, so you can leave gaps. C and D may be empty (2- or 3-option questions).
"""
import sys

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is missing:  pip install openpyxl")

HEADERS = ["ID", "Pool", "Topic", "Question", "A", "B", "C", "D", "Correct", "Explanation", "Source", "Version"]
WIDTHS = [9, 9, 12, 60, 28, 28, 28, 28, 9, 60, 40, 12]
POOL_NAMES = {"Normal": "N", "Hard": "H"}


def _s(v):
    return "" if v is None else str(v).strip()


def read_questions(path):
    """Returns a list of dicts {id, pool: 'N'|'H', topic, q, options, correct, explain, source, version, row}.
    Raises SystemExit with every problem listed (row numbers refer to the sheet)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "Questions" not in wb.sheetnames:
        sys.exit(f"{path}: no sheet named 'Questions'")
    ws = wb["Questions"]
    rows = ws.iter_rows(values_only=True)
    header = [_s(c) for c in next(rows)]
    col = {}
    for name in HEADERS:
        if name not in header:
            sys.exit(f"{path}: header row must contain {HEADERS}, missing {name!r}")
        col[name] = header.index(name)
    out, problems, seen_ids = [], [], set()
    for rno, row in enumerate(rows, 2):
        row = list(row) + [None] * len(HEADERS)
        get = lambda name: _s(row[col[name]])  # noqa: E731
        if not get("Question"):
            continue
        pool = get("Pool").capitalize()
        if pool not in POOL_NAMES:
            problems.append(f"row {rno}: Pool must be Normal or Hard (got {get('Pool')!r})")
            continue
        options = [get(k) for k in ("A", "B", "C", "D")]
        while options and not options[-1]:
            options.pop()
        if any(not o for o in options) or not 2 <= len(options) <= 4:
            problems.append(f"row {rno}: fill A and B (C, D optional, no gaps)")
            continue
        letter = get("Correct").upper()[:1]
        if letter not in "ABCD"[:len(options)] or not letter:
            problems.append(f"row {rno}: Correct must be one of {'/'.join('ABCD'[:len(options)])}")
            continue
        qid = get("ID") or f"row{rno}"
        if qid in seen_ids:
            problems.append(f"row {rno}: duplicate ID {qid}")
            continue
        seen_ids.add(qid)
        out.append({"id": qid, "pool": POOL_NAMES[pool], "topic": get("Topic"), "q": get("Question"),
                    "options": options, "correct": "ABCD".index(letter), "explain": get("Explanation"),
                    "source": get("Source"), "version": get("Version"), "row": rno})
    if problems:
        sys.exit(f"{path}: fix these rows first:\n  " + "\n  ".join(problems))
    if not any(q["pool"] == "N" for q in out) or not any(q["pool"] == "H" for q in out):
        sys.exit(f"{path}: need at least one Normal and one Hard question")
    return out


def write_questions(path, questions):
    """Creates a fresh workbook with the questions, dropdowns and a help sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Questions"
    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F5597")
        cell.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS[c - 1]
    for q in questions:
        opts = list(q["options"]) + [""] * (4 - len(q["options"]))
        ws.append([q["id"], "Normal" if q["pool"] == "N" else "Hard", q.get("topic", ""), q["q"],
                   opts[0], opts[1], opts[2], opts[3], "ABCD"[q["correct"]], q.get("explain", ""),
                   q.get("source", ""), q.get("version", "")])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{max(len(questions) + 1, 2)}"
    dv_pool = DataValidation(type="list", formula1='"Normal,Hard"', allow_blank=True, showErrorMessage=True,
                             errorTitle="Pool", error="Normal or Hard")
    dv_ans = DataValidation(type="list", formula1='"A,B,C,D"', allow_blank=True, showErrorMessage=True,
                            errorTitle="Correct", error="A, B, C or D")
    ws.add_data_validation(dv_pool)
    ws.add_data_validation(dv_ans)
    dv_pool.add("B2:B2000")
    dv_ans.add("I2:I2000")

    h = wb.create_sheet("How to edit")
    h.column_dimensions["A"].width = 110
    for line in [
        "HOW TO EDIT THE QUESTIONS",
        "",
        "- One row = one question. Add rows at the bottom, delete rows you do not want. Empty rows are ignored.",
        "- Pool: Normal or Hard (dropdown). Normal = the easier pool, Hard = deep-cut knowledge.",
        "- Question: the text shown in game (keep it to about 2 lines; plain letters, no emoji).",
        "- A, B, C, D: the answers. C and D may be empty for 2- or 3-answer questions (no gaps: fill A, B, then C, then D).",
        "- Correct: the letter of the right answer (dropdown A-D).",
        "- ID: any short unique text (N-077, H-080 ...). Topic / Explanation / Source / Version are notes for you, not shown in game.",
        "",
        "AFTER EDITING",
        "  1. Save this file (keep the name questions.xlsx, keep the sheet name Questions).",
        "  2. Run:  python tools/build.py      (needs Python 3 and:  pip install openpyxl)",
        "  3. The rebuilt packs are in dist/. The build stops and tells you the row number if something is wrong.",
        "",
        "One special question is worth keeping: 'Who made this modpack?' with B = W1ndLeaf.",
    ]:
        h.append([line])
    h["A1"].font = Font(bold=True, size=13)
    h["A10"].font = Font(bold=True)
    wb.save(path)


SETTINGS_ROWS = [   # (label, default, note)  - the general settings at the top of the Settings sheet
    ("Difficulty", "Normal", "Normal or Hard"),
    ("Interval (minutes)", 5, "1 - 30, minutes between questions"),
    ("Score display", "On", "On or Off - the sidebar with next question / correct / wrong"),
    ("Menu on first join", "Off", "On = the menu opens once per world when you join"),
]


def write_settings_sheet(path, rewards, punishments):
    """Adds (or replaces) the Settings sheet: general settings, then one On/Off row per reward and punishment.
    rewards / punishments: lists of (key, name)."""
    wb = openpyxl.load_workbook(path)
    if "Settings" in wb.sheetnames:
        del wb["Settings"]
    ws = wb.create_sheet("Settings", 1)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 70
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="2F5597")

    def header(*cells):
        ws.append(list(cells))
        for c in range(1, len(cells) + 1):
            ws.cell(row=ws.max_row, column=c).font = head
            ws.cell(row=ws.max_row, column=c).fill = fill
    header("Setting", "Value", "What it does - these are the values EVERY new world starts with (rebuild after editing)")
    for label, default, note in SETTINGS_ROWS:
        ws.append([label, default, note])
    ws.append([])
    header("Rewards", "On/Off", "Off = can never happen (you can still switch it in game, per world)")
    for _, name in rewards:
        ws.append([name, "On"])
    ws.append([])
    header("Punishments", "On/Off", "Off = can never happen")
    for _, name in punishments:
        ws.append([name, "On"])
    onoff = DataValidation(type="list", formula1='"On,Off"', allow_blank=True)
    diff = DataValidation(type="list", formula1='"Normal,Hard"', allow_blank=True)
    ws.add_data_validation(onoff)
    ws.add_data_validation(diff)
    diff.add("B2")
    onoff.add("B4:B5")
    onoff.add(f"B7:B{ws.max_row}")
    ws.freeze_panes = "A2"
    wb.save(path)


def read_settings(path, rewards, punishments):
    """Returns {"diff": 1|2, "interval": int, "sidebar": 0|1, "auto_menu": bool, "off_rewards": [keys], "off_punishments": [keys]}
    or None when the workbook has no Settings sheet."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "Settings" not in wb.sheetnames:
        return None
    by_name = {"Rewards": {name: key for key, name in rewards}, "Punishments": {name: key for key, name in punishments}}
    out = {"diff": 1, "interval": 5, "sidebar": 1, "auto_menu": False, "off_rewards": [], "off_punishments": []}
    section, problems = None, []
    for rno, row in enumerate(wb["Settings"].iter_rows(values_only=True), 1):
        label, value = (_s(row[0]) if row else ""), (_s(row[1]) if row and len(row) > 1 else "")
        if not label:
            continue
        if label in by_name:
            section = label
            continue
        if section:
            key = by_name[section].get(label)
            if key is None:
                problems.append(f"Settings row {rno}: unknown {section[:-1].lower()} {label!r}")
            elif value.lower() == "off":
                out["off_rewards" if section == "Rewards" else "off_punishments"].append(key)
            elif value.lower() not in ("on", ""):
                problems.append(f"Settings row {rno}: {label}: use On or Off")
        elif label == "Difficulty":
            if value.capitalize() not in ("Normal", "Hard"):
                problems.append(f"Settings row {rno}: Difficulty must be Normal or Hard")
            out["diff"] = 2 if value.capitalize() == "Hard" else 1
        elif label == "Interval (minutes)":
            try:
                out["interval"] = max(1, min(30, int(float(value))))
            except ValueError:
                problems.append(f"Settings row {rno}: Interval must be a number 1-30")
        elif label == "Score display":
            out["sidebar"] = 0 if value.lower() == "off" else 1
        elif label == "Menu on first join":
            out["auto_menu"] = value.lower() == "on"
    if problems:
        sys.exit(f"{path}: fix the Settings sheet first:\n  " + "\n  ".join(problems))
    return out


if __name__ == "__main__":  # quick check:  python questions_xlsx.py ../questions.xlsx
    qs = read_questions(sys.argv[1])
    print(len(qs), "questions,", sum(q["pool"] == "N" for q in qs), "Normal,", sum(q["pool"] == "H" for q in qs), "Hard")
