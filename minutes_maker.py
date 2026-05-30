"""Transcript-to-Excel meeting minutes generator.

The module contains UI-independent parsing and Excel rendering code so it can be
used from both the desktop app and tests. The Excel writer uses only Python's
standard library and emits a minimal .xlsx package that opens in Excel, Numbers,
and LibreOffice.
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

ACTION_PATTERNS = [
    re.compile(r"(?P<owner>[\wぁ-んァ-ヶ一-龠々ー・\s]{1,20})[：:、\s]+(?P<task>.+?)(?:[（(](?P<due>[^）)]+)[）)])?$"),
]

DECISION_WORDS = ("決定", "合意", "承認", "採択", "決まり", "決め", "することにな")
ACTION_WORDS = ("TODO", "ToDo", "タスク", "宿題", "対応", "担当", "アクション", "Action")
AGENDA_WORDS = ("議題", "アジェンダ", "agenda", "Agenda")
ATTENDEE_WORDS = ("参加者", "出席者", "参加", "attendees", "Attendees")
DATE_PATTERNS = (
    re.compile(r"(?P<year>20\d{2})[/-](?P<month>\d{1,2})[/-](?P<day>\d{1,2})"),
    re.compile(r"(?P<year>20\d{2})年\s*(?P<month>\d{1,2})月\s*(?P<day>\d{1,2})日"),
)


@dataclass
class ActionItem:
    owner: str = "未定"
    task: str = ""
    due: str = "未定"


@dataclass
class MeetingMinutes:
    title: str = "議事録"
    meeting_date: str = ""
    attendees: list[str] = field(default_factory=list)
    agenda: list[str] = field(default_factory=list)
    summary: list[str] = field(default_factory=list)
    decisions: list[str] = field(default_factory=list)
    action_items: list[ActionItem] = field(default_factory=list)


def normalize_lines(text: str) -> list[str]:
    """Return meaningful transcript lines with bullets and whitespace normalized."""
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    lines: list[str] = []
    for raw in normalized.split("\n"):
        line = raw.strip()
        line = re.sub(r"^[\-・*●○◆■□\d]+[.)、\s]*", "", line).strip()
        if line:
            lines.append(line)
    return lines


def split_values(value: str) -> list[str]:
    parts = re.split(r"[,、/／・]\s*|\s{2,}", value)
    return [part.strip() for part in parts if part.strip()]


def clean_after_label(line: str) -> str:
    return re.sub(r"^[^：:]+[：:]\s*", "", line).strip()


def extract_date(lines: Iterable[str]) -> str:
    for line in lines:
        for pattern in DATE_PATTERNS:
            match = pattern.search(line)
            if match:
                year = int(match.group("year"))
                month = int(match.group("month"))
                day = int(match.group("day"))
                return dt.date(year, month, day).isoformat()
    return dt.date.today().isoformat()


def extract_title(lines: list[str]) -> str:
    for line in lines[:8]:
        if any(word in line for word in ("会議", "MTG", "ミーティング", "定例", "議事録")):
            return clean_after_label(line) or "議事録"
    return "議事録"


def parse_action(line: str) -> ActionItem:
    value = clean_after_label(line)
    value = re.sub(r"^(TODO|ToDo|タスク|宿題|対応|担当|アクション|Action)\s*[：:]?\s*", "", value)
    for pattern in ACTION_PATTERNS:
        match = pattern.match(value)
        if match and len(match.group("task")) > 2:
            return ActionItem(
                owner=match.group("owner").strip() or "未定",
                task=match.group("task").strip(),
                due=(match.group("due") or "未定").strip(),
            )
    return ActionItem(task=value)


def parse_transcript(text: str) -> MeetingMinutes:
    """Parse a transcript into a structured minutes object using safe heuristics."""
    lines = normalize_lines(text)
    minutes = MeetingMinutes(title=extract_title(lines), meeting_date=extract_date(lines))
    current_section = "summary"

    for line in lines:
        lower = line.lower()
        if any(word in line for word in ATTENDEE_WORDS):
            current_section = "attendees"
            values = split_values(clean_after_label(line))
            minutes.attendees.extend(values)
            continue
        if any(word in line for word in AGENDA_WORDS):
            current_section = "agenda"
            value = clean_after_label(line)
            if value and value != line:
                minutes.agenda.append(value)
            continue
        if "決定事項" in line or lower.startswith("decisions"):
            current_section = "decisions"
            continue
        if "アクションアイテム" in line or "todo" in lower or lower.startswith("actions"):
            current_section = "actions"
            value = clean_after_label(line)
            if value and value != line:
                minutes.action_items.append(parse_action(value))
            continue

        if any(word in line for word in DECISION_WORDS):
            minutes.decisions.append(clean_after_label(line))
        elif any(word in line for word in ACTION_WORDS):
            minutes.action_items.append(parse_action(line))
        elif current_section == "attendees":
            minutes.attendees.extend(split_values(line))
        elif current_section == "agenda":
            minutes.agenda.append(line)
        elif current_section == "decisions":
            minutes.decisions.append(line)
        elif current_section == "actions":
            minutes.action_items.append(parse_action(line))
        else:
            minutes.summary.append(line)

    minutes.attendees = unique_keep_order(minutes.attendees)
    if not minutes.agenda:
        minutes.agenda = ["トランスクリプト内容の確認"]
    if not minutes.summary:
        minutes.summary = ["トランスクリプトから自動生成されました。必要に応じて内容を確認してください。"]
    if not minutes.decisions:
        minutes.decisions = ["明確な決定事項は検出されませんでした。"]
    if not minutes.action_items:
        minutes.action_items = [ActionItem(task="必要に応じてアクションアイテムを追記してください。")]
    return minutes


def unique_keep_order(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        cleaned = value.strip()
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            result.append(cleaned)
    return result


def xml_escape(value: object) -> str:
    return html.escape(str(value), quote=True)


def sheet_cell(reference: str, value: object, style: int = 0) -> str:
    escaped = xml_escape(value)
    return f'<c r="{reference}" t="inlineStr" s="{style}"><is><t>{escaped}</t></is></c>'


def row_xml(row_number: int, values: list[object], styles: list[int] | None = None) -> str:
    styles = styles or [0] * len(values)
    cells = []
    for index, value in enumerate(values, start=1):
        column = chr(ord("A") + index - 1)
        cells.append(sheet_cell(f"{column}{row_number}", value, styles[index - 1]))
    return f'<row r="{row_number}">{"".join(cells)}</row>'


def build_sheet_xml(minutes: MeetingMinutes) -> str:
    rows: list[str] = []
    merges = ["A1:D1", "B4:D4"]

    rows.append(row_xml(1, [minutes.title, "", "", ""], [1, 1, 1, 1]))
    rows.append(row_xml(3, ["開催日", minutes.meeting_date, "", ""], [2, 0, 0, 0]))
    rows.append(row_xml(4, ["参加者", "、".join(minutes.attendees) if minutes.attendees else "未入力", "", ""], [2, 0, 0, 0]))

    row = 6

    def section(title: str) -> None:
        nonlocal row
        rows.append(row_xml(row, [title, "", "", ""], [3, 3, 3, 3]))
        merges.append(f"A{row}:D{row}")
        row += 1

    def bullets(values: Iterable[str]) -> None:
        nonlocal row
        for value in values:
            rows.append(row_xml(row, ["・", value, "", ""], [0, 0, 0, 0]))
            merges.append(f"B{row}:D{row}")
            row += 1

    section("議題")
    bullets(minutes.agenda)
    row += 1
    section("要約")
    bullets(minutes.summary[:12])
    row += 1
    section("決定事項")
    bullets(minutes.decisions)
    row += 1
    section("アクションアイテム")
    rows.append(row_xml(row, ["No.", "担当", "内容", "期限"], [4, 4, 4, 4]))
    row += 1
    for index, item in enumerate(minutes.action_items, start=1):
        rows.append(row_xml(row, [index, item.owner, item.task, item.due]))
        row += 1

    merge_xml = "".join(f'<mergeCell ref="{ref}"/>' for ref in merges)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols><col min="1" max="1" width="8" customWidth="1"/><col min="2" max="2" width="28" customWidth="1"/><col min="3" max="3" width="28" customWidth="1"/><col min="4" max="4" width="28" customWidth="1"/></cols>
  <sheetData>{''.join(rows)}</sheetData>
  <mergeCells count="{len(merges)}">{merge_xml}</mergeCells>
</worksheet>'''


def build_styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="18"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def create_xlsx(minutes: MeetingMinutes, output_path: Path) -> None:
    files = {
        "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>''',
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''',
        "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="議事録" sheetId="1" r:id="rId1"/></sheets>
</workbook>''',
        "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>''',
        "xl/worksheets/sheet1.xml": build_sheet_xml(minutes),
        "xl/styles.xml": build_styles_xml(),
    }
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def save_minutes_excel(transcript: str, output_path: Path) -> Path:
    minutes = parse_transcript(transcript)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    create_xlsx(minutes, output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="トランスクリプト文章からExcel議事録を生成します。")
    parser.add_argument("input", type=Path, help="トランスクリプトのテキストファイル")
    parser.add_argument("output", type=Path, nargs="?", default=Path("minutes.xlsx"), help="出力するExcelファイル")
    args = parser.parse_args()
    transcript = args.input.read_text(encoding="utf-8")
    save_minutes_excel(transcript, args.output)
    print(f"Excel議事録を作成しました: {args.output}")


if __name__ == "__main__":
    main()
