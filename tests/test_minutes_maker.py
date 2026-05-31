import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from minutes_maker import (
    parse_action,
    parse_transcript,
    save_minutes_excel,
)


from pathlib import Path
import zipfile

from minutes_maker import parse_action, parse_transcript, save_minutes_excel


def test_parse_transcript_extracts_core_sections():
    transcript = """
    2026/6/15 開発定例会
    参加者：田中、佐藤
    議題：リリース準備
    決定：6月にβ版を公開する
    田中：資料作成（6/5）
    """

    minutes = parse_transcript(transcript)

    assert minutes.title == "2026/6/15 開発定例会"
    assert minutes.meeting_date == "2026-06-15"
    assert minutes.attendees == ["田中", "佐藤"]
    assert "リリース準備" in minutes.agenda
    assert "6月にβ版を公開する" in minutes.decisions[0]
    assert minutes.action_items[0].owner == "田中"
    assert minutes.action_items[0].task == "資料作成"
    assert minutes.action_items[0].due == "6/5"


def test_extracts_date_from_slash_date_prefixed_title():
    transcript = "2026/6/15 開発定例会"
    minutes = parse_transcript(transcript)

    assert minutes.meeting_date == "2026-06-15"


def test_extracts_date_from_japanese_date_prefixed_title():
    transcript = "2026年6月15日 開発定例会"
    minutes = parse_transcript(transcript)

    assert minutes.meeting_date == "2026-06-15"


def test_parse_action_with_full_width_colon_and_due_date():
    item = parse_action("田中：資料作成（6/5）")

    assert item.owner == "田中"
    assert item.task == "資料作成"
    assert item.due == "6/5"


def test_parse_action_with_owner_after_assignee_label():
    item = parse_action("担当：田中 資料作成（6/5）")

    assert item.owner == "田中"
    assert item.task == "資料作成"
    assert item.due == "6/5"


def test_parse_action_with_owner_after_todo_label():
    item = parse_action("TODO:田中 資料作成（6/5）")

    assert item.owner == "田中"
    assert item.task == "資料作成"
    assert item.due == "6/5"


def test_save_minutes_excel_creates_xlsx_package(tmp_path: Path):
    output = tmp_path / "minutes.xlsx"

    save_minutes_excel("参加者：田中\n決定：次回日程を確定", output)

    assert output.exists()
    with zipfile.ZipFile(output) as archive:
        names = set(archive.namelist())
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")

    assert "[Content_Types].xml" in names
    assert "xl/workbook.xml" in names
    assert "xl/worksheets/sheet1.xml" in names
    assert "議事録" in sheet_xml
    assert "田中" in sheet_xml
