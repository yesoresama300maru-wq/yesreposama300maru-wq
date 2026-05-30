from pathlib import Path
import zipfile

from minutes_maker import parse_transcript, save_minutes_excel


def test_parse_transcript_extracts_core_sections():
    transcript = """
    会議名：開発定例会
    開催日：2026年5月30日
    参加者：田中、佐藤
    議題：リリース準備
    決定：6月にβ版を公開する
    担当：田中 リリースノートを作成（6/3）
    """

    minutes = parse_transcript(transcript)

    assert minutes.title == "開発定例会"
    assert minutes.meeting_date == "2026-05-30"
    assert minutes.attendees == ["田中", "佐藤"]
    assert "リリース準備" in minutes.agenda
    assert "6月にβ版を公開する" in minutes.decisions[0]
    assert minutes.action_items[0].owner == "田中"


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
