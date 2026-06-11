from __future__ import annotations

from datetime import date
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

import processor
from tests.conftest import (
    build_attendance_pdf,
    build_attendance_workbook,
    build_basic_target_workbook,
    build_predefined_template_workbook,
    build_source_pdf,
    build_source_workbook_for_dates,
    build_source_workbook,
    build_target_workbook,
)


def _day_column(worksheet, layout: processor.SheetLayout, day_number: int) -> int:
    for col in range(layout.day_start_col, layout.day_end_col + 1):
        if processor.parse_sheet_day_header(worksheet.cell(5, col).value) == day_number:
            return col
    raise AssertionError(f"Day {day_number} was not generated")


def _target_rate_cell(worksheet, layout: processor.SheetLayout):
    blocks = processor.detect_sheet_blocks(worksheet, layout)
    return worksheet.cell(blocks[0].header_row - 1, layout.summary_start_col + 2)


def _workday_count_cell(worksheet, layout: processor.SheetLayout):
    blocks = processor.detect_sheet_blocks(worksheet, layout)
    return worksheet.cell(blocks[0].header_row - 1, layout.summary_start_col + 4)


def test_run_fill_predefined_mode_creates_new_aprilie_sheet_from_asset(tmp_path: Path, monkeypatch) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    preset_path = tmp_path / "preset.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_basic_target_workbook(target_path)
    build_predefined_template_workbook(preset_path)

    monkeypatch.setattr(processor, "PREDEFINED_TEMPLATE_WORKBOOK", preset_path)

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREDEFINED,
        target_month_name="aprilie",
    )

    assert result.output_file.exists()
    assert result.created_sheet_name == "aprilie"
    assert result.template_source_name == "martie"
    assert result.mapped_days == 5
    assert len(result.warnings) == 1
    assert "Zile eliminate din tabel" in result.warnings[0]
    assert "13.04.2026" in result.warnings[0]

    wb = load_workbook(result.output_file, data_only=False)
    ws = wb["aprilie"]
    layout = processor.get_sheet_layout(ws)

    assert layout.day_end_col - layout.day_start_col + 1 == 5
    assert ws.sheet_view.zoomScale == 55
    assert ws["C5"].fill.fgColor.rgb == "00F2E2D1"
    assert ws["G5"].value == 6.04
    assert ws["H5"].value == "VACI TRAN "
    target_rate_cell = _target_rate_cell(ws, layout)
    assert target_rate_cell.value == 166
    assert target_rate_cell.fill.fgColor.rgb == "00DDEBFF"
    assert target_rate_cell.font.bold is True
    assert target_rate_cell.alignment.horizontal == "center"
    assert ws["C6"].value == 120
    assert ws["D6"].value == "=100+12/6"
    assert ws["D7"].value == 14.67
    assert ws["C13"].value == 120
    assert ws["I6"].value == '=COUNTIF(C7:G7,">0")'
    assert ws["J6"].value == "=I6*$J$4"
    workday_count_cell = _workday_count_cell(ws, layout)
    assert ws["L6"].value == f"=${get_column_letter(workday_count_cell.column)}${workday_count_cell.row}"
    assert workday_count_cell.value == 5
    assert workday_count_cell.fill.fgColor.rgb == "00C6EFCE"
    assert workday_count_cell.font.bold is True
    assert workday_count_cell.alignment.horizontal == "center"
    assert wb.sheetnames[0] == "aprilie"
    wb.close()


def test_run_fill_excludes_non_sunday_days_missing_from_source(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    output_dir = tmp_path / "output"
    all_april_days = processor.build_target_day_columns(2026, 4)
    source_days = [
        current_day
        for current_day in all_april_days
        if current_day not in {date(2026, 4, 11), date(2026, 4, 13)}
    ]
    build_source_workbook_for_dates(source_path, source_days)
    build_target_workbook(target_path)

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]
    layout = processor.get_sheet_layout(worksheet)
    header_values = [
        worksheet.cell(5, col).value
        for col in range(layout.day_start_col, layout.day_end_col + 1)
    ]

    assert len(header_values) == 24
    assert 11.04 not in header_values
    assert 13.04 not in header_values
    assert worksheet.cell(5, layout.summary_start_col).value == "VACI TRAN "
    workday_count_cell = _workday_count_cell(worksheet, layout)
    assert worksheet.cell(6, layout.summary_start_col + 4).value == f"=${get_column_letter(workday_count_cell.column)}${workday_count_cell.row}"
    assert workday_count_cell.value == 24
    assert worksheet.cell(6, layout.summary_start_col + 8).value == 9500
    target_rate_cell = _target_rate_cell(worksheet, layout)
    target_rate_ref = f"${get_column_letter(target_rate_cell.column)}${target_rate_cell.row}"
    zile_lucratoare_ref = f"{get_column_letter(layout.summary_start_col + 4)}6"
    bani_in_plus_ref = f"{get_column_letter(layout.summary_start_col + 7)}6"
    salar_baza_ref = f"{get_column_letter(layout.summary_start_col + 8)}6"
    assert target_rate_cell.value == 166
    assert target_rate_cell.fill.fgColor.rgb == "00DDEBFF"
    assert worksheet.cell(6, layout.summary_start_col + 2).value == f"=AB6*{target_rate_ref}"
    assert worksheet.cell(6, layout.summary_start_col + 6).value == f"={salar_baza_ref}/{zile_lucratoare_ref}"
    assert worksheet.cell(6, layout.summary_start_col + 9).value == f"={salar_baza_ref}+{bani_in_plus_ref}"
    assert worksheet.cell(8, layout.summary_start_col).value == "=SUM(C8:Z8)"
    assert worksheet.cell(9, layout.summary_start_col).value == "=SUM(C9:Z9)"
    assert worksheet.cell(10, layout.summary_start_col).value == "=SUM(C10:Z10)"
    assert worksheet.cell(8, layout.summary_start_col + 2).value == "cost /vaca"
    assert worksheet.cell(9, layout.summary_start_col + 2).value == "bani  in plus"
    assert worksheet.cell(8, layout.summary_start_col + 4).value == "=AG6/11.5"
    assert worksheet.cell(9, layout.summary_start_col + 4).value == "=AA10*AE8"
    assert worksheet.cell(10, layout.summary_start_col + 5).value == "=AD10/175"
    assert result.warnings == [
        "Zile eliminate din tabel pentru că lipsesc din sursă: 11.04.2026, 13.04.2026"
    ]

    workbook.close()


def test_run_fill_previous_sheet_mode_overwrites_stale_global_target_rate(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path)

    workbook = load_workbook(target_path, data_only=False)
    worksheet = workbook["martie"]
    layout = processor.get_sheet_layout(worksheet)
    blocks = processor.detect_sheet_blocks(worksheet, layout)
    stale_cell = worksheet.cell(blocks[0].header_row - 1, layout.summary_start_col + 2)
    stale_cell.value = 123
    workbook.save(target_path)
    workbook.close()

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]
    layout = processor.get_sheet_layout(worksheet)
    target_rate_cell = _target_rate_cell(worksheet, layout)
    target_formula = worksheet.cell(6, layout.summary_start_col + 2).value

    assert target_rate_cell.value == 166
    assert target_rate_cell.fill.fgColor.rgb == "00DDEBFF"
    assert target_formula == f"=I6*${get_column_letter(target_rate_cell.column)}${target_rate_cell.row}"
    assert "*166" not in target_formula

    workbook.close()


def test_run_fill_does_not_generate_formula_ranges_that_include_their_own_cell(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    output_dir = tmp_path / "output"
    all_april_days = processor.build_target_day_columns(2026, 4)
    source_days = [
        current_day
        for current_day in all_april_days
        if current_day not in {date(2026, 4, 11), date(2026, 4, 13)}
    ]
    build_source_workbook_for_dates(source_path, source_days)
    build_target_workbook(target_path)

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]

    bad_formulas = []
    ref_errors = []
    for row in worksheet.iter_rows():
        for cell in row:
            formula = cell.value
            if not (isinstance(formula, str) and formula.startswith("=")):
                continue
            if "#REF!" in formula:
                ref_errors.append((cell.coordinate, formula))
            cell_row, cell_col = coordinate_to_tuple(cell.coordinate)
            for match in re.finditer(r"([A-Z]+\d+):([A-Z]+\d+)", formula):
                min_col, min_row, max_col, max_row = range_boundaries(match.group(0))
                if min_row <= cell_row <= max_row and min_col <= cell_col <= max_col:
                    bad_formulas.append((cell.coordinate, formula))

    assert ref_errors == []
    assert bad_formulas == []

    workbook.close()


def test_run_fill_ignores_attendance_co_days_removed_from_generated_table(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    all_april_days = processor.build_target_day_columns(2026, 4)
    source_days = [
        current_day
        for current_day in all_april_days
        if current_day not in {date(2026, 4, 11), date(2026, 4, 13)}
    ]
    build_source_workbook_for_dates(source_path, source_days)
    build_target_workbook(target_path)
    build_attendance_workbook(
        attendance_path,
        rows=[("CIOCLEA IOAN", "11,13", "11.04 - zi eliminată")],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]
    layout = processor.get_sheet_layout(worksheet)
    header_values = [
        worksheet.cell(5, col).value
        for col in range(layout.day_start_col, layout.day_end_col + 1)
    ]

    assert 11.04 not in header_values
    assert 13.04 not in header_values
    assert result.attendance_summary is not None
    assert result.attendance_summary.co_days_applied == 0
    assert result.attendance_summary.mentiuni_days_colored == 0

    workbook.close()


def test_run_fill_predefined_mode_colors_zero_result_block_using_global_off_fallback(
    tmp_path: Path,
    monkeypatch,
) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    preset_path = tmp_path / "preset.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path, zero_days={"02.04.2026"})
    build_basic_target_workbook(target_path)
    build_predefined_template_workbook(preset_path, include_offless_block=True)

    monkeypatch.setattr(processor, "PREDEFINED_TEMPLATE_WORKBOOK", preset_path)

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREDEFINED,
        target_month_name="aprilie",
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]

    for row in (12, 13, 14, 15, 16, 17):
        assert worksheet.cell(row, 4).fill.fgColor.rgb == "00FFF2CC"
    assert worksheet["D13"].value == 0
    assert worksheet["D14"].value == 0

    workbook.close()


def test_run_fill_previous_sheet_mode_recreates_existing_target_and_inserts_before_template(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path, include_existing_aprilie=True)

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
    )

    wb = load_workbook(result.output_file, data_only=False)
    assert wb.sheetnames[:2] == ["aprilie", "martie"]

    ws = wb["aprilie"]
    assert ws["A1"].value is None
    assert ws["G12"].value == 6.04
    assert ws["H5"].value == "VACI TRAN "
    assert ws["Q6"].value == "=P6+O6"
    wb.close()


def test_run_fill_previous_sheet_mode_keeps_normal_fill_for_zero_result_blocks(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path, zero_days={"02.04.2026"})
    build_target_workbook(target_path, include_offless_block=True)

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]

    assert worksheet["D13"].value == 0
    assert worksheet["D14"].value == 0
    assert worksheet["D12"].fill.fgColor.rgb == "00F2E2D1"
    for cell_ref in ("D13", "D14", "D15", "D16", "D17"):
        assert worksheet[cell_ref].fill.fgColor.rgb == "00FFF9F2"

    workbook.close()


def test_run_fill_with_attendance_applies_co_days_and_mentiuni(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path)
    build_attendance_workbook(
        attendance_path,
        rows=[("CIOCLEA IOAN", "02,06", "07.04 - a plecat la 15:00")],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]

    assert worksheet["D6"].value == 0
    assert worksheet["D7"].value == 0
    assert worksheet["G6"].value == 0
    assert worksheet["G7"].value == 0
    assert worksheet["D8"].value == "=IF(D7>0,D6/D7,0)"
    assert worksheet["D9"].value == "=IF(D7>0,11.5,0)"
    assert worksheet["D10"].value == "=IF(D7>0,D8-D9,0)"
    assert worksheet["D6"].fill.fgColor.rgb == "00FFF2CC"
    assert worksheet["T5"].value == "Mentiuni"
    assert worksheet["T6"].value == "07.04 - a plecat la 15:00"
    assert worksheet["T6"].fill.fgColor.rgb == "00FFF2CC"
    assert worksheet["T13"].value is None
    assert worksheet.row_dimensions[6].height >= 18
    assert result.attendance_summary is not None
    assert result.attendance_summary.matched_employees == 1
    assert result.attendance_summary.co_days_applied == 2
    assert result.attendance_summary.mentiuni_days_colored == 0
    assert result.attendance_summary.mentiuni_copied == 1

    workbook.close()


def test_run_fill_with_attendance_applies_n_days_like_co(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    source_days = [
        date(2026, 4, 1),
        date(2026, 4, 2),
        date(2026, 4, 3),
        date(2026, 4, 4),
        date(2026, 4, 6),
    ]
    build_source_workbook_for_dates(source_path, source_days)
    build_target_workbook(target_path)
    build_attendance_workbook(
        attendance_path,
        rows=[("CIOCLEA IOAN", "02", "03", "04.04 - observatie")],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]
    layout = processor.get_sheet_layout(worksheet)
    co_col = _day_column(worksheet, layout, 2)
    n_col = _day_column(worksheet, layout, 3)
    mentiuni_col = _day_column(worksheet, layout, 4)

    assert worksheet.cell(6, co_col).value == 0
    assert worksheet.cell(7, co_col).value == 0
    assert worksheet.cell(6, n_col).value == 0
    assert worksheet.cell(7, n_col).value == 0
    for row in (5, 6, 7, 8, 9, 10):
        assert worksheet.cell(row, n_col).fill.fgColor.rgb == "00FFF2CC"
        assert worksheet.cell(row, mentiuni_col).fill.fgColor.rgb == "00FFF8D6"
    assert result.attendance_summary is not None
    assert result.attendance_summary.co_days_applied == 1
    assert result.attendance_summary.n_days_applied == 1
    assert result.attendance_summary.mentiuni_days_colored == 1

    workbook.close()


def test_run_fill_accepts_pdf_source_and_pdf_attendance(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.pdf"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.pdf"
    output_dir = tmp_path / "output"
    build_source_pdf(source_path)
    build_target_workbook(target_path)
    build_attendance_pdf(
        attendance_path,
        rows=[("CIOCLEA IOAN", "02", "03", "04.04 - observatie")],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]
    layout = processor.get_sheet_layout(worksheet)
    n_col = _day_column(worksheet, layout, 3)

    assert worksheet.cell(6, n_col).value == 0
    assert worksheet.cell(7, n_col).value == 0
    assert result.mapped_days == 5
    assert result.attendance_summary is not None
    assert result.attendance_summary.n_days_applied == 1

    workbook.close()


def test_run_fill_with_attendance_colors_mentiuni_day_pale_yellow_without_changing_values(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    source_days = [
        date(2026, 4, 1),
        date(2026, 4, 2),
        date(2026, 4, 3),
        date(2026, 4, 4),
        date(2026, 4, 6),
        date(2026, 4, 7),
    ]
    build_source_workbook_for_dates(source_path, source_days)
    build_target_workbook(target_path)
    build_attendance_workbook(
        attendance_path,
        rows=[("CIOCLEA IOAN", None, "07.04 - a plecat la 15:00")],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]
    layout = processor.get_sheet_layout(worksheet)
    day_col = _day_column(worksheet, layout, 7)
    day_letter = get_column_letter(day_col)

    assert worksheet.cell(6, day_col).value != 0
    assert worksheet.cell(7, day_col).value != 0
    assert worksheet.cell(8, day_col).value == f"=IF({day_letter}7>0,{day_letter}6/{day_letter}7,0)"
    assert worksheet.cell(9, day_col).value == f"=IF({day_letter}7>0,11.5,0)"
    assert worksheet.cell(10, day_col).value == f"=IF({day_letter}7>0,{day_letter}8-{day_letter}9,0)"
    for row in (5, 6, 7, 8, 9, 10):
        assert worksheet.cell(row, day_col).fill.fgColor.rgb == "00FFF8D6"

    mentiuni_col = layout.summary_start_col + len(processor.SUMMARY_LABELS)
    assert worksheet.cell(5, mentiuni_col).value == "Mentiuni"
    assert worksheet.cell(6, mentiuni_col).value == "07.04 - a plecat la 15:00"
    assert result.attendance_summary is not None
    assert result.attendance_summary.mentiuni_days_colored == 1
    assert result.attendance_summary.mentiuni_copied == 1

    workbook.close()


def test_run_fill_with_attendance_keeps_co_fill_when_co_and_mentiuni_overlap(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    source_days = [
        date(2026, 4, 1),
        date(2026, 4, 2),
        date(2026, 4, 3),
        date(2026, 4, 4),
        date(2026, 4, 6),
        date(2026, 4, 7),
    ]
    build_source_workbook_for_dates(source_path, source_days)
    build_target_workbook(target_path)
    build_attendance_workbook(
        attendance_path,
        rows=[("CIOCLEA IOAN", "07", "07.04 - CO și mențiune")],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]
    layout = processor.get_sheet_layout(worksheet)
    day_col = _day_column(worksheet, layout, 7)

    assert worksheet.cell(6, day_col).value == 0
    assert worksheet.cell(7, day_col).value == 0
    for row in (5, 6, 7, 8, 9, 10):
        assert worksheet.cell(row, day_col).fill.fgColor.rgb == "00FFF2CC"
    assert result.attendance_summary is not None
    assert result.attendance_summary.co_days_applied == 1
    assert result.attendance_summary.mentiuni_days_colored == 0

    workbook.close()


def test_run_fill_with_attendance_inserts_mentiuni_when_neighbor_column_is_not_blank(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path)
    build_attendance_workbook(attendance_path)

    workbook = load_workbook(target_path, data_only=False)
    worksheet = workbook["martie"]
    layout = processor.get_sheet_layout(worksheet)
    premium_col = layout.summary_start_col + len(processor.SUMMARY_LABELS) - 1
    worksheet.cell(5, premium_col + 1).value = "RESERVED"
    workbook.save(target_path)
    workbook.close()

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]

    assert worksheet["T5"].value == "Mentiuni"
    assert worksheet["U5"].value == "RESERVED"

    workbook.close()


def test_run_fill_with_attendance_uses_approximate_matching_and_logs_unmatched(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path)

    workbook = load_workbook(target_path, data_only=False)
    worksheet = workbook["martie"]
    worksheet["B13"] = "NECHIFOR NELU FLORIN"
    workbook.save(target_path)
    workbook.close()

    build_attendance_workbook(
        attendance_path,
        rows=[
            ("NICHIFOR NELU-FLORIN", "02", "notă aproximativă"),
            ("PERSOANA LIPSA", "02", None),
        ],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    workbook = load_workbook(result.output_file, data_only=False)
    worksheet = workbook["aprilie"]

    assert worksheet["D13"].value == 0
    assert worksheet["T13"].value == "notă aproximativă"
    assert result.attendance_summary is not None
    assert result.attendance_summary.approximate_matches == 1
    assert result.attendance_summary.unmatched_names == ["PERSOANA LIPSA"]

    workbook.close()


def test_run_fill_with_attendance_logs_ambiguous_names(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path)

    workbook = load_workbook(target_path, data_only=False)
    worksheet = workbook["martie"]
    worksheet["B6"] = "POPESCU ION VASILE"
    worksheet["B13"] = "IONESCU ION VASILE"
    workbook.save(target_path)
    workbook.close()

    build_attendance_workbook(
        attendance_path,
        rows=[("ION VASILE", "02", "notă ambiguă")],
    )

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
    )

    assert result.attendance_summary is not None
    assert len(result.attendance_summary.ambiguous_names) == 1
    assert "POPESCU ION VASILE" in result.attendance_summary.ambiguous_names[0]
    assert "IONESCU ION VASILE" in result.attendance_summary.ambiguous_names[0]


def test_run_fill_with_attendance_month_mismatch_blocks_by_default_and_can_continue(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    attendance_path = tmp_path / "transatori_detinuti.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path)
    build_attendance_workbook(attendance_path, title="TRANŞARE - MARTIE 2026")

    try:
        processor.run_fill(
            source_path=source_path,
            target_path=target_path,
            output_dir=output_dir,
            template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
            target_month_name="aprilie",
            template_sheet_name="martie",
            attendance_path=attendance_path,
        )
    except processor.ProcessorError as exc:
        assert "Luna din Transatori+Detinuți" in str(exc)
    else:
        raise AssertionError("Expected ProcessorError")

    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
        attendance_path=attendance_path,
        allow_attendance_month_mismatch=True,
    )

    assert result.output_file.exists()


def test_run_fill_validates_source_month_mismatch(tmp_path: Path) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path)

    try:
        processor.run_fill(
            source_path=source_path,
            target_path=target_path,
            output_dir=output_dir,
            template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
            target_month_name="martie",
            template_sheet_name="martie",
        )
    except processor.ProcessorError as exc:
        message = str(exc)
        assert "aprilie" in message and "martie" in message
    else:
        raise AssertionError("Expected ProcessorError")


def test_run_fill_proceeds_with_selected_month_when_detection_fails(tmp_path: Path, monkeypatch) -> None:
    source_path = tmp_path / "situatie.xlsx"
    target_path = tmp_path / "salarii.xlsx"
    output_dir = tmp_path / "output"
    build_source_workbook(source_path)
    build_target_workbook(target_path, include_existing_aprilie=True)

    def _detection_fails(*args, **kwargs):
        raise processor.ProcessorError("Nu am putut detecta luna din fișierul sursă.")

    monkeypatch.setattr(processor, "detect_source_month_from_file", _detection_fails)

    # Even when automatic month detection is broken, the user's explicit month choice
    # must drive the run instead of aborting.
    result = processor.run_fill(
        source_path=source_path,
        target_path=target_path,
        output_dir=output_dir,
        template_mode=processor.TEMPLATE_MODE_PREVIOUS_SHEET,
        target_month_name="aprilie",
        template_sheet_name="martie",
    )

    assert result.output_file.exists()
    assert result.created_sheet_name == "aprilie"
