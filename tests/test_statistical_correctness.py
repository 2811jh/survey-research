import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parents[1] / "scripts"))
import crosstab as ct
import survey_drift as drift
from load_and_classify import classify_columns


def test_plain_names_basic_stats_structure():
    import basic_stats
    questions = basic_stats._build_question_structure(pd.DataFrame({"性别": ["男", "女"], "满意度": [4, 5], "uid": [1, 2]}))
    assert [q["label"] for q in questions] == ["性别", "满意度"]


def test_nps_detection_does_not_require_observed_promoters():
    df = pd.DataFrame({"Q1.NPS推荐": [0, 6, 7, 8], "Q2.性别": ["M", "M", "F", "F"]})
    assert ct._is_scoreable_question("Q1.NPS推荐", df) == "nps"


def test_basic_nps_excludes_invalid_values():
    import basic_stats
    result = basic_stats._stat_nps(pd.DataFrame({"Q1.NPS": [-1, 11, 9.5, 10, 8, 0]}), "Q1.NPS")
    assert result["total"] == 3
    assert result["nps"] == 0


def test_significance_keeps_question_and_option():
    df = pd.DataFrame({
        "Q1.A": [1] * 40 + [2] * 10 + [1] * 10 + [2] * 40,
        "Q2.B": [1] * 10 + [2] * 40 + [1] * 40 + [2] * 10,
        "Q3.性别": ["M"] * 50 + ["F"] * 50,
    })
    result = ct.run_crosstab(df, classify_columns(df), ["all"], ["Q3.性别"])
    sig = ct.calc_significance(result)
    assert sig["Q3.性别"]["M"]["Q1.A"]["1"]["delta_pp"] == 30
    assert sig["Q3.性别"]["M"]["Q2.B"]["1"]["delta_pp"] == -30
    heat = ct.calc_heatmap_data(result, sig)
    assert heat.at[("Q1.A", "1"), "Q3.性别\nM"] == 30
    summary = ct.get_crosstab_summary(result, sig)
    assert summary["Q3.性别"]["question"] in ("Q1.A", "Q2.B")


def test_significance_uses_valid_answers_and_complement():
    df = pd.DataFrame({"Q1.A": [1] * 10 + [None] * 40 + [1] * 10 + [2] * 40,
                       "Q2.性别": ["M"] * 50 + ["F"] * 50})
    result = ct.run_crosstab(df, classify_columns(df), ["Q1.A"], ["Q2.性别"])
    info = ct.calc_significance(result)["Q2.性别"]["M"]["Q1.A"]["1.0"]
    assert info["delta_pp"] == pytest.approx(66.7)
    assert info["p"] == ct.two_prop_z(10, 10, 10, 50)[1]
    assert info["significant"]


def test_nps_detects_detractor_shift():
    df = pd.DataFrame({"Q1.推荐": [10] * 20 + [8] * 80 + [10] * 20 + [0] * 80,
                       "version": ["old"] * 100 + ["new"] * 100})
    result = drift.build_findings(df, classify_columns(df), None, None,
                                  "Q1.推荐", [], bucket_col="version")
    change = result["metrics"][0]["adjacent"][0]
    assert change["delta_pp"] == -80
    assert change["p"] < .05
    assert change["drift"]


def test_custom_range_includes_end_day():
    dates = pd.Series(pd.to_datetime(["2026-11-13 23:59:59", "2026-11-14 00:00:00"]))
    labels, _ = drift._bucketize_custom(dates, [["event", "2026-11-11", "2026-11-13"]])
    assert labels.tolist() == ["event", None]


def test_week_buckets_are_unique_across_years():
    _, buckets = drift.bucketize(pd.Series(pd.to_datetime(["2020-07-06", "2026-07-06"])), "week")
    assert len(buckets) == 2
    assert "2020" in buckets[0]
    assert "2026" in buckets[1]


def test_missing_column_buckets_are_excluded():
    df = pd.DataFrame({"Q1.满意度": [4, 5, 4, 5], "version": ["a", None, "", "  "]})
    result = drift.build_findings(df, classify_columns(df), None, None, None,
                                  ["Q1.满意度"], bucket_col="version")
    assert result["buckets"] == ["a"]
    assert result["bucket_sizes"] == {"a": 1}


def test_nominal_codes_do_not_become_mean_metrics():
    df = pd.DataFrame({"Q1.年龄": [1, 2, 3, 4, 5] * 2, "version": ["a"] * 5 + ["b"] * 5})
    result = drift.build_findings(df, classify_columns(df), None, None, None, [], bucket_col="version")
    assert result["metrics"] == []
    assert ct.auto_detect_score_questions(df, {"valid_rows_map": {"Q1.年龄": "single"}}) == []


def test_plain_column_names_are_analyzable():
    result = classify_columns(pd.DataFrame({"性别": ["男", "女"], "满意度": [4, 5], "uid": [1, 2]}))
    assert result["single_choice"] == ["性别", "满意度"]
    assert result["meta"] == ["uid"]


def test_degenerate_mean_test_detects_different_constants():
    p, delta = ct._mean_t_test([5] * 40, [3] * 40)
    assert p < .05
    assert delta == 2


def test_nps_ignores_invalid_scores():
    result = drift.compute_nps(pd.Series([-1, 11, 9.5, 10, 8, 0]))
    assert result["n"] == 3
    assert result["nps"] == 0


def test_basic_stats_nps_without_ten_and_non_nps_ten(tmp_path):
    import basic_stats
    from openpyxl import load_workbook
    for name, values, expected in [
        ("Q1.推荐给朋友的可能性（量表题）", [0, 6, 8, 9], True),
        ("Q1.画面评分（量表题）", [1, 5, 9, 10], False),
    ]:
        df = pd.DataFrame({name: values})
        output = tmp_path / ("nps.xlsx" if expected else "rating.xlsx")
        basic_stats._write_stat_report(df, basic_stats._build_question_structure(df), str(output))
        book = load_workbook(output)
        cells = [cell.value for row in book.active for cell in row]
        assert ("NPS净推荐值" in cells) is expected
        book.close()


def test_excel_marks_each_group_without_turning_numbers_into_text(tmp_path):
    from openpyxl import Workbook
    df = pd.DataFrame({"Q1.A": [1] * 40 + [2] * 10 + [1] * 10 + [2] * 40,
                       "Q2.性别": ["M"] * 50 + ["F"] * 50})
    result = ct.run_crosstab(df, classify_columns(df), ["Q1.A"], ["Q2.性别"])
    book = Workbook()
    sheet = book.active
    sheet.append(["题目", "选项"] + result["col_labels"])
    for idx, row in result["percent_df"].iterrows():
        sheet.append(list(idx) + row.tolist())
    ct._apply_significance_heatmap(sheet, result["percent_df"], result["col_labels"], ct.calc_significance(result))
    columns = {label: i + 3 for i, label in enumerate(result["col_labels"])}
    for group in ("M", "F"):
        cell = sheet.cell(2, columns[f"Q2.性别\n{group}"])
        assert isinstance(cell.value, float)
        assert cell.comment is not None
    assert sheet.cell(2, columns["Q2.性别\nM"]).font.color.rgb != sheet.cell(2, columns["Q2.性别\nF"]).font.color.rgb
