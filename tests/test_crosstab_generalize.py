import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import pandas as pd
import pytest
from load_and_classify import identify_demographic_cols


def test_identify_demographic_识别性别年龄职业():
    df = pd.DataFrame({
        "Q1.满意度": [4, 5, 3],
        "Q33.请问您的性别是？": ["男", "女", "男"],
        "Q34.请问您的年龄是？": ["18-24", "25-30", "18-24"],
        "Q35.请问您的职业是？": ["学生", "工作", "学生"],
        "Q50.付费等级": ["免费", "付费", "免费"],
    })
    classification = {
        "single_choice": ["Q1.满意度", "Q33.请问您的性别是？", "Q34.请问您的年龄是？",
                          "Q35.请问您的职业是？", "Q50.付费等级"],
        "multi_choice": {}, "matrix_scale": {}, "text": [],
        "meta": [], "excluded": [], "valid_for_crosstab": [],
    }
    candidates = identify_demographic_cols(df, classification)
    assert "Q33.请问您的性别是？" in candidates
    assert "Q34.请问您的年龄是？" in candidates
    assert "Q35.请问您的职业是？" in candidates
    assert "Q50.付费等级" in candidates
    assert "Q1.满意度" not in candidates


def test_identify_demographic_无人口学题返回空():
    df = pd.DataFrame({"Q1.满意度": [4, 5], "Q2.玩法偏好": [1, 2]})
    classification = {
        "single_choice": ["Q1.满意度", "Q2.玩法偏好"],
        "multi_choice": {}, "matrix_scale": {}, "text": [],
        "meta": [], "excluded": [], "valid_for_crosstab": [],
    }
    candidates = identify_demographic_cols(df, classification)
    assert candidates == []


from crosstab import _short_col_label, default_output_filename, _five_point_scale_series


def test_short_col_label_提取关键词():
    assert _short_col_label("Q33.请问您的性别是？") == "性别"
    assert _short_col_label("Q34.请问您的年龄是？") == "年龄"
    assert _short_col_label("Q35.请问您的职业是？") == "职业"
    assert _short_col_label("Q50.付费等级") == "付费"


def test_short_col_label_无关键词截断():
    assert _short_col_label("Q10.您最喜欢的玩法") == "您最喜欢的玩法"


def test_default_output_filename_单分组():
    name = default_output_filename(["Q33.请问您的性别是？"], "survey_123_数据.csv")
    assert name == "survey_123_数据_交叉分析_按性别.xlsx"


def test_default_output_filename_多分组():
    cols = ["Q33.请问您的性别是？", "Q34.请问您的年龄是？", "Q35.请问您的职业是？"]
    name = default_output_filename(cols, "survey_123_数据.csv")
    assert name == "survey_123_数据_交叉分析_按性别_年龄_职业.xlsx"


def test_five_point_scale_series_识别():
    s = pd.Series([5, 4, 3, 2, 1, 5, 4, 4, 3, 5])
    assert _five_point_scale_series(s) is True


def test_five_point_scale_series_排除二元():
    s = pd.Series([1, 2, 1, 2, 1])
    assert _five_point_scale_series(s) is False


def test_five_point_scale_series_排除NPS():
    s = pd.Series([0, 1, 5, 7, 8, 9, 10, 3, 6, 4])
    assert _five_point_scale_series(s) is False


from crosstab import auto_detect_score_questions


def test_auto_detect_score_纳入五点量表题():
    """非满意度/NPS 关键词，但取值 1-5 的题应被识别。"""
    df = pd.DataFrame({
        "Q1.满意度": [5, 4, 3, 4, 5, 2, 1, 4, 5, 3],
        "Q13.整体印象": [4, 3, 5, 4, 2, 3, 4, 5, 1, 4],
        "Q2.性别": [1, 2, 1, 2, 1, 2, 1, 2, 1, 2],
    })
    ct_result = {
        "valid_rows_map": {"Q1.满意度": "single", "Q13.整体印象": "single", "Q2.性别": "single"},
    }
    scoreable = auto_detect_score_questions(df, ct_result)
    assert "Q1.满意度" in scoreable
    assert "Q13.整体印象" in scoreable
    assert "Q2.性别" not in scoreable
