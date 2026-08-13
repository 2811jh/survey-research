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
