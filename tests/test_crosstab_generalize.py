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


from crosstab import calc_scores
import pandas as pd


def test_calc_scores_含样本量行():
    """每个量表题的得分行下方应紧跟样本量行。"""
    df = pd.DataFrame({
        "Q1.满意度": [5, 4, 3, 4, 5, 2, 4, 5, 3, 4],
        "Q33.性别": ["男", "女", "男", "女", "男", "女", "男", "女", "男", "女"],
    })
    # 构造最小 ct_result：freq_df 用真实 crosstab
    freq = pd.crosstab(
        df["Q1.满意度"].astype(str),
        df["Q33.性别"],
        margins=True, margins_name="Q33.性别\n总计",
    )
    freq.index = pd.MultiIndex.from_arrays(
        [["Q1.满意度"] * len(freq), [str(x) for x in freq.index]],
        names=["题目", "选项"],
    )
    ct_result = {
        "freq_df": freq,
        "valid_rows_map": {"Q1.满意度": "single"},
        "col_labels": list(freq.columns),
        "col_totals": {c: int(freq[c].sum()) for c in freq.columns},
    }
    score_df = calc_scores(df, ct_result, ["Q1.满意度"])
    assert score_df is not None
    indices = [str(idx) for idx in score_df.index]
    assert any("得分" in i for i in indices), f"expected score row, got: {indices}"
    assert any("样本量" in i for i in indices), f"expected sample size row, got: {indices}"
    # 样本量行应该在得分行之后
    score_idx = next(i for i, x in enumerate(indices) if "得分" in x)
    sample_idx = next(i for i, x in enumerate(indices) if "样本量" in x)
    assert sample_idx == score_idx + 1, f"sample row should follow score row: {indices}"


from crosstab import calc_significance, two_prop_z


def test_two_prop_z_基本():
    z, p = two_prop_z(60, 100, 50, 100)
    assert abs(z) > 1.4
    assert p < 0.2


def test_two_prop_z_无差异():
    z, p = two_prop_z(50, 100, 50, 100)
    assert z == 0.0
    assert p == 1.0


def test_calc_significance_vs_分组维度总计():
    """构造已知差异：男组某选项占比 vs 性别总计占比 差 10pp，样本量足够显著。"""
    import pandas as pd
    freq = pd.DataFrame(
        {
            "Q33.性别\n男": [120, 80, 200],
            "Q33.性别\n女": [80, 120, 200],
            "Q33.性别\n总计": [200, 200, 400],
        },
        index=pd.MultiIndex.from_arrays(
            [["Q1"] * 3, ["选项A", "选项B", "总计"]],
            names=["题目", "选项"],
        ),
    )
    ct_result = {
        "freq_df": freq,
        "col_labels": ["Q33.性别\n男", "Q33.性别\n女", "Q33.性别\n总计"],
        "col_totals": {
            "Q33.性别\n男": 200,
            "Q33.性别\n女": 200,
            "Q33.性别\n总计": 400,
        },
    }
    sig = calc_significance(ct_result)
    # 男 vs 总计：选项A 60% vs 50%，差 10pp，应显著
    assert "Q33.性别" in sig
    assert "男" in sig["Q33.性别"]
    assert "选项A" in sig["Q33.性别"]["男"]
    info = sig["Q33.性别"]["男"]["选项A"]
    assert info["significant"] is True
    assert info["direction"] == "up"
    assert abs(info["delta_pp"] - 10.0) < 0.1


from crosstab import run_crosstab_pipeline


def test_run_crosstab_pipeline_auto_返回候选():
    df = pd.DataFrame({
        "Q1.满意度": [5, 4, 3, 4, 5],
        "Q33.请问您的性别是？": ["男", "女", "男", "女", "男"],
        "Q34.请问您的年龄是？": ["18-24", "25-30", "18-24", "25-30", "18-24"],
    })
    df.to_csv("/tmp/test_crosstab_auto.csv", index=False, encoding="utf-8-sig")
    result = run_crosstab_pipeline(
        file_path="/tmp/test_crosstab_auto.csv",
        row_questions=["all"],
        col_questions=["auto"],
        calc_scores_mode="auto",
    )
    assert result["status"] == "need_input"
    assert result["reason"] == "col_candidates"
    assert "Q33.请问您的性别是？" in result["candidates"]
    assert "Q34.请问您的年龄是？" in result["candidates"]
