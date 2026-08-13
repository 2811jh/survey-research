import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import pandas as pd
import pytest
from survey_drift import detect_time_col


def test_detect_default_结束答题时间优先级最高():
    df = pd.DataFrame({"结束答题时间": pd.date_range("2026-01-01", periods=3),
                       "提交时间": pd.date_range("2026-02-01", periods=3),
                       "Q1": [1, 2, 3]})
    col, source = detect_time_col(df, None)
    assert col == "结束答题时间"
    assert source == "default"


def test_detect_显式指定覆盖默认():
    df = pd.DataFrame({"结束答题时间": pd.date_range("2026-01-01", periods=3),
                       "其他时间列": pd.date_range("2026-02-01", periods=3)})
    col, source = detect_time_col(df, "其他时间列")
    assert col == "其他时间列"
    assert source == "explicit"


def test_detect_无默认列时按关键词扫描():
    df = pd.DataFrame({"答题日期": pd.date_range("2026-01-01", periods=3),
                       "Q1": [1, 2, 3]})
    col, source = detect_time_col(df, None)
    assert col == "答题日期"
    assert source == "auto_detect"


def test_detect_无任何时间列返回None():
    df = pd.DataFrame({"Q1": [1, 2, 3], "Q2": [4, 5, 6]})
    col, source = detect_time_col(df, None)
    assert col is None
    assert source == "not_found"


def test_detect_关键词命中但非时间类型不误判():
    # 「答题时长」含「答题」但是数值列，不应被识别
    df = pd.DataFrame({"答题时长": [30, 45, 60], "Q1": [1, 2, 3]})
    col, source = detect_time_col(df, None)
    assert col is None
    assert source == "not_found"


import numpy as np
from survey_drift import build_findings


def _make_classification():
    return {
        "single_choice": ["Q1.满意度", "Q2.性别"],
        "multi_choice": {},
        "matrix_scale": {},
        "text": [],
        "meta": [],
        "excluded": [],
        "valid_for_crosstab": ["Q1.满意度", "Q2.性别"],
    }


def test_build_findings_列分桶模式():
    df = pd.DataFrame({
        "Q35.用户版本号": ["v1.0", "v1.0", "v2.0", "v2.0", "v3.0", "v3.0"],
        "Q1.满意度": [5, 4, 3, 2, 1, 5],
        "Q2.性别": [1, 2, 1, 2, 1, 2],
    })
    cls = _make_classification()
    findings = build_findings(df, cls, granularity=None, time_col=None,
                             nps_col=None, satisfaction_cols=None, min_n=2,
                             bucket_col="Q35.用户版本号")
    assert findings["bucket_mode"] == "column"
    assert findings["bucket_col"] == "Q35.用户版本号"
    assert findings["buckets"] == ["v1.0", "v2.0", "v3.0"]
    assert findings["granularity"] is None
    assert "time_col_source" in findings


def test_build_findings_列分桶_显式顺序():
    df = pd.DataFrame({
        "Q35.用户版本号": ["v1.0", "v1.0", "v2.0", "v2.0", "v3.0", "v3.0"],
        "Q1.满意度": [5, 4, 3, 2, 1, 5],
        "Q2.性别": [1, 2, 1, 2, 1, 2],
    })
    cls = _make_classification()
    findings = build_findings(df, cls, granularity=None, time_col=None,
                             nps_col=None, satisfaction_cols=None, min_n=2,
                             bucket_col="Q35.用户版本号",
                             bucket_order=["v3.0", "v2.0", "v1.0"])
    assert findings["buckets"] == ["v3.0", "v2.0", "v1.0"]


def test_build_findings_季度粒度():
    df = pd.DataFrame({
        "结束答题时间": pd.to_datetime(["2026-01-15", "2026-04-15", "2026-07-15", "2026-10-15"]),
        "Q1.满意度": [5, 4, 3, 2],
    })
    cls = {"single_choice": ["Q1.满意度"], "multi_choice": {}, "matrix_scale": {},
           "text": [], "meta": [], "excluded": [], "valid_for_crosstab": ["Q1.满意度"]}
    findings = build_findings(df, cls, granularity="quarter", time_col="结束答题时间",
                              nps_col=None, satisfaction_cols=None, min_n=1,
                              time_col_source="default")
    assert findings["bucket_mode"] == "time"
    assert findings["granularity"] == "quarter"
    assert len(findings["buckets"]) == 4
    assert findings["buckets"][0] == "26年Q1"


def test_build_findings_自定义区间():
    df = pd.DataFrame({
        "结束答题时间": pd.to_datetime(["2026-10-15", "2026-11-12", "2026-11-20", "2026-11-25"]),
        "Q1.满意度": [5, 4, 3, 2],
    })
    cls = {"single_choice": ["Q1.满意度"], "multi_choice": {}, "matrix_scale": {},
           "text": [], "meta": [], "excluded": [], "valid_for_crosstab": ["Q1.满意度"]}
    ranges = [["双11前", "2026-10-01", "2026-11-10"],
              ["双11期", "2026-11-11", "2026-11-13"],
              ["双11后", "2026-11-14", "2026-11-30"]]
    findings = build_findings(df, cls, granularity="custom_ranges", time_col="结束答题时间",
                              nps_col=None, satisfaction_cols=None, min_n=1,
                              custom_ranges=ranges, time_col_source="default")
    assert findings["bucket_mode"] == "time"
    assert findings["granularity"] == "custom_ranges"
    assert "双11前" in findings["buckets"]
    assert "双11期" in findings["buckets"]
    assert "双11后" in findings["buckets"]


def test_time_col_source_字段写入_findings():
    df = pd.DataFrame({
        "结束答题时间": pd.date_range("2026-01-01", periods=3),
        "Q1.满意度": [5, 4, 3],
    })
    cls = {"single_choice": ["Q1.满意度"], "multi_choice": {}, "matrix_scale": {},
           "text": [], "meta": [], "excluded": [], "valid_for_crosstab": ["Q1.满意度"]}
    findings = build_findings(df, cls, granularity="week", time_col="结束答题时间",
                              nps_col=None, satisfaction_cols=None, min_n=1,
                              time_col_source="default")
    assert findings["time_col_source"] == "default"
