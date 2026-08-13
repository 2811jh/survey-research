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
