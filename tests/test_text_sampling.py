import json
import subprocess
import sys
from pathlib import Path

import pandas as pd

SCRIPTS = Path(__file__).parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))
from text_extract import clean_column_texts


def test_default_extract_samples_one_thousand():
    df = pd.DataFrame({"建议": [f"希望优化加载速度编号{i}" for i in range(1200)]})
    result = clean_column_texts(df, "建议")
    assert result["sample_n"] == 1000
    assert result["valid_count"] == 1200
    assert result["texts"] == clean_column_texts(df, "建议")["texts"]


def test_default_extract_keeps_all_below_limit():
    df = pd.DataFrame({"建议": [f"希望优化加载速度编号{i}" for i in range(900)]})
    assert clean_column_texts(df, "建议")["sample_n"] == 900
    assert clean_column_texts(df, "建议", sample_n=0)["sample_n"] == 900


def test_cli_default_samples_one_thousand(tmp_path):
    source = tmp_path / "texts.csv"
    pd.DataFrame({"建议": [f"希望优化加载速度编号{i}" for i in range(1200)]}).to_csv(source, index=False)
    result = subprocess.run([sys.executable, "-X", "utf8", str(SCRIPTS / "text_extract.py"),
                             "--file_path", str(source), "--column", "建议"],
                            stdin=subprocess.DEVNULL, capture_output=True, text=True, encoding="utf-8")
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["sample_n"] == 1000
