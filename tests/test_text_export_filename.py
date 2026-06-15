import importlib.util
from pathlib import Path

MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "text_export.py"
spec = importlib.util.spec_from_file_location("text_export", MODULE_PATH)
text_export = importlib.util.module_from_spec(spec)
spec.loader.exec_module(text_export)


def test_default_output_filename_uses_question_summary():
    result = [{"question": "Q5.【可跳过】为什么您对MC移动版比较满意呢？"}]

    assert text_export.default_output_filename(result) == "Q5_MC满意原因.xlsx"


def test_default_output_filename_handles_activity_suggestions():
    result = [{"question": "Q70.【可跳过】您对“五一特惠节”类似的活动还有哪些建议或期待？"}]

    assert text_export.default_output_filename(result) == "Q70_五一特惠建议.xlsx"


def test_default_output_filename_falls_back_to_text_analysis():
    result = [{"question": "Q101.其他开放题"}]

    assert text_export.default_output_filename(result) == "Q101_文本分析.xlsx"
