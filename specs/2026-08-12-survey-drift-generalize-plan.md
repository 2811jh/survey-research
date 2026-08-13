# 异动分析通用化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 survey-research 阶段 6 异动诊断从「回流问卷专用」泛化为通用问卷时序/分桶对比能力，并打通下载/清洗 → 异动诊断的文档衔接。

**Architecture:** 改动分为三层：(1) `scripts/survey_drift.py` 增加分桶模式双形态 + 时间列自动检测 + 新粒度；(2) `references/18-drift-workflow.md` 全篇中性化与新增模式文档；(3) SKILL.md 与 09/10 references 的衔接补全。所有现有命令行为保持向后兼容。

**Tech Stack:** Python 3 + pandas + scipy + openpyxl + argparse CLI；Markdown 文档；YAML frontmatter。

**Spec:** `survey-research/specs/2026-08-12-survey-drift-generalize-design.md`

---

## File Structure

| 文件 | 责任 | 改动类型 |
|------|------|---------|
| `scripts/survey_drift.py` | 异动诊断 CLI 主脚本 | 修改 |
| `references/18-drift-workflow.md` | 异动诊断工作流文档 | 大改 |
| `references/09-survey-download.md` | 下载文档 | 小改（加后续操作）|
| `references/10-survey-clean.md` | 清洗文档 | 小改（加后续操作）|
| `SKILL.md` | 主入口 | 修改（阶段 6 标题、路由 B 衔接、description）|
| `tests/test_survey_drift_generalize.py` | 新增能力测试 | 新建 |

---

## Task 1: 输出名与脚本注释中性化

**Files:**
- Modify: `scripts/survey_drift.py:3-13, 1065-1068`

- [ ] **Step 1: 改脚本顶部注释**

`scripts/survey_drift.py:3-13` 现状：
```python
"""
问卷分析工具 - 时间异动诊断 (survey_drift)
==========================================

单份回流问卷按 周/月/天 分桶，逐题相邻期显著性检验，双门槛判异动，
Agent 写一句话结论，导出 4-Sheet Excel。

子命令:
    analyze  分桶 + 检验 → drift_findings.json
    export   findings + conclusions → Excel
"""
```

改为：
```python
"""
问卷分析工具 - 异动诊断 (survey_drift)
==========================================

按 周/月/天/季度/自定义区间/任意列 分桶，逐题相邻期显著性检验，
双门槛判异动，Agent 写一句话结论，导出 4-Sheet Excel。

子命令:
    analyze  分桶 + 检验 → drift_findings.json
    export   findings + conclusions → Excel
"""
```

- [ ] **Step 2: 改 default_output_filename 函数**

`scripts/survey_drift.py:1065-1068` 现状：
```python
def default_output_filename(granularity):
    label = {"week": "按周", "month": "按月", "day": "按天"}.get(granularity, granularity)
    from datetime import datetime
    return f"回流异动诊断_{label}_{datetime.now():%Y%m%d_%H%M}.xlsx"
```

改为：
```python
def default_output_filename(granularity, bucket_col=None):
    if bucket_col:
        # 列分桶模式：用列名简化作 label
        label = _short_col_label(bucket_col)
    else:
        label = {"week": "按周", "month": "按月", "day": "按天",
                 "quarter": "按季度", "custom_ranges": "按自定义区间"}.get(granularity, granularity)
    from datetime import datetime
    return f"问卷异动诊断_{label}_{datetime.now():%Y%m%d_%H%M}.xlsx"


def _short_col_label(col):
    """从列名提取简短 label：Q35.用户版本号 → 用户版本号；无前缀取整列。"""
    s = str(col)
    # 去掉 Q\d+. 前缀
    m = re.match(r"Q\d+\.\s*(.+)", s)
    return m.group(1) if m else s
```

- [ ] **Step 3: 验证未破坏现有调用**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -c "from scripts.survey_drift import default_output_filename; print(default_output_filename('week')); print(default_output_filename('quarter')); print(default_output_filename('custom_ranges')); print(default_output_filename('week', 'Q35.用户版本号'))"
```

Expected output:
```
问卷异动诊断_按周_YYYYMMDD_HHMM.xlsx
问卷异动诊断_按季度_YYYYMMDD_HHMM.xlsx
问卷异动诊断_按自定义区间_YYYYMMDD_HHMM.xlsx
问卷异动诊断_用户版本号_YYYYMMDD_HHMM.xlsx
```

- [ ] **Step 4: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/survey_drift.py && git commit -m "refactor(drift): neutralize output filename and script header for generalization"
```

---

## Task 2: 时间列自动检测

**Files:**
- Modify: `scripts/survey_drift.py:274-282` (load_df 附近新增 detect_time_col)
- Modify: `scripts/survey_drift.py:1071-1100` (_cmd_analyze)
- Modify: `scripts/survey_drift.py:1134` (argparse default)
- Test: `tests/test_survey_drift_generalize.py`

- [ ] **Step 1: 写 detect_time_col 失败测试**

Create `tests/test_survey_drift_generalize.py`:
```python
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
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_survey_drift_generalize.py -v
```

Expected: FAIL with `ImportError: cannot import name 'detect_time_col'`

- [ ] **Step 3: 实现 detect_time_col**

在 `scripts/survey_drift.py` 的 `load_df` 函数（line 274-281）之后插入：
```python
_TIME_COL_KEYWORDS = ("时间", "日期", "date", "time", "提交", "答题")


def _is_parseable_as_datetime(series, sample_size=50):
    """抽样检查列是否可解析为时间。纯数值列（如答题时长秒数）应被排除。"""
    non_null = series.dropna()
    if len(non_null) == 0:
        return False
    sample = non_null.sample(min(sample_size, len(non_null)), random_state=42)
    parsed = pd.to_datetime(sample, errors="coerce")
    # 解析成功率 ≥ 80% 视为时间列
    return parsed.notna().mean() >= 0.8


def detect_time_col(df, explicit):
    """时间列自动检测。返回 (col_name, source)。
    source 取值：explicit / default / auto_detect / not_found。"""
    if explicit::
            return explicit, "explicit"
        return None, "not_found"
    # 优先级 1：默认列
    if "结束答题时间" in df.columns and _is_parseable_as_datetime(df["结束答题时间"]):
        return "结束答题时间", "default"
    # 优先级 2：关键词扫描
    candidates = []
    for col in df.columns:
        s = str(col).lower()
        if any(kw in s for kw in _TIME_COL_KEYWORDS):
            if _is_parseable_as_datetime(df[col]):
                candidates.append(col)
    if candidates:
        # 取第一个命中的（按列顺序）
        return candidates[0], "auto_detect"
    return None, "not_found"
```

- [ ] **Step 4: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_survey_drift_generalize.py -v
```

Expected: 5 passed

- [ ] **Step 5: 接入 _cmd_analyze**

`scripts/survey_drift.py:1071-1100` 现状 `_cmd_analyze` 开头：
```python
def _cmd_analyze(args):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from load_and_classify import classify_columns
    df = load_df(args.file_path)
    if args.time_col not in df.columns:
        return {"status": "need_input", "reason": "time_col_missing",
                "message": f"时间列 '{args.time_col}' 不存在，可用列：{list(df.columns[:20])}"}
    classification = classify_columns(df)
```

改为：
```python
def _cmd_analyze(args):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from load_and_classify import classify_columns
    df = load_df(args.file_path)
    # 模式 B：列分桶（非时间维度），跳过时间列检测
    if args.bucket_col:
        if args.bucket_col not in df.columns:
            return {"status": "need_input", "reason": "bucket_col_missing",
                    "message": f"分桶列 '{args.bucket_col}' 不存在，可用列：{list(df.columns[:20])}"}
        time_col = None
        time_col_source = "not_applicable"
    else:
        time_col, time_col_source = detect_time_col(df, args.time_col)
        if time_col is None:
            return {"status": "need_input", "reason": "time_col_missing",
                    "message": f"未找到时间列，可用列：{list(df.columns[:20])}；请用 --time_col 指定"}
    classification = classify_columns(df)
```

继续替换 build_findings 调用部分（line 1085-1100）：
```python
    findings = build_findings(df, classification, args.granularity, time_col,
                              nps_col, sat_cols, args.min_n,
                              bucket_col=args.bucket_col, bucket_order=args.bucket_order,
                              custom_ranges=args.custom_ranges, time_col_source=time_col_source)
```

- [ ] **Step 6: 改 argparse 默认值**

`scripts/survey_drift.py:1134` 现状：
```python
    pa.add_argument("--time_col", default="结束答题时间")
```

改为：
```python
    pa.add_argument("--time_col", default=None,
                    help="时间列名；缺省自动检测（默认列 + 关键词扫描）")
```

- [ ] **Step 7: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/survey_drift.py tests/test_survey_drift_generalize.py && git commit -m "feat(drift): auto-detect time column with keyword scanning + datetime validation"
```

---

## Task 3: build_findings 支持列分桶模式

**Files:**
- Modify: `scripts/survey_drift.py:318-402` (build_findings)
- Test: `tests/test_survey_drift_generalize.py`

- [ ] **Step 1: 写列分桶失败测试**

追加到 `tests/test_survey_drift_generalize.py`：
```python
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
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_survey_drift_generalize.py::test_build_findings_列分桶模式 -v
```

Expected: FAIL with `TypeError: build_findings() got an unexpected keyword argument 'bucket_col'`

- [ ] **Step 3: 改 build_findings 签名与逻辑**

`scripts/survey_drift.py:318-402` 现状签名：
```python
def build_findings(df, classification, granularity, time_col,
                   nps_col, satisfaction_cols, min_n=30):
    if time_col not in df.columns:
        raise KeyError(f"时间列不存在：{time_col}")
    dt = pd.to_datetime(df[time_col], errors="coerce")
    labels, ordered = bucketize(dt, granularity)
    sizes_all = {b: int((labels == b).sum()) for b in ordered}
    low_n_buckets = [b for b, n in sizes_all.items() if n < min_n]
```

改为：
```python
def build_findings(df, classification, granularity, time_col,
                   nps_col, satisfaction_cols, min_n=30,
                   bucket_col=None, bucket_order=None,
                   custom_ranges=None, time_col_source="default"):
    # 模式 B：列分桶
    if bucket_col:
        if bucket_col not in df.columns:
            raise KeyError(f"分桶列不存在：{bucket_col}")
        labels = df[bucket_col].astype(str)
        if bucket_order:
            ordered = [b for b in bucket_order if b in labels.unique()]
            # 补上 order 里没有的桶（避免丢数据）
            extras = [b for b in labels.unique() if b not in ordered]
            ordered = ordered + extras
        else:
            # 按出现顺序去重
            ordered = list(dict.fromkeys(labels.tolist()))
        bucket_mode = "column"
        granularity_out = None
        time_col_out = None
    else:
        # 模式 A：时间分桶
        if time_col not in df.columns:
            raise KeyError(f"时间列不存在：{time_col}")
        dt = pd.to_datetime(df[time_col], errors="coerce")
        if granularity == "custom_ranges":
            labels, ordered = _bucketize_custom(dt, custom_ranges)
        elif granularity == "quarter":
            labels, ordered = _bucketize_quarter(dt)
        else:
            labels, ordered = bucketize(dt, granularity)
        bucket_mode = "time"
        granularity_out = granularity
        time_col_out = time_col

    sizes_all = {b: int((labels == b).sum()) for b in ordered}
    low_n_buckets = [b for b, n in sizes_all.items() if n < min_n]
```

继续把函数末尾 return 改为（line 397-402）：
```python
    return {
        "granularity": granularity_out, "time_col": time_col_out,
        "time_col_source": time_col_source,
        "bucket_mode": bucket_mode,
        "bucket_col": bucket_col,
        "custom_ranges": custom_ranges,
        "buckets": ordered, "bucket_sizes": sizes_all, "low_n_buckets": low_n_buckets,
        "metrics": metrics, "questions": questions,
        "nps_col": nps_col, "satisfaction_cols": metric_cols,
    }
```

- [ ] **Step 4: 新增 quarter 和 custom_ranges 分桶函数**

在 `bucketize` 函数（line 51）之后插入：
```python
def quarter_label(dt):
    q = (dt.month - 1) // 3 + 1
    return f"{str(dt.year)[2:]}年Q{q}"


def _bucketize_quarter(dt_series):
    """按季度分桶。返回 (label_series, ordered_labels)。"""
    labels = dt_series.apply(lambda d: quarter_label(d) if pd.notna(d) else None)
    valid = pd.DataFrame({"label": labels, "dt": dt_series}).dropna(subset=["label"])
    order = valid.groupby("label")["dt"].min().sort_values()
    return labels, list(order.index)


def _bucketize_custom(dt_series, custom_ranges):
    """按自定义区间分桶。custom_ranges = [[label, start, end], ...]。
    区间为左闭右闭（含两端日期）。返回 (label_series, ordered_labels)。"""
    if not custom_ranges:
        raise ValueError("granularity=custom_ranges 时必须传 --custom_ranges")
    parsed = []
    for item in custom_ranges:
        if len(item) != 3:
            raise ValueError(f"custom_ranges 每项需为 [label, start, end]，实际：{item}")
        label, start, end = item
        parsed.append((label, pd.to_datetime(start), pd.to_datetime(end)))

    def _label_for(d):
        if pd.isna(d):
            return None
        for label, start, end in parsed:
            if start <= d <= end:
                return label
        return None

    labels = dt_series.apply(_label_for)
    ordered = [item[0] for item in parsed]
    # 过滤掉空桶
    present = set(labels.dropna().unique())
    ordered = [b for b in ordered if b in present]
    return labels, ordered
```

- [ ] **Step 5: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_survey_drift_generalize.py -v
```

Expected: 7 passed

- [ ] **Step 6: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/survey_drift.py tests/test_survey_drift_generalize.py && git commit -m "feat(drift): support column-bucketing mode and custom/quarter granularity"
```

---

## Task 4: argparse 扩展新参数

**Files:**
- Modify: `scripts/survey_drift.py:1131-1138` (analyze subparser)

- [ ] **Step 1: 扩展 argparse**

`scripts/survey_drift.py:1131-1138` 现状：
```python
    pa = sub.add_parser("analyze", help="分桶 + 检验 → findings JSON")
    pa.add_argument("--file_path", required=True)
    pa.add_argument("--granularity", required=True, choices=["week", "month", "day"])
    pa.add_argument("--time_col", default=None,
                    help="时间列名；缺省自动检测（默认列 + 关键词扫描）")
    pa.add_argument("--nps_col", default=None)
    pa.add_argument("--satisfaction_cols", nargs="*", default=None)
    pa.add_argument("--min_n", type=int, default=30)
    pa.add_argument("--findings_out", default=None)
```

改为：
```python
    pa = sub.add_parser("analyze", help="分桶 + 检验 → findings JSON")
    pa.add_argument("--file_path", required=True)
    pa.add_argument("--granularity", required=False, default=None,
                    choices=["week", "month", "day", "quarter", "custom_ranges"],
                    help="时间分桶粒度；传 --bucket_col 时可省略")
    pa.add_argument("--time_col", default=None,
                    help="时间列名；缺省自动检测（默认列 + 关键词扫描）")
    pa.add_argument("--nps_col", default=None)
    pa.add_argument("--satisfaction_cols", nargs="*", default=None)
    pa.add_argument("--min_n", type=int, default=30)
    pa.add_argument("--findings_out", default=None)
    # 模式 B：列分桶
    pa.add_argument("--bucket_col", default=None,
                    help="非时间维度分桶：指定任意离散列（版本号/活动批次/渠道等）；与 --granularity 互斥")
    pa.add_argument("--bucket_order", default=None,
                    help="列分桶桶顺序，逗号分隔，如 v1.0,v2.0,v3.0；不传则按出现顺序")
    # 自定义区间
    pa.add_argument("--custom_ranges", default=None,
                    help='自定义区间，JSON 数组：[["双11前","2026-10-01","2026-11-10"],...]；--granularity=custom_ranges 时必传')
```

- [ ] **Step 2: 在 _cmd_analyze 里解析 --bucket_order 和 --custom_ranges**

在 `_cmd_analyze` 函数（Task 2 改过的版本）开头 `df = load_df(args.file_path)` 之后插入参数解析：
```python
    # 解析 --bucket_order：逗号分隔 → list
    bucket_order = None
    if args.bucket_order:
        bucket_order = [s.strip() for s in args.bucket_order.split(",") if s.strip()]
    # 解析 --custom_ranges：JSON 字符串 → list
    custom_ranges = None
    if args.custom_ranges:
        try:
            custom_ranges = json.loads(args.custom_ranges)
        except json.JSONDecodeError as e:
            return {"status": "error", "message": f"--custom_ranges JSON 解析失败：{e}"}
    # 互斥校验
    if args.bucket_col and (args.granularity or args.time_col):
        # 列分桶模式忽略 granularity/time_col，不报错但提示
        pass
    if not args.bucket_col and not args.granularity:
        return {"status": "need_input", "reason": "no_bucket_mode",
                "message": "必须指定 --granularity（时间分桶）或 --bucket_col（列分桶）"}
```

把这两个变量传给 build_findings（在 Task 2 已加 `bucket_col=args.bucket_col, bucket_order=args.bucket_order, custom_ranges=args.custom_ranges`，这里要更新变量名一致）。

注意：argparse 把 `--bucket_order` 解析成 `args.bucket_order`，上面解析成局部变量 `bucket_order`，传给 build_findings 时用局部变量。修改 Task 2 里 build_findings 调用：
```python
    findings = build_findings(df, classification, args.granularity, time_col,
                              nps_col, sat_cols, args.min_n,
                              bucket_col=args.bucket_col, bucket_order=bucket_order,
                              custom_ranges=custom_ranges, time_col_source=time_col_source)
```

- [ ] **Step 3: 验证 CLI 帮助文本**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python scripts/survey_drift.py analyze --help
```

Expected: 帮助文本列出所有新参数（`--bucket_col`、`--bucket_order`、`--custom_ranges`、`--granularity` 含 quarter/custom_ranges）。

- [ ] **Step 4: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/survey_drift.py && git commit -m "feat(drift): expose --bucket_col/--bucket_order/--custom_ranges/--granularity quarter in CLI"
```

---

## Task 5: export 适配新 findings 字段

**Files:**
- Modify: `scripts/survey_drift.py:1103-1124` (_cmd_export)
- Modify: `scripts/survey_drift.py:491+` (export_excel 方法与样本 Sheet)

- [ ] **Step 1: 改 default_output_filename 调用**

`scripts/survey_drift.py:1120-1122` 现状：
```python
    out = args.output_path or os.path.join(
        os.path.dirname(os.path.abspath(args.findings)),
        default_output_filename(findings["granularity"]))
```

改为：
```python
    out = args.output_path or os.path.join(
        os.path.dirname(os.path.abspath(args.findings)),
        default_output_filename(findings.get("granularity"), findings.get("bucket_col")))
```

- [ ] **Step 2: 在「方法与样本」Sheet 顶部加分桶模式标注**

找到 export_excel 里生成「ℹ️ 方法与样本」Sheet 的代码块（在 line 491 之后某处）。先读文件定位：
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && grep -n "方法与样本" scripts/survey_drift.py
```

找到写表头的第一行后，在第一行内容之前插入：
```python
    # 分桶模式标注
    bucket_mode = findings.get("bucket_mode", "time")
    if bucket_mode == "column":
        mode_line = f"分桶方式：列分桶，分桶列={findings.get('bucket_col')}"
    else:
        gran = findings.get("granularity", "week")
        mode_line = f"分桶方式：时间分桶，粒度={gran}"
        if gran == "custom_ranges":
            mode_line += "（自定义区间）"
        if findings.get("time_col_source") == "auto_detect":
            mode_line += f"，时间列自动识别={findings.get('time_col')}"
    ws_method.cell(row=1, column=1, value=mode_line)  # 在原有第一行之前
    # 原有第一行及之后内容下移 1 行（调整 row 偏移量）
```

注意：具体行偏移需读现有代码后调整。实施时先 `read_file` 看「方法与样本」Sheet 的写入代码起始行，然后在第一行 cell 写入之前加这一行，后续行号 +1。

- [ ] **Step 3: 验证 export 仍能正常输出**

Run（用现有 findings 文件，如有）:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python scripts/survey_drift.py export --findings <现有 findings.json 路径> --conclusions <现有 conclusions.json 路径>
```

Expected: 输出 `问卷异动诊断_按周_*.xlsx`，且「方法与样本」Sheet 顶部新增分桶方式标注行。

- [ ] **Step 4: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/survey_drift.py && git commit -m "feat(drift): export adapts to bucket_mode and adds mode annotation to method sheet"
```

---

## Task 6: 18-drift-workflow.md 全篇中性化与新能力文档

**Files:**
- Modify: `references/18-drift-workflow.md` (整体)

- [ ] **Step 1: 改标题与触发条件**

`references/18-drift-workflow.md:1-11` 现状：
```markdown
# 时间异动诊断工作流程（阶段 6）

> 📌 **何时读取本文档**：用户有单份含时间列的回流问卷数据，想按周/月/天自动对比、诊断满意度/NPS/单选/多选的异动时，由 SKILL.md 阶段 6 指引跳转至此。

## 触发条件

- "按周/月/天诊断这份回流数据的变化"
- "逐题对比各周/各月的满意度和 NPS 有没有显著变化"
- "回流数据有没有异常波动 / 哪个指标这期掉了/涨了、显著吗"
- MC 等持续回流问卷，想及时发现数据异动
```

改为：
```markdown
# 问卷异动诊断工作流程（阶段 6）

> 📌 **何时读取本文档**：用户有含时间列的问卷数据（回流/满意度/NPS/活动跟踪等），或想按任意离散列对比各桶异动时，由 SKILL.md 阶段 6 指引跳转至此。

## 触发条件

- "按周/月/天/季度诊断这份问卷数据的变化"
- "逐题对比各周/各月的满意度和 NPS 有没有显著变化"
- "问卷数据有没有异常波动 / 哪个指标这期掉了/涨了、显著吗"
- "对比版本 A vs 版本 B 的满意度异动" / "活动前后对比" / "渠道对比"
- MC 等持续回流问卷，想及时发现数据异动
```

- [ ] **Step 2: 改前置段**

`references/18-drift-workflow.md:12-16` 现状：
```markdown
## 前置

- 输入为**含时间列的量化原始 CSV**（列名为编码后 Q1/Q2…，时间列默认 `结束答题时间`）。
- 粒度（周/月/天）由用户指定；未指定时用 `ask_user_question` 让用户三选一。
```

改为：
```markdown
## 前置

- 输入为**含时间列的量化原始 CSV**（列名为编码后 Q1/Q2…，时间列默认 `结束答题时间`）
- 数据来源三选一：
  - **本地文件**：用户直接给路径（路径 A）
  - **下载产物**：`survey_download.py download` 下载的 `files.quantified_data` 文件路径（路径 B）
  - **清洗后产物**：`survey_download.py clean` 清洗后下载的 `files.quantified_data` 文件路径（路径 B+清洗）
- 分桶模式二选一：
  - **时间分桶**（默认）：`--granularity week|month|day|quarter|custom_ranges`，时间列自动检测（默认列 → 关键词扫描 → ask 兜底）
  - **列分桶**：`--bucket_col` 指定任意离散列（版本号/活动批次/渠道等），跳过时间列识别
- 未指定模式时用 `ask_user_question` 让用户选粒度或分桶列
- 下载产物默认带 `结束答题时间` 列，与异动诊断默认时间列对齐，无需额外配置
- `value_labels.json` 放在数据同目录即可被自动加载，适合 MC 月度等场景复用
```

- [ ] **Step 3: 改 Step A 命令示例与可选参数**

`references/18-drift-workflow.md:19-28` 现状：
```markdown
### Step A：运行 analyze，分桶 + 逐题检验

```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --granularity week \
  --findings_out "{数据目录}/drift_findings.json"
```

可选参数：`--time_col`（默认 `结束答题时间`）、`--nps_col`、`--satisfaction_cols`（不传则按关键词自动识别）、`--min_n`（默认 30）。
```

改为：
```markdown
### Step A：运行 analyze，分桶 + 逐题检验

**时间分桶模式（默认）：**
```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --granularity week \
  --findings_out "{数据目录}/drift_findings.json"
```

**列分桶模式（非时间维度对比）：**
```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --bucket_col "Q35.用户版本号" \
  --findings_out "{数据目录}/drift_findings.json"
```

**自定义区间模式（活动/版本节点对比）：**
```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --granularity custom_ranges \
  --custom_ranges '[["双11前","2026-10-01","2026-11-10"],["双11期","2026-11-11","2026-11-13"],["双11后","2026-11-14","2026-11-30"]]' \
  --findings_out "{数据目录}/drift_findings.json"
```

可选参数：
- `--time_col`：时间列名；缺省自动检测（默认列 `结束答题时间` → 关键词扫描「时间/日期/date/time/提交/答题」→ ask 兜底）
- `--nps_col` / `--satisfaction_cols`：不传则按关键词 + 五点量表自动识别
- `--min_n`：默认 30（桶内样本不足不判异动）
- `--bucket_col`：列分桶模式必传；与 `--granularity` 互斥
- `--bucket_order`：列分桶桶顺序，逗号分隔，如 `v1.0,v2.0,v3.0`
- `--custom_ranges`：`--granularity=custom_ranges` 时必传，JSON 数组
```

- [ ] **Step 4: 在 drift_findings.json 结构段补新字段**

`references/18-drift-workflow.md:69-76` 现状：
```markdown
## drift_findings.json 结构

- 顶层：`granularity`、`time_col`、`buckets`（旧→新有序）、`bucket_sizes`、`low_n_buckets`、`metrics`、`questions`、`nps_col`、`satisfaction_cols`。
```

改为：
```markdown
## drift_findings.json 结构

- 顶层：`granularity`（列分桶模式为 null）、`bucket_mode`（"time" / "column"）、`bucket_col`（列分桶模式才有）、`time_col`（时间分桶模式才有）、`time_col_source`（explicit / default / auto_detect / not_applicable）、`custom_ranges`（custom_ranges 粒度才有）、`buckets`（旧→新有序）、`bucket_sizes`、`low_n_buckets`、`metrics`、`questions`、`nps_col`、`satisfaction_cols`。
```

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add references/18-drift-workflow.md && git commit -m "docs(drift): neutralize workflow doc and document column/custom-range bucketing modes"
```

---

## Task 7: 09-survey-download.md / 10-survey-clean.md 加后续操作

**Files:**
- Modify: `references/09-survey-download.md` (末尾追加)
- Modify: `references/10-survey-clean.md` (末尾追加)

- [ ] **Step 1: 09-survey-download.md 文末追加后续操作**

在 `references/09-survey-download.md` 文末（line 118 之后）追加：
```markdown

## 后续操作

下载完成后可衔接的分析流程：

- **做异动诊断**（按周/月/天/季度/自定义区间/列分桶对比各题异动）→ 读取 `18-drift-workflow.md`，
  用下载返回 JSON 中 `files.quantified_data` 的文件路径作为 `survey_drift.py analyze --file_path` 的输入。
  下载的量化数据默认带 `结束答题时间` 列，与异动诊断默认时间列对齐，无需额外配置。
- **做基础统计 / 交叉分析 / 文本分析** → 回到 SKILL.md 主流程，按阶段 1→2→3→4 顺序执行。
```

- [ ] **Step 2: 10-survey-clean.md 文末追加后续操作references/10-survey-clean.md` 文末（line 98 之后）追加：
```markdown

## 后续操作

清洗 + 下载完成后可衔接的分析流程：

- **做异动诊断**（按周/月/天/季度/自定义区间/列分桶对比各题异动）→ 读取 `18-drift-workflow.md`，
  用清洗后下载返回 JSON 中 `files.quantified_data` 的文件路径作为 `survey_drift.py analyze --file_path` 的输入。
  分桶基于清洗后数据，排除无效作答的影响。
- **做基础统计 / 交叉分析 / 文本分析** → 回到 SKILL.md 主流程。
```

- [ ] **Step 3: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add references/09-survey-download.md references/10-survey-clean.md && git commit -m "docs(drift): add 'next steps' section to download/clean docs linking to drift workflow"
```

---

## Task 8: SKILL.md 阶段 6 标题、路由 B 衔接、description 触发语

**Files:**
- Modify: `SKILL.md` (阶段 6 段落、路由 B 段落、description)

- [ ] **Step 1: 找 SKILL.md 阶段 6 和路由 B 段落位置**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && grep -n "时间异动诊断\|阶段 6\|路由 B\|下载成功后" SKILL.md
```

记录行号后实施编辑。

- [ ] **Step 2: 改阶段 6 标题**

找到 SKILL.md 中阶段 6 标题「时间异动诊断」所在行，改为「异动诊断」。

- [ ] **Step 3: 改阶段 6 前置段，加入口说明**

阶段 6 段落开头加入：
```markdown
数据可来自三条路径：
- 本地 CSV/Excel（路径 A）
- `survey_download.py download` 下载的 `files.quantified_data` 文件路径（路径 B）
- `survey_download.py clean` 清洗后下载的 `files.quantified_data` 文件路径（路径 B+清洗）

下载/清洗产物的 `quantified_data` 路径可直接传给 `survey_drift.py analyze --file_path`。
```

- [ ] **Step 4: 路由 B 流程末尾加分支提示**

找到 SKILL.md 路由 B 末尾「下载成功后自动进入阶段 1」一句后追加：
```markdown
如用户明确要求「按周/月对比」「异动诊断」「版本对比」等，可直接跳到阶段 6，用 `quantified_data` 路径作为输入。
```

- [ ] **Step 5: 扩展 description 触发语**

在 SKILL.md frontmatter 的 description 段（Task 3 之前已加过部分触发语）扩展为：
```yaml
当用户说"按周/月/天诊断这份回流数据的变化"、"逐题对比各周/各月的满意度和 NPS 有没有显著变化"、
  "回流数据有没有异常波动"、"哪个指标这期掉了/涨了、显著吗"、"异动分析"、"回流报告"、
  "满意度月度跟踪"、"NPS 月度对比"、"活动前后对比"、"版本 A vs 版本 B 异动"、"渠道对比"等
  涉及问卷时序对比、异动诊断的场景时，也应触发。
```

- [ ] **Step 6: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add SKILL.md && git commit -m "docs(drift): neutralize stage 6 title, add routing B handoff and expand description triggers"
```

---

## Task 9: 回归测试与端到端验证

**Files:**
- Test: `tests/test_survey_drift_generalize.py` (扩展)
- Test: 现有 `tests/test_survey_drift.py`

- [ ] **Step 1: 扩展测试覆盖新粒度**

追加到 `tests/test_survey_drift_generalize.py`：
```python
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
```

- [ ] **Step 2: 运行全部新测试**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_survey_drift_generalize.py -v
```

Expected: 全部 passed

- [ ] **Step 3: 运行原有测试确认向后兼容**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_survey_drift.py -v
```

Expected: 原有测试全部 passed（如有失败需排查是否改动引入回归）

- [ ] **Step 4: 端到端 CLI 验证**

Run 时间分桶模式：
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python scripts/survey_drift.py analyze --file_path <测试 CSV> --granularity week --findings_out /tmp/findings.json
```

Expected: stdout JSON 含 `status=success`、`bucket_mode=time`、`time_col_source=default 或 auto_detect`、文件名 `问卷异动诊断_按周_*.xlsx`。

Run 列分桶模式：
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python scripts/survey_drift.py analyze --file_path <测试 CSV> --bucket_col "Q35.用户版本号" --findings_out /tmp/findings_col.json
```

Expected: stdout JSON 含 `status=success`、`bucket_mode=column`、`bucket_col=Q35.用户版本号`、`granularity=null`。

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add tests/test_survey_drift_generalize.py && git commit -m "test(drift): cover quarter/custom_ranges granularity and time_col_source field"
```

---

## Self-Review 检查

实施完成后执行以下检查：

1. **Spec coverage**: 对照 spec 各节，确认都有对应 Task：
   - 第 4 节分桶双形态 → Task 3
   - 第 5 节时间列自动检测 → Task 2
   - 第 6 节衔接设计 → Task 6/7/8
   - 第 7 节文案中性化 → Task 1/5/6/8
   - 第 8 节 Excel 4 Sheet → Task 5
   - 第 9 节 findings JSON 扩展 → Task 3
   - 第 10 节向后兼容 → Task 9

2. **Placeholder scan**: 检查本计划是否有 TBD/TODO/「类似 Task N」等占位符。如有则补全。

3. **Type consistency**: `detect_time_col` 返回 `(col, source)` 二元组，build_findings 签名新增 `bucket_col/bucket_order/custom_ranges/time_col_source`，_cmd_analyze 调用一致。default_output_filename 新增 `bucket_col` 参数。各 Task 间参数名一致。

4. **行号偏移**: Task 2/3/4 改动 `survey_drift.py` 后行号会变，后续 Task 引用行号时需重新 `grep -n` 定位。
