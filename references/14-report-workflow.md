# 报告生成工作流程（阶段 5）

> 📌 **何时读取本文档**：需要生成报告时，由 SKILL.md 阶段 5 指引跳转至此。

## 目标

**默认生成 HTML 满意度报告**（`.html`，含交互式 ECharts 图表），单文件离线可用。
所有产出物统一存放到 SKILL.md「📁 输出目录规则」指定的时间戳文件夹中。

## 执行流程

直接使用 `html_report.py` 一键生成报告，无需先写 Markdown：

→ **读取 `references/15-satisfaction-report.md` 获取完整指标计算逻辑和报告命令。**

```bash
python {SKILL_DIR}/scripts/html_report.py \
  --file_path "量化数据.csv" \
  --survey_name "问卷名称" \
  --survey_id 问卷ID \
  --date_range "起止日期" \
  --clean_desc "清洗逻辑描述" \
  --cross_cols '["Q54.性别列名","Q56.职业列名"]' \
  --theme default \
  --output "输出目录/报告文件名.html"
```

**最终输出确认**：告知用户所有文件已保存在哪个文件夹中，并列出文件清单：
- `survey_XXX【量化数据】xxx.csv` — 原始数据
- `MC_满意度报告_XXX_xxx.html` — HTML 满意度报告

## 后续操作提示

→ **执行 SKILL.md 中的「⭐ 后续操作提示（必须执行）」章节**，根据当前已完成的阶段智能裁剪后展示。