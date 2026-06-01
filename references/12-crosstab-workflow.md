# 交叉分析工作流程（阶段 3）

> 📌 **何时读取本文档**：用户要求对比不同人群差异时，由 SKILL.md 阶段 3 指引跳转至此。

## 触发条件

- 用户明确要求对比不同人群（如"分析不同性别/付费玩家的差异"）
- 用户说"请你重点分析一下不同 xxx"
- 用户要求全面分析且数据中存在明显的分组变量

## 执行步骤

### 1. 确定行列变量

- 行变量（被分析的题目）：通常用 `["all"]` 表示全部
- 列变量（分组维度）：用户指定的分组变量列名

### 2. 判断是否需要合并选项

如果分组变量是量表题（如满意度 1-5 分），考虑合并为二分类：
```
--merge_rules '{"Q1.满意度": {"不满意": [1,2,3], "满意": [4,5]}}'
```

### 3. 执行交叉分析

```bash
python {SKILL_DIR}/scripts/crosstab.py \
  --file_path "用户文件路径" \
  --row_questions '["all"]' \
  --col_questions '["Q17.性别"]' \
  --calc_scores auto
```

关键参数：
- `--row_questions`：行变量 JSON 列表。`["all"]` = 所有可分析题目
- `--col_questions`：列变量 JSON 列表。填入确认的分组变量列名
- `--merge_rules`：可选，合并选项规则
- `--calc_scores`：`auto` = 自动检测满意度/NPS 题并计算得分
- `--output_path`：可选，默认 `{文件名}_交叉分析.xlsx`

### 4. 读取输出 JSON

脚本 stdout 返回 JSON，包含：
- `percent_table`：各题各选项在不同分组的百分比
- `diff_summary`：每题的最大差异选项和差异值
- `score_summary`：满意度/NPS 得分（如有）

重点关注 `diff_summary` 中差异值 > 0.05（5pp）的题目，这些是有意义的发现。

> 📖 **参考** `references/05-survey-interpretation.md`：差异判断需结合样本量——小样本（< 100）时 5pp 差异不一定显著；NPS 分差 < 5 分通常为噪音；李克特量表不要只看均值差，需看分布变化。

### 5. 第二轮导出（可选）

分析完 JSON 后，撰写结构化报告 JSON，再次调用脚本导出含报告的 Excel：
```bash
python {SKILL_DIR}/scripts/crosstab.py \
  --file_path "用户文件路径" \
  --row_questions '["all"]' \
  --col_questions '["Q17.性别"]' \
  --calc_scores auto \
  --report_json '{"per_question":[{"question":"Q1...", "finding":"发现..."}], "key_findings":["发现1"], "recommendations":["建议1"], "summary":"总结"}'
```

---

完成后继续执行下一阶段（如有），或跳转到 `references/14-report-workflow.md` 生成报告。
