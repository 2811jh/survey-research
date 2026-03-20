# 报告生成工作流程（阶段 5）

> 📌 **何时读取本文档**：需要生成综合报告、转换报告格式时，由 SKILL.md 阶段 5 指引跳转至此。

## 目标

**默认输出 Word 格式报告**（`.docx`），同时保留 Markdown 源文件。
如用户要求其他格式（Excel / TXT / 仅 Markdown），按需调整。

## 执行流程

1. 先按下方模板撰写 Markdown 报告内容，保存为 `{文件名}_分析报告.md`
2. **自动调用** `report_export.py` 将 Markdown 转为 Word：
   ```bash
   python {SKILL_DIR}/scripts/report_export.py --input "{报告.md路径}" --format docx
   ```
   > Word 格式需要 `python-docx`，如环境缺失先执行 `pip install python-docx`
3. 告知用户报告已生成，Word 版为主交付物

## Markdown 报告结构

```markdown
# {问卷名称} 分析报告

## 一、调研概述
- 样本量：XXX 份有效问卷
- 调研时间：（如数据中有）
- 分组信息：按 XX 分组（各组样本量）

## 二、关键发现
1. 发现一（最重要的洞察）
2. 发现二
3. 发现三
（3-5 条，每条 1-2 句话）

## 三、基础统计要点
（从基础统计结果中提取最值得关注的数据点）

## 四、交叉分析核心结论
（如果执行了交叉分析，列出差异最显著的发现）

## 五、文本分析核心结论
（如果执行了文本分析，列出主要主题和关键原声）

## 六、综合洞察
（关联交叉分析和文本分析的结果，形成更深层的理解）
（写法参考：[定量发现]+[定性来源] → 综合解读 → 行动方向；如两者分歧则单独说明原因）

## 七、策略建议
1. 建议一（附数据依据：XX% 用户受影响，受影响约 N~M 人）
2. 建议二
3. 建议三
（按影响力×证据强度×可行性排优先级，用范围值而非精确数字）
```

**报告文件命名**：`{文件名}_分析报告.md`，保存在输入文件同目录。

**最终输出确认**：告知用户生成了哪些文件：
- `{文件名}_分析报告.md` — Markdown 摘要报告
- `{文件名}_基础统计.xlsx` — 基础统计详细数据
- `{文件名}_交叉分析.xlsx` — 交叉分析详细数据（如有）
- `{文件名}_文本分析.xlsx` — 文本分析详细数据（如有）

## 报告格式转换（按需）

如果用户要求输出 Word / Excel / TXT 格式的报告（如"请给我一份 Word 报告"、"导出成 docx"），
在生成 Markdown 报告后，调用格式转换脚本：

```bash
# 转 Word（.docx）
python {SKILL_DIR}/scripts/report_export.py --input "{报告.md路径}" --format docx

# 转 Excel（.xlsx）
python {SKILL_DIR}/scripts/report_export.py --input "{报告.md路径}" --format xlsx

# 转纯文本（.txt）
python {SKILL_DIR}/scripts/report_export.py --input "{报告.md路径}" --format txt

# 也可指定输出路径
python {SKILL_DIR}/scripts/report_export.py --input "{报告.md路径}" --format docx --output "自定义路径.docx"
```

支持格式：`md`（默认）、`txt`、`xlsx`、`docx`（Word）。
Word 格式需要额外依赖 `python-docx`，如环境缺失，先执行 `pip install python-docx`。

## 后续操作提示

→ **执行 SKILL.md 中的「⭐ 后续操作提示（必须执行）」章节**，根据当前已完成的阶段智能裁剪后展示。
