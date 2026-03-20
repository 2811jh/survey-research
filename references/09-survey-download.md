# 下载问卷数据

## 命令

```bash
python {SKILL_DIR}/scripts/survey_download.py download --id 问卷ID --output_dir "输出目录"
```

## 参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--id` | 问卷 ID（与 `--name` 二选一） | — |
| `--name` | 问卷名称，模糊匹配（与 `--id` 二选一） | — |
| `--type` | 导出类型：`both` / `text` / `quantified` | `both` |
| `--start` | 起始日期 `YYYY-MM-DD` | 问卷创建时间 |
| `--end` | 结束日期 `YYYY-MM-DD` | 当前时间 |
| `--output_dir` | 输出目录 | 当前工作目录 |
| `--select` | 多个匹配时的选择序号（从 0 开始） | — |
| `--clean` | 下载前自动配置清洗条件（流程详见 `clean.md`，务必先 `--dry-run` 预览确认） | 不清洗 |

## 流程

### 1. 定位问卷

**用户给了 ID** → 直接用 `--id`。

**用户给了名称** → 先搜索：

```bash
python {SKILL_DIR}/scripts/survey_download.py search --name "关键词"
```

根据返回的 JSON：
- 匹配 1 个 → 直接用该 ID
- 匹配多个 → 用 `ask_user_question` 让用户选，选项格式：`[序号] 问卷名称 (ID: xxx, 回收: xxx份, 创建: xxx)`
- 匹配 0 个 → 告知用户，建议换关键词或提供 ID

### 2. 下载

```bash
python {SKILL_DIR}/scripts/survey_download.py download --id 确定的ID --output_dir "目录"
```

默认导出文本+量化两种数据，全部时间范围。用户有特殊要求时用 `--type`、`--start`、`--end` 调整。

### 3. 告知结果

下载成功后告知：问卷名称、ID、文件路径、文件大小。

## 输出格式

成功：
```json
{
  "status": "success",
  "survey_name": "《我的世界》山头服调研",
  "survey_id": 90450,
  "files": {
    "text_data": "C:\\path\\to\\survey_90450【文本数据】xxx.csv",
    "quantified_data": "C:\\path\\to\\survey_90450【量化数据】xxx.xlsx"
  }
}
```

多个匹配：
```json
{
  "status": "multiple_matches",
  "matches": [{"id": 90450, "name": "xxx", "status": "回收中", "responses": 419}, ...]
}
```

## 大文件处理

数据量超 20000 条的问卷，服务端会返回 ZIP 压缩包（内含多个分片文件）。脚本会自动解压。
- CSV 分片：自动合并为单个 CSV
- XLSX 分片：自动合并为单个 CSV（XLSX 合并太慢，转 CSV 更实用）
- 合并需要 `pandas` + `openpyxl`，如果未安装会保留分片文件并提示

如果用户反馈文件打不开或太大，建议用 `--type text` 只下载 CSV 文本数据。
