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
| `--skip-existing` | 输出目录已有同问卷同类型文件时跳过下载，直接复用 | 不跳过 |
| `--no-stat` | 跳过下载系统统计报表（默认同时下载统计报表） | 不跳过 |

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
    "text_data": "C:\\path\\to\\survey_90450_我的世界山头服调研【文本数据】20260101-20260410.xlsx",
    "quantified_data": "C:\\path\\to\\survey_90450_我的世界山头服调研【量化数据】20260101-20260410.csv",
    "stat_data": "C:\\path\\to\\survey_90450_我的世界山头服调研_基础统计_20260101-20260410.xlsx"
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
- **合并时自动过滤水印行**：服务端导出文件末尾会附加「网易内部文件，泄密必究！！！(NetEase Internal documents must be investigated for leakage!!!)」一行水印及前后空行，合并函数会自动剥离，确保数据干净
- **原始文件保留**：解压前会把 ZIP 包复制到 `_raw/` 子目录，分片合并前每份分片也会复制一份到 `_raw/`。合并完成后 `_raw/` 中保留：原始 ZIP + 全部分片，用户可随时回溯原始数据

如果用户反馈文件打不开或太大，建议用 `--type text` 只下载 CSV 文本数据。

## 输出目录结构示例

下载 28k 条数据的问卷（触发分片）后，输出目录的结构：

```
{输出目录}/
├── _raw/                                              ← 原始文件保留区
│   ├── survey_xxx【量化数据】xxx.csv                  ← 原始下载的 ZIP（已重命名为 .csv 但是 ZIP）
│   ├── survey_xxx【量化数据】xxx_1.csv                ← 分片 1 副本
│   └── survey_xxx【量化数据】xxx_2.csv                ← 分片 2 副本
└── survey_xxx【量化数据】xxx.csv                       ← 合并后的最终文件（去水印干净）
```
