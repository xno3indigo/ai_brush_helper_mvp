# PPT-DB 映射助手 MVP

这个 MVP 聚焦现在已经明确的需求：

```text
PPT 模板 + spec 指标规则 + 已导入数据库的 DP 结果表
→ 自动推断 PPT 与结果表的对应关系
→ 生成配置表草稿
→ 审查配置表和缺口清单
→ 后续交给渲染流程生成成品 PPT
```

当前脚本**不会写数据库**，也**不会直接修改 PPT**。它先把“PPT 模板里需要什么”和“DP 结果表里有什么”对齐，输出一组可以人工审查、也可以继续喂给渲染代码的配置表草稿。

## 当前支持的模式

- `infer-config`：没有配置表时，从 PPT、spec、结果表导出中自动推断并生成配置表草稿。
- `mapping`：已经有配置表时，检查配置表、PPT 模板、结果表之间是否匹配。
- `render`：读取配置表或 `mapping_spec.generated.yaml`，替换 PPT 文本占位符并刷新支持的图表缓存。
- `extract`：旧模式，用于从问卷 Excel 和人工刷数代码里抽上下文；当前主流程一般不需要。

## 推荐输入

`infer-config` 模式需要这些材料：

- PPT 模板：包含 `${VIEW...}` 文本占位符和图表位置。
- spec 指标规则：可选，用来说明指标口径、页面规则、特殊映射规则。
- DB 导出目录：从数据库导出的 DP 结果表。
- 历史配置表导出：可选，用来复用上一期的 `group_sign`、`pivot`、字段显示名。

如果你还没有配置表，DB 导出目录可以只放结果表：

```text
result_tables/
  adoption_curve_total_25q3.csv
  share_of_voice_total_25q3.csv
  message_recall_dp_total_25q3.csv
```

如果你已经有配置表，`mapping` 模式的 DB 导出目录建议长这样：

```text
db_exports/
  bh_database_table.csv
  bh_database_table_field.csv
  bh_charts_replaces.csv
  adoption_curve_total_25q3.csv
  share_of_voice_total_25q3.csv
  message_recall_dp_total_25q3.csv
```

也支持 `.xlsx` / `.xlsm`。其中：

- `bh_database_table`：图表对应的结果表，重点字段是 `page`、`sort`、`name`、`group_sign`、`pivot`。
- `bh_database_table_field`：图表字段对应的数据库列，重点字段是 `table_id`、`name`、`database_column_name`。
- `bh_charts_replaces`：文本占位符对应值，重点字段是 `name`、`value`。
- 其他结果表：用于校验 `bh_database_table.name` 是否真的有导出、字段是否存在。

## 使用示例

### 生成配置表草稿

这是替代“人工写代码生成配置表”的主要入口：

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode infer-config \
  --pptx /Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx \
  --spec examples/spec_rules_sample.yaml \
  --db-export-dir examples/result_tables_sample \
  --wave 25q3 \
  --out reports/infer_config_sample
```

如果有上一期配置表，可以一起给它：

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode infer-config \
  --pptx /Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx \
  --spec examples/spec_rules_sample.yaml \
  --db-export-dir result_tables_25q3 \
  --history-config-dir db_exports_25q2 \
  --wave 25q3 \
  --out reports/25q3_infer_config
```

会生成：

- `bh_database_table.generated.csv`
- `bh_database_table_field.generated.csv`
- `bh_charts_replaces.generated.csv`
- `mapping_review.csv`
- `config_spec.generated.yaml`
- `validation_report.json`

### 校验已有配置表

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode mapping \
  --pptx /Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx \
  --spec examples/spec_rules_sample.yaml \
  --db-export-dir db_exports \
  --wave 25q3 \
  --out reports/25q3_mapping
```

如果暂时没有 spec，可以不传：

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode mapping \
  --pptx /Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx \
  --db-export-dir db_exports \
  --wave 25q3 \
  --out reports/25q3_mapping
```

仓库里也放了一组最小配置表示例导出，方便先试通 `mapping`：

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode mapping \
  --pptx /Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx \
  --db-export-dir examples/db_exports_sample \
  --wave 25q3 \
  --out reports/sample_mapping
```

### 渲染 PPT

如果已经有配置表和结果表，可以直接渲染：

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode render \
  --pptx /Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx \
  --db-export-dir examples/db_exports_sample \
  --wave 25q3 \
  --out reports/render_sample
```

也可以让渲染器读取 `mapping_spec.generated.yaml`：

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode render \
  --pptx /Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx \
  --mapping-spec reports/sample_mapping/mapping_spec.generated.yaml \
  --db-export-dir examples/db_exports_sample \
  --wave 25q3 \
  --out reports/render_from_spec_sample
```

渲染输出：

- `*.rendered.pptx`：渲染后的 PPT。
- `render_validation.csv`：每页文本替换和每个图表刷新的状态。
- `render_report.json`：渲染摘要，包括残留占位符数量、成功/失败图表数。
- `summary.json`：整体摘要。

## 会生成什么

`infer-config` 模式会生成：

- `bh_database_table.generated.csv`：图表配置表草稿。
- `bh_database_table_field.generated.csv`：图表字段配置表草稿。
- `bh_charts_replaces.generated.csv`：文本替换配置表草稿，无法推断的值会写成 `TODO_REVIEW`。
- `mapping_review.csv`：每条图表映射的置信度、原因、需要审查的点。
- `config_spec.generated.yaml`：配置表草稿对应的 YAML spec。
- `validation_report.json`：低置信度、TODO、待审查字段统计。

字段置信度规则：

- `high`：字段来自历史配置，或在 spec 的 `required_columns` 中明确声明。
- `medium`：字段名属于常见图表角色，例如 `brand`、`value`、`base`、`message`、`stage`。
- `low`：只能确认它来自结果表表头，但没有更多语义证据。

`mapping` 模式会生成：

- `ppt_placeholders.csv`：PPT 模板里的 `${VIEW...}` 占位符。
- `ppt_slides.csv`：每页 PPT 的标题猜测、文本摘要、图表数量。
- `ppt_to_db_mapping.csv`：核心文件，列出 PPT 文本/图表与数据库表、字段、占位符值的映射。
- `missing_placeholders.csv`：PPT 有占位符，但 `bh_charts_replaces` 没有值。
- `missing_tables.csv`：`bh_database_table` 配了结果表，但导出目录里没找到结果表。
- `missing_columns.csv`：字段配置引用了某列，但结果表导出里没有该列。
- `mapping_spec.generated.yaml`：根据 DB 配置生成的可审查映射 spec。
- `summary.json`：本次运行摘要。

`render` 模式会生成：

- 渲染后的 PPTX。
- `render_validation.csv`：文本和图表级别校验结果。
- `render_report.json`：渲染摘要。
- `summary.json`：整体摘要。

## 如何复现到其他项目

每次换一个问卷或 PPT 模板，不需要重新读问卷 Excel，只要准备：

1. 新项目的 PPT 模板。
2. 已经导入数据库的 DP 结果表导出。
3. 可选的 spec 指标规则，写明页面、指标、特殊排序、特殊字段名、图表规则。
4. 可选的上一期配置表导出：`bh_database_table`、`bh_database_table_field`、`bh_charts_replaces`。

然后先跑 `--mode infer-config` 生成配置表草稿，审查 `mapping_review.csv` 和三张 generated CSV。确认后再跑 `--mode mapping` 校验配置表和 PPT/结果表是否匹配，最后跑 `--mode render` 输出成品 PPT。

## 当前边界

这个 MVP 已经开始解决“人工写代码生成配置表”和“按配置渲染 PPT”的问题。当前渲染器边界：

1. 已支持读取配置表或 `mapping_spec.generated.yaml`。
2. 已支持用 `bh_charts_replaces` 替换 PPT 文本占位符。
3. 已支持常见图表的缓存数据刷新，并尽量同步嵌入 Excel。
4. 已支持渲染后校验残留 `${...}` 和图表刷新状态。
5. 复杂图表、特殊透视结构、被拆分到多个文本 run 的占位符，可能仍需要在 `render_validation.csv` 中人工审查。
