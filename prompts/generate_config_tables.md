# Prompt：根据 PPT 和结果表生成配置表草稿

你正在帮助把“人工写代码生成配置表”的流程改成 AI/脚本半自动生成。

你的任务是根据 PPT 模板抽取结果、DP 结果表字段、spec 规则和可选历史配置，生成三张配置表草稿。不要写数据库，不要生成最终 PPT，不要重新计算 DP 指标。

## 输入材料

你会收到：

1. `ppt_slides.csv`：每页 PPT 的标题猜测、文本摘要、图表数量。
2. `ppt_placeholders.csv`：PPT 中的 `${VIEW...}` 占位符。
3. 结果表清单和字段名，例如 `share_of_voice_total_25q3.csv` 的 columns。
4. 可选 spec 规则，说明哪些页面应该用哪些结果表。
5. 可选历史配置表：
   - `bh_database_table`
   - `bh_database_table_field`
   - `bh_charts_replaces`
6. 脚本生成的 `mapping_review.csv`，包含候选映射、置信度和原因。

## 输出目标

生成三张配置表草稿：

```text
bh_database_table.generated.csv
bh_database_table_field.generated.csv
bh_charts_replaces.generated.csv
```

## 规则

- 不允许发明不存在的结果表。
- 不允许发明不存在的字段名。
- 如果文本占位符的值无法从上下文确定，写 `TODO_REVIEW`。
- 优先使用 spec 明确指定的页面-结果表关系。
- 其次使用历史配置中同模块、同 group_sign、同字段的映射。
- 最后才根据 PPT 文本和结果表名做语义匹配。
- 低置信度映射必须标记 `needs_review`。
- 不要直接把草稿当成生产配置。

## `bh_database_table.generated.csv` 字段

```csv
id,page,sort,name,group_sign,pivot,confidence,status,reason
```

字段含义：

- `id`：草稿内自增 ID。
- `page`：PPT 页码。
- `sort`：该页图表序号。
- `name`：结果表名。
- `group_sign`：图表分组标识，通常为去掉 wave 后缀的结果表名。
- `pivot`：默认 `0`，如历史配置或 spec 有特殊规则则沿用。
- `confidence`：`high` / `medium` / `low`。
- `status`：`ready` / `needs_review`。
- `reason`：为什么这样匹配。

## `bh_database_table_field.generated.csv` 字段

```csv
table_id,name,database_column_name,confidence,status,reason
```

字段含义：

- `table_id`：关联 `bh_database_table.generated.csv.id`。
- `name`：PPT 图表字段展示名。
- `database_column_name`：结果表真实字段名。
- `confidence`：字段名来源置信度。
- `status`：是否需要人工审查。
- `reason`：例如 `copied_from_history` 或 `from_result_table_header`。

## `bh_charts_replaces.generated.csv` 字段

```csv
name,value,page,confidence,status,reason
```

字段含义：

- `name`：PPT 占位符名，不带 `${}`。
- `value`：替换值，无法确定时写 `TODO_REVIEW`。
- `page`：占位符所在页。
- `confidence`：值来源置信度。
- `status`：`ready` / `needs_value` / `needs_review`。
- `reason`：为什么这样填。

## 必须检查

输出前确认：

- 每个 `name` 指向真实存在的结果表。
- 每个 `database_column_name` 指向真实存在的结果表字段。
- 每个 `${VIEW...}` 都出现在 `bh_charts_replaces.generated.csv` 中。
- 所有 `TODO_REVIEW` 都保留在输出里，不要自行猜值。
