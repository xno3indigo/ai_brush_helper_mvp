# 配置表生成与 PPT-DB 映射审查清单

在允许把 generated 配置表导入数据库或渲染成品 PPT 之前，先按这份清单检查。

## 生成配置表草稿

- 确认 `mapping_review.csv` 中所有 `low` 和 `medium` 置信度行都已人工看过。
- 确认 `bh_database_table.generated.csv` 里的 `page`、`sort`、`name` 正确。
- 确认 `bh_database_table_field.generated.csv` 里的字段名来自真实结果表。
- 确认 `bh_charts_replaces.generated.csv` 中的 `TODO_REVIEW` 已补值或明确保留为待处理。
- 确认只有 `high` 置信度且业务上确认过的行才可以视为 `ready`。

## PPT 占位符

- 确认 `missing_placeholders.csv` 为空，或每一行都能解释。
- 确认每个 `${VIEW...}` 在 `bh_charts_replaces` 中都有唯一值。
- 确认占位符值没有旧 wave、旧日期、旧样本量。
- 确认数据库里只写 `VIEW...` 或写 `${VIEW...}` 时都能被规范化匹配。

## 图表配置

- 确认每个需要刷新的图表都有 `bh_database_table` 配置。
- 确认 `page` 对应 PPT 页码。
- 确认 `sort` 对应该页第几个图表或渲染器约定的图表序号。
- 确认 `name` 指向正确的 DP 结果表。
- 确认 `group_sign`、`pivot` 和渲染器读取逻辑一致。

## 字段配置

- 确认 `missing_columns.csv` 为空。
- 确认 `bh_database_table_field.table_id` 能正确关联到 `bh_database_table.id`。
- 确认 `database_column_name` 在结果表导出中真实存在。
- 确认 `name` 是 PPT 图表里需要展示的系列名、分类名或渲染器约定字段名。

## 结果表导出

- 确认 `missing_tables.csv` 为空。
- 确认导出的结果表是目标 wave，例如 `25q3`。
- 确认表名大小写、后缀、业务模块命名和数据库一致。
- 确认结果表不是空表。

## spec 规则

- 确认 `mapping_spec.generated.yaml` 中每页的 `text_bindings` 和 `chart_bindings` 都符合 spec。
- 对特殊页面标备注：排序、隐藏项、合并项、百分比格式、小数位、base 文案。
- 对无法自动判断的页面，不要直接渲染，先在 spec 中人工补充规则。

## 渲染后校验

- 成品 PPT 中不能残留 `${...}`。
- 每个图表都能打开编辑数据，不应出现空数据源。
- 页码、标题、脚注、wave 标签没有沿用旧项目内容。
- 抽查关键页面与 `ppt_to_db_mapping.csv` 的表名、字段、值一致。
