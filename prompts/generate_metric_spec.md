# Prompt：生成 PPT-DB 映射 Spec

你正在帮助搭建一个 PPT 自动渲染流程。

你的任务是根据 PPT 模板抽取结果、数据库配置表、结果表字段和人工 spec 规则，生成一份可审查的 YAML 映射 spec。不要重新推导问卷公式，不要生成上游 DP 计算代码。

## 输入材料

你会收到：

1. `ppt_placeholders.csv`：PPT 占位符、页码、标题猜测、附近文本。
2. `ppt_slides.csv`：PPT 每页文本摘要、图表数量。
3. `ppt_to_db_mapping.csv`：根据 `bh_database_table`、`bh_database_table_field`、`bh_charts_replaces` 生成的初始映射。
4. `missing_placeholders.csv`、`missing_tables.csv`、`missing_columns.csv`。
5. 人工 spec 指标规则，可选。

## 规则

- 不允许发明不存在的数据库表、字段或占位符。
- 如果一个 PPT 位置和数据库结果表的关系不确定，写 `TODO`，并在 `questions` 中说明。
- 优先相信数据库配置表，其次参考 PPT 文本，最后才使用命名相似度推断。
- 必须保留 `page`、`sort`、`group_sign`、`pivot`。
- 区分文本替换和图表数据绑定：
  - 文本替换来自 `bh_charts_replaces`。
  - 图表数据来自 `bh_database_table` + `bh_database_table_field` + DP 结果表。
- 如果缺表、缺字段、缺占位符值，必须写进 `validation_errors`。

## 输出格式

只返回 YAML：

```yaml
wave: 25q3
pages:
  - page: 7
    status: ready
    text_bindings:
      - placeholder: "${VIEW_7_1}"
        source: bh_charts_replaces
        value: "N=659"
    chart_bindings:
      - sort: 1
        table: "share_of_voice_total_25q3"
        group_sign: "share_of_voice_total"
        pivot: "0"
        fields:
          - db_column: "brand"
            ppt_name: "品牌"
          - db_column: "value"
            ppt_name: "数值"
    validation_errors: []
    questions: []
```

## 必须检查

最终输出 YAML 之前，确认：

- 每个 `${VIEW...}` 要么已映射，要么列入 `validation_errors`。
- 每个 `bh_database_table.name` 都能找到对应结果表，除非结果表导出未提供。
- 每个 `database_column_name` 都能在结果表导出中找到。
- 每页的图表数量和 `chart_bindings` 数量是否明显不一致。
