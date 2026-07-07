# PPT-DB 映射助手 MVP 项目交接文档

更新时间：2026-07-07

## 1. 项目定位

本项目的目标是减少“人工写刷数代码 / 人工写配置表”的工作量。

当前关注的核心链路是：

```text
PPT 模板 + 已导入数据库的 DP 结果表 + 可选 spec/Excel 辅助信息
-> 自动识别 PPT 图表与数据库结果表的对应关系
-> 生成 bh_database_table / bh_database_table_field / bh_charts_replaces 配置表草稿
-> 审查低置信映射和缺失项
-> 可选把高置信配置写入数据库配置表
-> 后续渲染器读取配置表，把数据库结果刷新到 PPT
```

目前项目已经完成了一个可工作的 MVP，但还不是全自动、全模板通用的生产系统。最成熟的部分是 `tracking_dlbcl` 数据库 + Tisle PPT 第 7/8 页的局部配置生成和数据库写入。

## 2. 项目目录

项目路径：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp
```

主要文件：

```text
README.md
docs/PROJECT_HANDOFF.md
review_checklist.md

ai_brush_helper/
  run_mvp.py
  run_enhanced_flow.py
  dp_workbook_importer.py
  ppt_chart_inspector.py
  auto_mapper.py
  ppt_renderer.py
  generate_pages_7_8_dlbcl_config.py
  render_pages_7_8_dlbcl.py
  common.py

examples/
  spec_rules_sample.yaml
  spec_rules_tisle_template.yaml
  db_exports_sample/
  result_tables_sample/

reports/
  ...
```

### 2.1 本机路径索引

下面是当前项目交接最常用的本机绝对路径。

项目根目录：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp
```

GitHub 远端：

```text
https://github.com/xno3indigo/ai_brush_helper_mvp.git
```

核心文档：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/README.md
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/docs/PROJECT_HANDOFF.md
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/review_checklist.md
```

核心代码：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/run_mvp.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/run_enhanced_flow.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/dp_workbook_importer.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/ppt_chart_inspector.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/auto_mapper.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/ppt_renderer.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/generate_pages_7_8_dlbcl_config.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/render_pages_7_8_dlbcl.py
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/common.py
```

spec 和示例数据：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/examples/spec_rules_sample.yaml
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/examples/spec_rules_tisle_template.yaml
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/examples/db_exports_sample
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/examples/result_tables_sample
```

当前重点输入文件：

```text
/Users/pingchaolee/Downloads/Tisle_26W1_EC_template with 25W2 data_0705.pptx
/Users/pingchaolee/Downloads/Tisle ESCC 26H1_DP_20260703T1800.xlsx
```

历史参考输入文件：

```text
/Users/pingchaolee/Downloads/dp-az-feature-az-prism-25q3
/Users/pingchaolee/Downloads/prism_总问卷_20251029_25q3(N=659).xlsx
/Users/pingchaolee/Downloads/Breztri_report_az_Prism_template_1014.pptx
/Users/pingchaolee/Downloads/Breztri_report_az_Prism_20251106160455.pptx
/Users/pingchaolee/Desktop/dp-az-main
/Users/pingchaolee/Desktop/自动化文档需要的图
```

当前重点输出目录：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite_dryrun
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_visual
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_test
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/db_exports_20260707
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/tisle_26h1_enhanced_v3
```

当前重点输出文件：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/mapping_review.pages_7_8.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/bh_database_table.generated.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/bh_database_table_field.generated.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/db_write_report.pages_7_8.json
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/summary.pages_7_8.json
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/ppt_inspect/ppt_chart_inventory.csv
```

局部渲染测试产物：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/Tisle_pages_7_8_tracking_dlbcl.current.rendered.pptx
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/render_pages_7_8_data.xlsx
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/render_pages_7_8_report.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/Tisle_pages_7_8_tracking_dlbcl.current.rendered.pdf
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/pdf_preview/page-06.png
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/pdf_preview/page-07.png
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/pdf_preview/page-08.png
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_test/Tisle_pages_7_8_tracking_dlbcl.rendered.pptx
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_test/render_pages_7_8_report.csv
```

数据库导出 Excel：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/db_exports_20260707/tracking_dlbcl_config_and_data_tables_20260707.xlsx
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/db_exports_20260707/tracking_dlbcl_config_and_data_tables_20260707.summary.json
```

增强流程最近产物：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/tisle_26h1_enhanced_v3/summary.enhanced.json
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/tisle_26h1_enhanced_v3/03_mapping/mapping_review.enhanced.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/tisle_26h1_enhanced_v3/03_mapping/mapping_spec.enhanced.json
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/tisle_26h1_enhanced_v3/04_render/Tisle_26W1_EC_template_with_26H1_data.enhanced.v3.rendered.pptx
```

数据库：

```text
MySQL host: 192.168.20.7
MySQL port: 3306
MySQL user: root
MySQL password: root
MySQL database: tracking_dlbcl
```

## 3. 当前能力总览

### 3.1 通用 MVP 流程

入口：

```text
ai_brush_helper/run_mvp.py
```

支持模式：

```text
infer-config  从 PPT、spec、结果表导出中推断配置表草稿
mapping       校验已有配置表、PPT、结果表之间的匹配关系
render        根据配置表或 mapping spec 尝试渲染 PPT
extract       早期模式，用于从问卷 Excel 和历史代码抽上下文
```

这部分适合用来验证一个项目的基本可行性，但对复杂 PPT 内嵌 Excel 图表、特殊图表结构、没有清晰占位符的模板，仍需要补规则或人工审查。

### 3.2 增强流程：Excel 到 PPT 图表映射

入口：

```text
ai_brush_helper/run_enhanced_flow.py
```

它把 DP Excel 解析成结果表，再读取 PPT chart inventory，生成增强 mapping，并可尝试渲染。

典型命令：

```bash
python3 -m ai_brush_helper.run_enhanced_flow \
  --mode all \
  --excel "/path/to/DP.xlsx" \
  --pptx "/path/to/template.pptx" \
  --sheet "DP_问卷" \
  --wave 26h1 \
  --target-wave 26W1 \
  --spec /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/examples/spec_rules_tisle_template.yaml \
  --out /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/enhanced_run \
  --min-confidence medium
```

已知表现：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/tisle_26h1_enhanced_v3/summary.enhanced.json
```

最近一次结果：

```text
bindings: 350
updated_files: 54
ok: 27
no_matching_cells: 307
unsupported_xlsb: 15
skipped_low_confidence: 1
```

结论：增强流程能批量识别和尝试渲染，但距离可靠生产仍有距离。主要问题是很多 PPT 图表内嵌 workbook 使用旧模板结构、辅助列、泛化系列名，导致自动投影不到正确单元格。

### 3.3 局部成熟流程：tracking_dlbcl 第 7/8 页配置生成与写库

入口：

```text
ai_brush_helper/generate_pages_7_8_dlbcl_config.py
```

这个脚本当前是最接近用户真实需求的一段：

```text
PPT 模板 + tracking_dlbcl 数据库结果表
-> 读取第 7/8 页图表结构
-> 读取数据库结果表 schema 和样例数据
-> 自动识别图表位置对应的指标角色
-> 生成配置表 CSV / YAML / review 文件
-> 可选写入数据库 bh_* 配置表
```

它当前针对第 7/8 页做了局部规则：

```text
第 7 页：
  通过图表坐标识别 NPS / SOV / TOM / SOC / Adoption Rate

第 8 页：
  通过图表类型和顺序识别 brand_equity 原始分数 / 比例
```

这不是全页泛化逻辑，但这段是后续通用化的原型。

## 4. 当前数据库状态

数据库连接默认值写在脚本参数中：

```text
host: 192.168.20.7
port: 3306
user: root
password: root
database: tracking_dlbcl
```

相关结果表：

```text
brand_awareness_total_26w2
  brand, tom, unaided, mopb, n

nps_score_total_26w2
  brand, NPS, 推荐者, 贬损者, n

sov_total_26w2
  brand, 覆盖率, sov, 拜访频率, n

brand_equity_total_26w2
  id, attr, ranking,
  赞必佳（芦比替定）_132,
  免疫 + 化疗（不含芦比替定_170,
  福可维（安罗替尼）_161,
  安泰适（塔拉妥单抗）_112
```

配置表结构特点：

```text
bh_database_table.id 不是 auto_increment，写入时必须手动生成 id。
bh_database_table_field.id 是 auto_increment，写入时不需要指定 id。
bh_charts_replaces 当前没有写入内容。
```

最近一次已执行真实写库：

```text
输出目录：
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite

写入前：
bh_database_table: 0
bh_database_table_field: 0
bh_charts_replaces: 0

写入后：
bh_database_table: 5
bh_database_table_field: 21
bh_charts_replaces: 0
```

写入报告：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/db_write_report.pages_7_8.json
```

## 5. 第 7/8 页当前映射结果

已高置信写入数据库：

```text
page 7 chart_sort 1 -> brand_awareness_total_26w2.tom
page 7 chart_sort 3 -> sov_total_26w2.sov
page 7 chart_sort 4 -> nps_score_total_26w2.NPS
page 8 chart_sort 1 -> brand_equity_total_26w2 raw_score
page 8 chart_sort 2 -> brand_equity_total_26w2 score_ratio
```

未写入、需要审查：

```text
page 7 chart_sort 2 -> SOC
原因：数据库中没有找到 soc_total_26w2 或包含 soc 字段的结果表。

page 7 chart_sort 5 -> Adoption Rate
原因：数据库中没有找到 adoption/penetration 对应结果表；该图表内嵌 workbook 是 xlsb，当前读取器不支持解析 xlsb。
```

最新局部渲染测试：

```text
脚本：
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/render_pages_7_8_dlbcl.py

输出目录：
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707

生成 PPT：
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/Tisle_pages_7_8_tracking_dlbcl.current.rendered.pptx

导出数据 Excel：
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/render_pages_7_8_data.xlsx

渲染报告：
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_local_render_20260707/render_pages_7_8_report.csv
```

最新渲染测试结果：

```text
ok:
  slide 7 chart_sort 1 -> brand_awareness_total_26w2.tom
  slide 7 chart_sort 3 -> sov_total_26w2.sov
  slide 7 chart_sort 4 -> nps_score_total_26w2.NPS
  slide 8 chart_sort 1 -> brand_equity_total_26w2 raw score
  slide 8 chart_sort 2 -> brand_equity_total_26w2 score/100

skipped:
  slide 7 chart_sort 2 -> SOC source table/field not found
  slide 7 chart_sort 5 -> Adoption Rate source table not found and embedded workbook is xlsb
```

注意：第 7 页横向小条形图使用静态品牌文字叠加 chart 数据，PowerPoint 渲染时类别轴是反向显示的。因此 `render_pages_7_8_dlbcl.py` 会把 SOV/NPS 的写入矩阵按 bottom-to-top plot order 输出，确保视觉行和品牌文字对齐。

审查文件：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/mapping_review.pages_7_8.csv
```

## 6. 如何运行

### 6.1 Dry run：只生成配置表，不写数据库

```bash
cd /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp

python3 -m ai_brush_helper.generate_pages_7_8_dlbcl_config \
  --pptx "/Users/pingchaolee/Downloads/Tisle_26W1_EC_template with 25W2 data_0705.pptx" \
  --out /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dryrun
```

重点检查：

```text
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dryrun/mapping_review.pages_7_8.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dryrun/bh_database_table.generated.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dryrun/bh_database_table_field.generated.csv
/Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dryrun/summary.pages_7_8.json
```

### 6.2 写入数据库配置表

```bash
python3 -m ai_brush_helper.generate_pages_7_8_dlbcl_config \
  --pptx "/Users/pingchaolee/Downloads/Tisle_26W1_EC_template with 25W2 data_0705.pptx" \
  --out /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite \
  --write-db
```

安全规则：

```text
默认只在 bh_database_table / bh_database_table_field / bh_charts_replaces 为空时写入。
如果配置表已有数据，脚本会跳过写入，并在 db_write_report.pages_7_8.json 中记录 skipped_non_empty_config_tables。
```

如果确认要清空已有配置后重写：

```bash
python3 -m ai_brush_helper.generate_pages_7_8_dlbcl_config \
  --pptx "/Users/pingchaolee/Downloads/Tisle_26W1_EC_template with 25W2 data_0705.pptx" \
  --out /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite_replace \
  --write-db \
  --replace-existing-config
```

注意：`--replace-existing-config` 会删除下面三张表的已有配置：

```text
bh_database_table
bh_database_table_field
bh_charts_replaces
```

### 6.3 查询数据库确认写入

可以用如下 SQL 或 Python 查询：

```sql
select id, page, sort, name, original_name, row_filters
from bh_database_table
order by id;

select table_id, origin_name, name, database_column_name, type, column_index
from bh_database_table_field
order by table_id, column_index;
```

## 7. 输出文件说明

局部脚本会输出：

```text
bh_database_table.generated.csv
  生成的数据库图表配置草稿。

bh_database_table_field.generated.csv
  生成的字段配置草稿。

mapping_review.pages_7_8.csv
  最重要的审查文件。包含每个 PPT chart 的位置、内嵌 workbook、目标数据库表、字段、置信度和原因。

mapping_spec.pages_7_8.json / yaml
  更结构化的映射说明，可供后续渲染器读取。

db_samples.pages_7_8.json
  数据库结果表字段和样例行，用于人工审查。

ppt_inspect/ppt_chart_inventory.csv / json
  PPT chart inventory，包含 slide、chart_sort、chart_x、chart_y、chart type、series、内嵌 workbook 等信息。

db_write_report.pages_7_8.json
  写库状态。只有传 --write-db 时才会是 inserted 或 skipped。

summary.pages_7_8.json
  总结本次绑定数量、字段数量、高置信数量、写库状态等。
```

## 8. 关键实现逻辑

### 8.1 PPT 读取

文件：

```text
ai_brush_helper/ppt_chart_inspector.py
```

功能：

```text
读取 pptx zip 包。
枚举 ppt/slides/slide*.xml。
识别每个 slide 中的 c:chart。
解析 chart XML，获取 chart_type、series_names、category_values。
读取 chart 关联的内嵌 workbook。
读取 chart 的坐标 chart_x/chart_y/chart_cx/chart_cy。
输出 ppt_chart_inventory.csv/json。
```

为什么要读坐标：

```text
第 7 页的 chart workbook 表头很多是“系列 1 / 系列 2 / 系列 3”，不能只靠 workbook 表头判断 NPS/SOV/TOM/SOC。
因此当前用图表在页面上的视觉位置判断指标角色。
```

### 8.2 数据库表画像

文件：

```text
ai_brush_helper/generate_pages_7_8_dlbcl_config.py
```

核心函数：

```text
result_table_names()
table_columns()
table_sample()
build_table_profiles()
score_profile()
best_profile()
```

做法：

```text
排除 bh_* / sys_* 表。
读取每张结果表的字段和样例行。
根据 required_columns、表名关键词、total 优先级打分。
选择得分最高的结果表。
```

### 8.3 第 7 页指标识别

核心函数：

```text
infer_slide7_visual_role()
```

当前规则：

```text
x < 2,000,000 且 y > 3,000,000 -> Adoption Rate
4,000,000 <= x < 8,000,000 且 y < 3,000,000 -> NPS
x >= 8,000,000 且 y < 3,000,000 -> SOV
4,000,000 <= x < 8,000,000 且 y > 3,000,000 -> TOM
x >= 8,000,000 且 y > 3,000,000 -> SOC
```

这是一套针对当前 Tisle 第 7 页布局的局部规则。后续要泛化，需要把坐标规则抽成配置或让模型根据附近文本框自动判断。

### 8.4 配置表写入

核心函数：

```text
write_config_to_database()
config_table_counts()
next_database_table_id()
remap_table_ids()
insert_rows()
```

写入逻辑：

```text
先检查 bh_database_table / bh_database_table_field / bh_charts_replaces 是否为空。
如果非空且未传 --replace-existing-config，则跳过。
如果需要替换，则先 delete 三张配置表。
为 bh_database_table 生成连续 id。
同步 remap bh_database_table_field.table_id。
按数据库真实字段严格 insert。
写入 db_write_report.pages_7_8.json。
```

## 9. 已知问题和风险

### 9.1 还不是全自动通用映射器

`generate_pages_7_8_dlbcl_config.py` 当前专门服务于：

```text
tracking_dlbcl
Tisle PPT 第 7/8 页
26w2 结果表
```

它证明了路线可行，但还没泛化到任意 PPT、任意页、任意数据库。

### 9.2 第 7 页 SOC 缺结果表

当前数据库没有 `soc_total_26w2` 或 `soc` 字段。脚本能识别 PPT 中的 SOC 位置，但无法生成可写配置。

可修复方式：

```text
方案 A：让 DP 结果生成流程产出 soc_total_26w2。
方案 B：如果 SOC 实际来自其他表/字段，补一条指标别名规则。
方案 C：在 spec 中声明 SOC 的来源表和字段。
```

### 9.3 第 7 页 Adoption Rate 缺结果表且 xlsb 不支持

PPT 中 Adoption Rate 对应 chart_sort 5，内嵌 workbook 是 `.xlsb`。

当前读取器：

```text
支持 xlsx / xlsm
不支持 xlsb
```

可修复方式：

```text
方案 A：增加 xlsb 解析依赖，例如 pyxlsb。
方案 B：让模板方把内嵌 workbook 转成 xlsx。
方案 C：不依赖内嵌 workbook，从数据库结果表和 spec 直接重建 chart 数据结构。
```

### 9.4 渲染器仍需加强

`ppt_renderer.py` 和 `render_pages_7_8_dlbcl.py` 已经能做局部渲染验证，但全 PPT 渲染仍有大量边界：

```text
复杂 chart 的缓存数据和内嵌 workbook 同步难。
旧模板可能有辅助列、隐藏计算列。
散点图、组合图、xlsb 图表支持有限。
PPT 上的文本占位符可能被拆分成多个 run。
```

当前项目主线建议先把“配置生成和写库”做准，再继续扩大渲染范围。

### 9.5 数据库写入需要谨慎

`--replace-existing-config` 会删除三张配置表已有数据：

```text
bh_database_table
bh_database_table_field
bh_charts_replaces
```

使用前必须确认没有其他项目共用同一数据库或同一配置表。

## 10. 后续建议

### 10.1 把局部规则配置化

当前第 7 页坐标规则写在 Python 函数中。建议抽成 YAML：

```yaml
pages:
  7:
    charts:
      - region: left_bottom
        metric: adoption_rate
        source_candidates: [adoption_rate_total_26w2, penetration_total_26w2]
      - region: middle_top
        metric: NPS
        source_candidates: [nps_score_total_26w2]
      - region: right_top
        metric: sov
        source_candidates: [sov_total_26w2]
      - region: middle_bottom
        metric: tom
        source_candidates: [brand_awareness_total_26w2]
      - region: right_bottom
        metric: soc
        source_candidates: [soc_total_26w2]
```

这样换模板时不需要改代码。

### 10.2 增加 spec 指标规则入口

建议让脚本支持：

```bash
--spec examples/spec_rules_tisle_template.yaml
```

spec 中声明：

```text
每页有哪些指标。
指标别名。
指标对应候选表。
指标对应字段。
排序规则。
缺失时是否允许 fallback。
```

### 10.3 增加数据库写入前校验

目前只校验字段存在和配置表是否为空。建议补充：

```text
目标 chart 是否都被识别。
source_table 是否都有样例数据。
value 字段是否是数值型。
brand/attr 维度是否能和 PPT workbook 行列匹配。
同一 page/sort 是否重复。
row_filters 是否符合渲染器约定。
```

### 10.4 补齐 SOC / Adoption Rate 结果表

这是当前第 7 页配置不完整的直接原因。

需要确认：

```text
SOC 的业务定义是什么？
SOC 是否等于处方占比 / share of choice / C2？
Adoption Rate 的公式是否使用品牌渗透率、P3M 处方过、P3M 经常处方、P3M 最常处方？
这些指标是否已经在数据库中，只是表名不是 soc/adoption？
```

确认后可补：

```text
soc_total_26w2
adoption_rate_total_26w2
```

或在 spec 中声明对应旧表字段。

### 10.5 把写库能力接入通用 run_mvp

现在写库能力只在局部脚本中。建议后续加入：

```bash
python3 ai_brush_helper/run_mvp.py \
  --mode infer-config \
  ... \
  --write-db
```

并统一支持：

```text
dry run
write-db
replace-existing-config
db-write-report
```

## 11. 推荐交接顺序

给下一个开发者的阅读顺序：

```text
1. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/README.md
2. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/docs/PROJECT_HANDOFF.md
3. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/generate_pages_7_8_dlbcl_config.py
4. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/ppt_chart_inspector.py
5. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/mapping_review.pages_7_8.csv
6. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/reports/dlbcl_pages_7_8_config_dbwrite/db_write_report.pages_7_8.json
7. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/run_mvp.py
8. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/run_enhanced_flow.py
9. /Users/pingchaolee/Documents/Codex/2026-06-30/new-chat/outputs/ai_brush_helper_mvp/ai_brush_helper/ppt_renderer.py
```

建议先复现第 7/8 页 dry run，再考虑泛化。

## 12. 当前 Git 状态提示

截至 2026-07-07，代码、README、示例 spec、局部渲染脚本和交接文档均应提交并推送到 GitHub。

```text
remote:
https://github.com/xno3indigo/ai_brush_helper_mvp.git
```

本次建议提交范围：

```text
M ai_brush_helper/render_pages_7_8_dlbcl.py
M docs/PROJECT_HANDOFF.md
```

`reports/` 下的 PPT、Excel、CSV、PDF、PNG 等运行产物默认不提交；它们已经在文档中记录了本机绝对路径。

交接前或接手后建议确认：

```bash
git status --short
git log -1 --oneline
git remote -v
python3 -m compileall ai_brush_helper
```
