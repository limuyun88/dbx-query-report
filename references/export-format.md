# 通用导出输入格式

`scripts/export_report.py` 读取 UTF-8 JSON，只处理已经执行并核对完整的聚合结果，不连接数据库。顶层字段：

| 字段 | 必需 | 说明 |
| --- | --- | --- |
| `title` | 是 | 报表标题 |
| `source` | 是 | 本次连接/数据来源的用户可见描述 |
| `period` | 是 | 日期范围与所用日期字段 |
| `definition` | 是 | 指标、状态、退款与 NULL 等口径 |
| `sql` | 是 | 已确认并执行的完整带注释 SQL 字符串，多个查询用字符串数组 |
| `executed` | 是 | 必须为 `true`，表明输入来自已执行查询 |
| `complete` | 是 | 必须为 `true`，表明已核对 DBX 返回结果未截断 |
| `datasets` | 是 | 一个或多个结果集 |

每个结果集含 `title`、`columns` 和 `rows`。列定义为 `{"key":"month","label":"月份","type":"text"}`，类型限 `text`、`integer`、`decimal`、`percent`；`rows` 是按 key 组织的对象数组。整数或小数最好以十进制字符串传入，避免浮点数丢精度。NULL 用 JSON `null`，不要改成0。可选 `chart` 为 `{"category":"name","value":"sales","group":"month"}`；`group` 可省略，存在时按月份等字段分别画图。只对非空数值生成柱形图，无图时省略。可选 `note` 说明结果边界。

完整模拟输入见仓库 `tests/fixtures/report.json`。示例只用于测试，与真实连接、字段和业务口径无关。每个结果集按输入顺序展示，排名须由 SQL 在数据库中确定。脚本拒绝缺失列、重复列名、非数值金额、过大输入及不安全的表名。输出目录须尚不存在，避免覆盖。

调用：

```text
python scripts/export_report.py --input tests/fixtures/report.json --out-dir example-output
```
