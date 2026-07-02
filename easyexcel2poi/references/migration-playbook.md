# EasyExcel 迁移 Playbook

## Inventory

先运行扫描器：

```bash
python easyexcel2poi/scripts/scan_easyexcel_usage.py <project-root> > easyexcel-usage.json
python easyexcel2poi/scripts/generate_migration_worklog.py <project-root> > easyexcel2poi-worklog.md
```

把命中项归入以下 bucket：

| Bucket | 典型命中 |
| --- | --- |
| read | `EasyExcel.read`、`ExcelReader`、`ReadListener`、`AnalysisEventListener`、`PageReadListener` |
| write | `EasyExcel.write`、`ExcelWriter`、`WriteSheet`、`WriteTable`、`doWrite` |
| fill | `withTemplate`、`doFill`、`FillConfig`、`FillWrapper` |
| metadata | `@ExcelProperty`、`@ExcelIgnore`、日期/数字/样式/宽度/高度/合并注解 |
| extension | converter、read/write handler、exception callback、web stream |
| dependency | Maven `com.alibaba:easyexcel`、`easyexcel-core`、`easyexcel-support` |

## TDD 顺序

1. 给当前 EasyExcel 代码补 characterization tests。
2. read 路径断言领域对象、listener 副作用、异常处理、sheet 选择、无模型 `Map`。
3. write 路径先生成 EasyExcel workbook 或语义 snapshot。
4. 写 POI 测试对比同一 fixture/baseline，并确认它先失败。
5. 每次只实现一个 bucket 的最小 POI 行为。
6. 当前 bucket 测试通过后，再跑模块测试。
7. 行为测试全绿后，才移除 EasyExcel import 和 dependency。

## Baseline 策略

优先把语义快照放在 `src/test/resources/easyexcel-baseline`。可接受格式：

- `.xlsx` 或 `.xls`：原始 EasyExcel workbook。
- `.json` 或 `.txt`：归一化语义快照。
- `.snapshot`：已有 demo 使用的文本语义快照格式。

导出 Excel 至少比较：

- sheet 顺序和名称
- cell 类型和展示值
- formula
- merged regions
- 业务可见的 row height 和 column width
- comment 和 hyperlink
- data validation
- picture 数量和 anchor

忽略 OOXML 包装噪声：zip entry 顺序、时间戳、relationship id、creator metadata。

## 依赖移除

替换完成后运行：

```bash
python easyexcel2poi/scripts/check_no_easyexcel.py <module-root>
mvn -B -ntp -f <module-root>/pom.xml dependency:tree -Dincludes=com.alibaba
mvn -B -ntp -f <module-root>/pom.xml test
```

不要删除项目仍需要的 Apache POI 直接依赖。EasyExcel 曾经传递引入的 POI，如果迁移后仍用于读写 Excel，就必须显式保留。

## 异常处理原则

- fixture 缺失：停止，先生成或向用户索要。
- baseline 对比失败：修 POI 输出，不改弱 comparator。
- Maven dependency tree 仍有 `com.alibaba`：继续定位依赖来源，不能报告完成。
- 编译失败：优先检查 EasyExcel 类型是否残留在 listener、handler、converter 或注解中。
