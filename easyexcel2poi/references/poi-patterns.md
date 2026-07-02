# Apache POI 最小迁移模式

本文件只记录 EasyExcel 到 Apache POI 的关键行为映射。业务逻辑以现有测试和 baseline 为准，不按本文件猜业务规则。

## 注解映射

| EasyExcel | POI 迁移做法 |
| --- | --- |
| `@ExcelProperty(value, index)` | 本地元数据注解或字段配置；优先使用显式 `index`，否则按字段声明顺序 |
| `@ExcelIgnore` | 从读写字段集合中排除 |
| `@DateTimeFormat` | 读时按 fixture 格式解析；写时设置 Excel date cell + data format |
| `@NumberFormat` | 写 numeric cell，并设置 data format；不要写成字符串除非基线就是字符串 |
| `@ColumnWidth` | `sheet.setColumnWidth`，单位按 EasyExcel 旧表现校准 |
| `@ContentRowHeight` / `@HeadRowHeight` | `row.setHeightInPoints` |
| `@ContentStyle` / `@HeadStyle` | 创建或复用 `CellStyle`，比较语义样式 |
| `@ContentLoopMerge` / `@OnceAbsoluteMerge` | `sheet.addMergedRegion` |

## Read 对象

- 普通 `.xlsx` 用 `XSSFWorkbook`；扩展名不确定或可能是 `.xls` 时用 `WorkbookFactory.create(InputStream)`。
- 先从 head row 建 header map，再按注解名称或 `index` 映射列。
- 字符串展示值用 `DataFormatter`，不要直接 `cell.toString()`。
- 日期数值 cell 用 `DateUtil` 或 `cell.getLocalDateTimeCellValue()`。
- 空行、空 cell、异常继续/中断行为必须以 EasyExcel 基线测试为准。

## Listener 生命周期

| EasyExcel 行为 | POI 等价行为 |
| --- | --- |
| `invoke(data, context)` | 每解析一行后调用本地 listener 的 `invoke(row)` |
| `doAfterAllAnalysed(context)` | 所有目标行处理完成后只调用一次 |
| `onException(exception, context)` | 捕获单行异常后按旧测试决定继续或抛出 |
| `PageReadListener` | 按旧 page size 累积，满批次回调，最后 flush 尾批次 |

不要把 listener 里的业务逻辑搬进 POI 工具类，只替换读取驱动。

## Write 对象

- 先构建列元数据：字段、标题、index、format、width、style、converter。
- 显式 `index` 优先，未设置 index 的列按字段声明顺序补齐。
- 多行表头用 `String[]` 或列表表达，重复父级标题要合并。
- 日期写 numeric Excel date cell，并给 `CellStyle` 设置 data format。
- 数字写 numeric cell；只有 EasyExcel 基线是字符串时才写字符串。
- formula、comment、hyperlink、rich text 必须用 POI 原生对象写出，并纳入语义比较。

## 图片、合并和样式

- 图片用 `workbook.addPicture`、`CreationHelper.createClientAnchor`、`Drawing.createPicture`。
- 合并区域写入前检查与已有区域是否重叠。
- 样式可缓存复用，避免无限创建 `CellStyle`。
- row height、column width、font、fill、border、alignment 只有业务可见或基线覆盖时才作为必须项。

## Fill 模板

- `{name}` 表示 scalar 替换。
- `{.field}` 或命名 wrapper placeholder 表示 list fill。
- list fill 前先复制或 shift 模板行，再写入数据。
- 模板原有样式要复制到新增行。
- 横向 fill 必须按旧 EasyExcel 基线确定列增长方向和间距。

## Converter / Handler

- converter 只负责 Java 值和 Excel cell 值之间转换，不承载业务流程。
- read/write handler 替换为本地接口时，保持回调时机：workbook、sheet、row、cell 创建前后。
- handler 顺序以旧 EasyExcel 行为测试为准；不确定时写 characterization test。

## 语义比较

用 POI 同时读取 EasyExcel baseline 和 POI candidate，比较 workbook 事实，不比较 zip bytes。最小快照包含：

```text
sheets=1
sheet[0]=SheetName
merged=[A1:C1]
pictures=2
row[0]=STRING:Header|NUMERIC:2020-01-01
```

使用到以下能力时，必须扩展快照：formula、comment、hyperlink、data validation、row height、column width、style、merged region、picture anchor。

不得为了通过测试忽略业务可见差异。只能忽略 OOXML 包装噪声：zip entry 顺序、时间戳、relationship id、creator metadata。

## Maven 依赖

迁移模块应显式依赖：

```xml
<dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi</artifactId>
</dependency>
<dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
</dependency>
```

如果测试运行时 POI/Log4j 有日志实现警告，补测试或运行时日志实现依赖；不要重新引入 EasyExcel。
