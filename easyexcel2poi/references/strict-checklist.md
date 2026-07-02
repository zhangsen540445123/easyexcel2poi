# 弱模型严格执行清单

当模型可能跳过 TDD、猜测 POI 行为、忘记验证时，必须按本文件顺序执行。每一步都要记录命令、退出码和输出摘要。任一命令失败，立即停止。

## 0. 工作日志

先生成工作日志：

```bash
python easyexcel2poi/scripts/generate_migration_worklog.py <project-root> > easyexcel2poi-worklog.md
```

要求：

- `easyexcel2poi-worklog.md` 已生成。
- 日志里列出 read/write/fill/listener/handler/converter/dependency bucket。
- 后续每个命令都把退出码和摘要填回日志。

## 1. 扫描 EasyExcel 使用

```bash
python easyexcel2poi/scripts/scan_easyexcel_usage.py <project-root> > easyexcel-usage.json
python easyexcel2poi/scripts/scan_easyexcel_usage.py <project-root> --summary-only
```

要求：

- `easyexcel-usage.json` 存在。
- summary 不为空时，必须逐项迁移。
- 本步骤禁止修改生产代码。

## 2. 迁移前基线测试

对每个 EasyExcel 解析路径：

- 使用原始 Excel fixture。
- 断言解析后的领域对象、listener 副作用、异常继续/中断行为、sheet/table 选择。
- 没有 fixture 时，先用旧 EasyExcel 导出路径生成；不能生成就停止向用户索要。

对每个 EasyExcel 导出路径：

- 在替换实现前生成 EasyExcel workbook 或语义 snapshot。
- 保存到 `src/test/resources/easyexcel-baseline`。
- 目标项目没有 comparator 时，复制 `assets/java-test-support/SemanticWorkbookSnapshot.java`。

基线命令：

```bash
mvn -B -ntp -Dtest=<baseline-test> test
python easyexcel2poi/scripts/check_baseline_artifacts.py <module-root>
```

要求：

- 旧 EasyExcel 实现下 baseline test 通过。
- baseline 目录存在且包含 `.xlsx`、`.xls`、`.json`、`.txt` 或 `.snapshot` 文件。

## 3. POI 红灯测试

在写 POI 实现前，先写使用同一 fixture 或 baseline 的 POI 测试：

```bash
mvn -B -ntp -Dtest=<poi-test> test
```

要求：

- 测试必须失败。
- 失败原因必须是 POI 行为缺失或与 EasyExcel 基线不一致。
- 如果测试直接通过，说明测试无效，必须重写测试。

## 4. 单 bucket 迁移

一次只迁移一个 bucket，顺序如下：

1. read
2. write
3. fill
4. web stream
5. listeners/handlers/converters
6. dependency removal

每完成一个 bucket，运行：

```bash
mvn -B -ntp -Dtest=<bucket-test> test
```

要求：

- 当前 bucket 测试通过。
- 不相关 bucket 不做顺手重构。
- 行为不一致时修 POI 实现，不能削弱断言或 comparator。

## 5. 移除 EasyExcel

只有行为测试全部通过后，才能移除 EasyExcel import 和 Maven dependency：

```bash
python easyexcel2poi/scripts/check_no_easyexcel.py <module-root>
mvn -B -ntp -f <module-root>/pom.xml dependency:tree -Dincludes=com.alibaba
```

要求：

- `check_no_easyexcel.py` 输出 `PASS`。
- dependency tree 退出码为 0，且没有 EasyExcel 依赖条目。

## 6. 最终 gate

```bash
python easyexcel2poi/scripts/migration_gate.py <module-root> --require-baselines
```

只有输出 `FINAL: PASS` 才能报告完成。

## 禁止行为表

| 模型常见借口 | 必须执行的规则 |
| --- | --- |
| “先改代码再补测试也一样” | 不一样。必须先看见 POI 测试红灯 |
| “只是删除依赖，很简单” | 行为测试全绿后才能删除依赖 |
| “字节不一样但应该没问题” | `.xlsx` 不能比字节，要比语义快照 |
| “这个差异不重要，忽略吧” | 只能忽略 OOXML 包装噪声，业务可见差异必须修 |
| “命令失败但我知道怎么改” | 停止，报告命令、退出码和输出 |

## 工作日志模板

```text
Inventory command:
Exit code:
Summary:

Baseline test command:
Exit code:
Result:

Baseline artifact command:
Exit code:
Result:

POI RED command:
Exit code:
Expected failure:

Bucket completed:
Bucket test command:
Exit code:
Result:

Dependency removal command:
Exit code:
Result:

Final gate command:
Exit code:
Result:
```
