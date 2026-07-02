---
name: easyexcel2poi
description: "Use when migrating Java SpringBoot/Maven business code from Alibaba EasyExcel to Apache POI without changing business logic; 适用于 EasyExcel 迁移 Apache POI、中文 TDD、基线对比、读写填充 listener handler converter 重构、移除 com.alibaba.excel 依赖。"
---

# EasyExcel 迁移 Apache POI

用这个 skill 迁移 Java SpringBoot/Maven 项目里的 EasyExcel 业务代码。目标是：业务逻辑不变，Excel IO 边界改为 Apache POI，并用 TDD 和语义 workbook 对比证明行为一致。

默认使用弱模型严格路径：先扫描，再建 EasyExcel 基线，再写 POI 红灯测试，再按 bucket 逐个迁移，最后跑统一 gate。不要跳步。

## 不可协商

- 保持业务逻辑不变。只能改 Excel IO 边界、Excel 元数据、测试和 fixture。
- 必须 TDD。替换 EasyExcel 前，先给现有行为补 characterization tests。
- 导出 Excel 必须用 Apache POI 语义快照比较 EasyExcel 基线，不用字节相等。
- 解析 Excel 必须使用真实 fixture。没有 fixture 时，先用旧 EasyExcel 路径反推生成，或停止向用户索要。
- 只有目标模块无 `com.alibaba.excel` 业务代码残留、无 EasyExcel 依赖、Maven 测试通过，才算完成。

## 弱模型严格路径

1. 运行 `scripts/scan_easyexcel_usage.py <project-root>`，保存 EasyExcel 使用清单。
2. 运行 `scripts/generate_migration_worklog.py <project-root>`，生成中文工作日志并照着填。
3. 阅读 `references/strict-checklist.md`，按命令级清单执行；任何失败立即停止。
4. 迁移前阅读 `references/migration-playbook.md`，确认 read/write/fill/web/listener/converter/dependency bucket。
5. 实现 POI 行为时阅读 `references/poi-patterns.md`。
6. 目标项目没有 workbook comparator 时，复制 `assets/java-test-support/SemanticWorkbookSnapshot.java` 到测试代码中。
7. 最终运行 `scripts/migration_gate.py <module-root> --require-baselines`。

## 禁止行为

| 禁止 | 正确做法 |
| --- | --- |
| 先改实现再补测试 | 先写 EasyExcel 基线测试和 POI 红灯测试 |
| 一次性替换所有 read/write/fill | 一次只迁移一个 bucket |
| 为了通过测试削弱 comparator | 修正 POI 输出，或记录用户批准的差异 |
| 用文件字节相等比较 `.xlsx` | 用 POI 解析后的语义快照比较 |
| 测试没绿就删 EasyExcel 依赖 | 行为测试全绿后再移除依赖 |
| 命令失败后继续猜测 | 停止并报告命令、退出码、输出 |

## 参考文件

- `references/strict-checklist.md`：弱模型必须遵守的命令级执行清单。
- `references/migration-playbook.md`：迁移 bucket、TDD 顺序、基线策略和依赖移除。
- `references/poi-patterns.md`：Apache POI 读、写、填充和 EasyExcel 特性映射。

## 最终验收

- EasyExcel characterization tests 能证明旧行为；POI 测试使用同一 fixture 或 baseline。
- `scripts/check_no_easyexcel.py <module-root>` 输出 `PASS`。
- `mvn -B -ntp -f <module-root>/pom.xml dependency:tree -Dincludes=com.alibaba` 无 EasyExcel 依赖。
- `mvn -B -ntp -f <module-root>/pom.xml test` 退出码为 0。
- `scripts/migration_gate.py <module-root> --require-baselines` 输出 `FINAL: PASS`。
