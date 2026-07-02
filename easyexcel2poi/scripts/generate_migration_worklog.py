#!/usr/bin/env python3
"""根据 EasyExcel 扫描结果生成中文迁移工作日志模板。"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, Iterable, List

from scan_easyexcel_usage import scan


BUCKET_RULES = {
    "read": ("builders", "reader_writer", "listeners"),
    "write": ("builders", "reader_writer", "handlers"),
    "fill": ("fill",),
    "metadata": ("annotations",),
    "extension": ("listeners", "handlers", "converters"),
    "dependency": ("maven_dependency",),
}


def count_bucket(summary: Dict[str, int], names: Iterable[str]) -> int:
    return sum(summary.get(name, 0) for name in names)


def format_hits(hits: List[Dict[str, object]], limit: int = 8) -> List[str]:
    lines: List[str] = []
    for item in hits[:limit]:
        lines.append(f"- `{item['file']}:{item['line']}` {item['text']}")
    if len(hits) > limit:
        lines.append(f"- ... 还有 {len(hits) - limit} 条命中")
    return lines


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("project_root", help="要迁移的项目或模块根目录")
    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    result = scan(project_root)
    summary: Dict[str, int] = result["summary"]  # type: ignore[assignment]
    hits: Dict[str, List[Dict[str, object]]] = result["hits"]  # type: ignore[assignment]

    print("# easyexcel2poi 迁移工作日志")
    print()
    print(f"- 项目根目录: `{project_root}`")
    print("- 规则: 每个命令必须记录退出码和输出摘要；任一失败立即停止。")
    print()
    print("## 扫描摘要")
    print()
    for name, count in summary.items():
        print(f"- {name}: {count}")
    print()
    print("## Bucket 清单")
    print()
    for bucket, names in BUCKET_RULES.items():
        count = count_bucket(summary, names)
        status = "需要处理" if count else "未发现"
        print(f"- {bucket}: {status} ({count})")
    print()
    print("## 关键命中")
    print()
    for name, items in hits.items():
        if not items:
            continue
        print(f"### {name}")
        for line in format_hits(items):
            print(line)
        print()
    print("## 必跑命令记录")
    print()
    print("```text")
    print("Inventory command:")
    print(f"python easyexcel2poi/scripts/scan_easyexcel_usage.py {project_root} > easyexcel-usage.json")
    print("Exit code:")
    print("Summary:")
    print()
    print("Baseline test command:")
    print("mvn -B -ntp -Dtest=<baseline-test> test")
    print("Exit code:")
    print("Result:")
    print()
    print("Baseline artifact command:")
    print(f"python easyexcel2poi/scripts/check_baseline_artifacts.py {project_root}")
    print("Exit code:")
    print("Result:")
    print()
    print("POI RED command:")
    print("mvn -B -ntp -Dtest=<poi-test> test")
    print("Exit code:")
    print("Expected failure:")
    print()
    print("Bucket completed:")
    print("Bucket test command:")
    print("mvn -B -ntp -Dtest=<bucket-test> test")
    print("Exit code:")
    print("Result:")
    print()
    print("Dependency removal command:")
    print(f"python easyexcel2poi/scripts/check_no_easyexcel.py {project_root}")
    print(f"mvn -B -ntp -f {project_root / 'pom.xml'} dependency:tree -Dincludes=com.alibaba")
    print("Exit code:")
    print("Result:")
    print()
    print("Final gate command:")
    print(f"python easyexcel2poi/scripts/migration_gate.py {project_root} --require-baselines")
    print("Exit code:")
    print("Result:")
    print("```")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
