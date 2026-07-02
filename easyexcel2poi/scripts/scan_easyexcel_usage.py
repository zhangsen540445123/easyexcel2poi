#!/usr/bin/env python3
"""扫描 Java Maven 项目中的 EasyExcel 使用点。"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Dict, Iterable, List


PATTERNS = {
    "imports": re.compile(r"import\s+com\.alibaba\.excel(?:\.|\b)"),
    "builders": re.compile(r"\bEasyExcel\.(read|write|writerSheet|readSheet|writerTable)\b"),
    "reader_writer": re.compile(r"\b(ExcelReader|ExcelWriter|ReadSheet|WriteSheet|WriteTable)\b"),
    "annotations": re.compile(r"@(ExcelProperty|ExcelIgnore|DateTimeFormat|NumberFormat|ColumnWidth|ContentRowHeight|HeadRowHeight|ContentStyle|HeadStyle|ContentLoopMerge|OnceAbsoluteMerge)\b"),
    "listeners": re.compile(r"\b(AnalysisEventListener|ReadListener|PageReadListener)\b"),
    "handlers": re.compile(r"\b(CellWriteHandler|RowWriteHandler|SheetWriteHandler|WorkbookWriteHandler|WriteHandler)\b"),
    "converters": re.compile(r"\b(Converter|ReadConverterContext|WriteConverterContext|StringImageConverter)\b"),
    "fill": re.compile(r"\b(withTemplate|doFill|FillConfig|FillWrapper)\b"),
    "maven_dependency": re.compile(r"<artifactId>(easyexcel|easyexcel-core|easyexcel-support)</artifactId>"),
}


def iter_files(root: Path) -> Iterable[Path]:
    ignored_parts = {"target", ".git", ".idea", ".mvn", "build", "out"}
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if any(part in ignored_parts for part in path.parts):
            continue
        if path.suffix in {".java", ".xml"}:
            yield path


def scan(root: Path) -> Dict[str, object]:
    hits: Dict[str, List[Dict[str, object]]] = {name: [] for name in PATTERNS}
    for path in iter_files(root):
        try:
            lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
        except OSError:
            continue
        easyexcel_file = any("com.alibaba.excel" in line or "EasyExcel" in line for line in lines)
        for line_number, line in enumerate(lines, 1):
            for name, pattern in PATTERNS.items():
                if name not in {"imports", "builders", "maven_dependency"} and not easyexcel_file:
                    continue
                if pattern.search(line):
                    hits[name].append(
                        {
                            "file": str(path.relative_to(root)),
                            "line": line_number,
                            "text": line.strip(),
                        }
                    )

    summary = {name: len(items) for name, items in hits.items()}
    return {"root": str(root), "summary": summary, "hits": hits}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", help="要扫描的项目根目录")
    parser.add_argument("--summary-only", action="store_true", help="只输出各类命中数量")
    args = parser.parse_args()

    root = Path(args.root).resolve()
    result = scan(root)
    if args.summary_only:
        print(json.dumps(result["summary"], indent=2, ensure_ascii=False))
    else:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
