#!/usr/bin/env python3
"""检查迁移后的 Java/Maven 模块是否仍残留 EasyExcel API 或依赖。"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable, List, Tuple


FORBIDDEN = [
    re.compile(r"com\.alibaba\.excel"),
    re.compile(r"\bEasyExcel\b"),
    re.compile(r"\bExcelWriter\b"),
    re.compile(r"\bExcelReader\b"),
    re.compile(r"<artifactId>\s*(easyexcel|easyexcel-core|easyexcel-support)\s*</artifactId>"),
]

SEARCH_SUFFIXES = {".java", ".xml", ".gradle", ".kts"}
IGNORED_PARTS = {"target", ".git", ".idea", ".mvn", "build", "out"}


def iter_files(root: Path) -> Iterable[Path]:
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if any(part in IGNORED_PARTS for part in path.parts):
            continue
        if path.suffix in SEARCH_SUFFIXES or path.name in {"pom.xml", "build.gradle"}:
            yield path


def scan(root: Path) -> List[Tuple[Path, int, str]]:
    hits: List[Tuple[Path, int, str]] = []
    for path in iter_files(root):
        text = path.read_text(encoding="utf-8", errors="ignore")
        for line_number, line in enumerate(text.splitlines(), 1):
            if any(pattern.search(line) for pattern in FORBIDDEN):
                hits.append((path.relative_to(root), line_number, line.strip()))
    return hits


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("module_root", help="迁移后的模块根目录")
    args = parser.parse_args()

    root = Path(args.module_root).resolve()
    hits = scan(root)
    if hits:
        print("FAIL: 发现 EasyExcel 残留")
        for path, line_number, line in hits:
            print(f"{path}:{line_number}: {line}")
        return 1
    print("PASS: 未发现 EasyExcel API 或依赖残留")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
