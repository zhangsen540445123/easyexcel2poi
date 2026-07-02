#!/usr/bin/env python3
"""检查 EasyExcel baseline 目录是否存在并包含可比较基线文件。"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, List


ALLOWED_SUFFIXES = {".xlsx", ".xls", ".json", ".txt", ".snapshot"}
IGNORED_PARTS = {"target", ".git", ".idea", ".mvn", "build", "out"}


def iter_baselines(root: Path) -> Iterable[Path]:
    baseline_dir = root / "src" / "test" / "resources" / "easyexcel-baseline"
    if not baseline_dir.exists():
        return []
    files: List[Path] = []
    for path in baseline_dir.rglob("*"):
        if not path.is_file():
            continue
        if any(part in IGNORED_PARTS for part in path.parts):
            continue
        if path.suffix.lower() in ALLOWED_SUFFIXES:
            files.append(path)
    return files


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("module_root", help="要检查的 Maven 模块根目录")
    args = parser.parse_args()

    root = Path(args.module_root).resolve()
    baseline_dir = root / "src" / "test" / "resources" / "easyexcel-baseline"
    if not baseline_dir.exists():
        print(f"FAIL: baseline 目录不存在: {baseline_dir}")
        return 1
    if not baseline_dir.is_dir():
        print(f"FAIL: baseline 路径不是目录: {baseline_dir}")
        return 1

    files = list(iter_baselines(root))
    if not files:
        suffixes = ", ".join(sorted(ALLOWED_SUFFIXES))
        print(f"FAIL: baseline 目录中没有可比较文件，允许后缀: {suffixes}")
        return 1

    print(f"PASS: 发现 {len(files)} 个 EasyExcel baseline 文件")
    for path in files:
        print(path.relative_to(root))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
