#!/usr/bin/env python3
"""运行 EasyExcel 迁移 Apache POI 的最终门禁。"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List


def resolve_executable(name: str) -> str:
    """在 Windows 上解析 mvn.cmd 等 shell 友好的可执行文件名。"""
    candidates = [name]
    if os.name == "nt" and not name.lower().endswith((".cmd", ".bat", ".exe", ".com")):
        candidates.extend(f"{name}{suffix}" for suffix in (".cmd", ".bat", ".exe", ".com"))

    for candidate in candidates:
        found = shutil.which(candidate)
        if found:
            return found
        if Path(candidate).exists():
            return str(Path(candidate).resolve())
    return name


def run(label: str, command: List[str], cwd: Path) -> bool:
    print(f"\n== {label} ==")
    print(" ".join(command))
    try:
        completed = subprocess.run(command, cwd=str(cwd), text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    except FileNotFoundError as exc:
        print(f"无法启动命令: {exc}")
        print(f"FAIL: {label} 无法启动")
        return False
    print(completed.stdout)
    if completed.returncode == 0:
        print(f"PASS: {label}")
        return True
    print(f"FAIL: {label} exited {completed.returncode}")
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("module_root", help="迁移后的 Maven 模块根目录")
    parser.add_argument("--mvn", default="mvn", help="Maven 可执行文件")
    parser.add_argument("--skip-tests", action="store_true", help="跳过 Maven test 门禁")
    parser.add_argument("--require-baselines", action="store_true", help="要求存在 EasyExcel baseline 文件")
    args = parser.parse_args()

    module_root = Path(args.module_root).resolve()
    skill_root = Path(__file__).resolve().parents[1]
    mvn = resolve_executable(args.mvn)
    checks = []
    if args.require_baselines:
        checks.append(
            (
                "EasyExcel baseline 文件",
                [sys.executable, str(skill_root / "scripts" / "check_baseline_artifacts.py"), str(module_root)],
            )
        )
    checks.extend(
        [
            (
                "无 EasyExcel 源码残留",
                [sys.executable, str(skill_root / "scripts" / "check_no_easyexcel.py"), str(module_root)],
            ),
            (
                "无 Alibaba 依赖",
                [mvn, "-B", "-ntp", "-f", str(module_root / "pom.xml"), "dependency:tree", "-Dincludes=com.alibaba"],
            ),
        ]
    )
    if not args.skip_tests:
        checks.append(
            (
                "Maven 测试",
                [mvn, "-B", "-ntp", "-f", str(module_root / "pom.xml"), "test"],
            )
        )

    failures = 0
    for label, command in checks:
        if not run(label, command, module_root):
            failures += 1

    if failures:
        print(f"\nFINAL: FAIL ({failures} gate(s) failed)")
        return 1
    print("\nFINAL: PASS (all gates passed)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
