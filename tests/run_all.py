# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "python-docx>=1.1.2",
#     "pywin32; sys_platform == 'win32'",
#     "pymupdf>=1.24.0",
# ]
# ///
"""Master Test Suite Runner for all Antigravity Plugins.

Discovers and executes skill test suites across plugins, collects statistics,
and generates unified test execution reports.
"""

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

# UTF-8 stdout
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

REPO_ROOT = Path(__file__).resolve().parents[1]
TESTS_ROOT = REPO_ROOT / "tests"
UV_PATH = Path(shutil.which("uv") or r"C:\Users\kekss\.local\bin\uv.exe")


def run_test_script(script_path: Path) -> tuple[int, str]:
    """Execute a skill run_tests.py using uv."""
    cmd = [str(UV_PATH), "run", str(script_path)]
    print(f"\n---> Запуск тестов: {script_path.relative_to(REPO_ROOT)}")
    res = subprocess.run(
        cmd,
        cwd=str(REPO_ROOT),
        capture_output=False,
    )
    return res.returncode, str(script_path.relative_to(REPO_ROOT))


def main():
    parser = argparse.ArgumentParser(description="Master test runner for Antigravity plugins")
    parser.add_argument("--plugin", "-p", help="Run tests only for specific plugin (e.g. docx-tools)")
    args = parser.parse_args()

    print("=" * 75)
    print("ГЛАВНЫЙ РАННЕР ТЕСТОВ ANTIGRAVITY PLUGINS")
    print(f"Корень репозитория: {REPO_ROOT}")
    if args.plugin:
        print(f"Фильтр плагина: {args.plugin}")
    print("=" * 75)

    test_runners = []
    for plugin_dir in sorted(TESTS_ROOT.iterdir()):
        if not plugin_dir.is_dir() or plugin_dir.name.startswith("."):
            continue
        if args.plugin and plugin_dir.name != args.plugin:
            continue

        for skill_dir in sorted(plugin_dir.iterdir()):
            if not skill_dir.is_dir():
                continue
            runner = skill_dir / "run_tests.py"
            if runner.is_file():
                test_runners.append(runner)

    if not test_runners:
        print("Тестовые раннеры не найдены.")
        sys.exit(0)

    results = []
    for runner in test_runners:
        code, rel_path = run_test_script(runner)
        results.append((rel_path, code == 0))

    print("\n" + "=" * 75)
    print("ИТОГ ВЫПОЛНЕНИЯ ВСЕХ НАБОРОВ ТЕСТОВ:")
    print("=" * 75)
    all_passed = True
    for path, passed in results:
        status = "✅ PASS" if passed else "❌ FAIL"
        print(f"[{status}] {path}")
        if not passed:
            all_passed = False

    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
