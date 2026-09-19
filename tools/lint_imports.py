"""Run import-linter with UTF-8 stdio.

Why this exists
---------------
GitHub Actions 的 Windows runner 把 stdout 重定向。rich 的
``detect_legacy_windows()`` 依据控制台 VT 支持判定（``GetConsoleMode``），在
重定向下为 False，于是 import-linter 走 ``legacy_windows_render``，输出编码落到
locale 的 cp1252。报告含框线符号（``╔`` ``─`` ``▶`` 等），cp1252 无法表示，
在 ``rich/console.py`` 抛 ``UnicodeEncodeError``。

实测（2026-09-19，Windows + 强制 PYTHONIOENCODING=cp1252 复现）：

===========================  ========
方案                          结果
===========================  ========
``PYTHONIOENCODING=utf-8``   通过（唯一有效）
``PYTHONUTF8=1``            失败（优先级低于 PYTHONIOENCODING）
``python -X utf8``          失败（同上）
``TERM=dumb``               失败
``NO_COLOR=1``              失败
``legacy_windows=False``    失败（import-linter 未暴露该选项）
===========================  ========

因此这里显式设置 stdio 编码后再调用 ``lint-imports``，让本地提交钩子与 CI
行为一致——中文 Windows（cp936 可表示框线字符）不会暴露这个问题，英文
Windows 才会。
"""

from __future__ import annotations

import os
import subprocess
import sys


def main() -> int:
    environment = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    try:
        return subprocess.call(["lint-imports", *sys.argv[1:]], env=environment)
    except FileNotFoundError:
        print(
            "lint-imports 不在 PATH 中；请用 `uv run python tools/lint_imports.py` 调用。",
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
