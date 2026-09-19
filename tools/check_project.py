"""项目结构自检：只包含不需要人工判断、且 import-linter 不覆盖的客观规则。

用法：
    uv run python tools/check_project.py

退出码：0 = 全部通过，1 = 有检查失败。

依赖方向与循环依赖**不在这里**——它们由 ``uv run lint-imports`` 负责
（契约见 pyproject.toml 的 ``[tool.importlinter]``）。import-linter 会追出
间接依赖链并给出打破环的建议，比手写 AST 扫描准确，因此不再重复实现。

这里只留两项：
  * 每个模块是否已在 docs/architecture.md 中登记
  * 敏感文件是否被误跟踪（AGENTS.md 的安全边界）
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PACKAGES = ("course_selection", "course_progress")
ARCHITECTURE_DOC = ROOT / "docs" / "architecture.md"

# 每个模块必须出现在架构文档的哪个范围内。
# ``course_progress`` 限定在自己的小节，否则同名的 ``cli.py`` 会被
# ``course_selection/cli.py`` 的登记掩盖。
SECTION_SCOPE: dict[str, tuple[str, ...] | None] = {
    "course_selection": None,        # 模块分散在 §2 / §3.1 / §3.2 / §3.3
    "course_progress": ("3.4",),
}

# 与 AGENTS.md「Safety boundary」一致：这些必须保持未跟踪。
# `.env.example` 是模板，故意允许入仓。
SENSITIVE_PATTERNS = (
    ".private/",
    "storage_state.json",
    ".har",
    ".xls",
    ".xlsx",
    ".docx",
    ".doc",
)
SENSITIVE_EXACT = (".env",)


def _modules(package: str) -> list[str]:
    directory = ROOT / package
    return sorted(
        p.stem for p in directory.glob("*.py")
        if p.stem not in ("__init__", "__main__")
    )


def _sections(doc: str) -> dict[str, str]:
    """Split the architecture doc into '### x.y' sections keyed by their number."""
    sections: dict[str, str] = {}
    for part in re.split(r"\n(?=### )", doc):
        heading = part.split("\n", 1)[0]
        if heading.startswith("### "):
            sections[heading[4:].split()[0]] = part
    return sections


def check_modules_documented() -> list[str]:
    """Every module must be named inside its own package's section of the module map."""
    doc = ARCHITECTURE_DOC.read_text(encoding="utf-8")
    sections = _sections(doc)
    problems: list[str] = []
    for package in PACKAGES:
        scope_keys = SECTION_SCOPE.get(package)
        scope = doc if scope_keys is None else "\n".join(
            sections.get(key, "") for key in scope_keys
        )
        modules = _modules(package)
        missing = [m for m in modules if f"`{m}.py`" not in scope]
        if missing:
            problems.append(f"{package}: 未登记 {', '.join(missing)}")
        else:
            print(f"      ok   {package}: {len(modules)}/{len(modules)} 已登记")
    return problems


def check_no_sensitive_tracked() -> list[str]:
    """Sensitive artefacts must stay untracked (see AGENTS.md)."""
    tracked = subprocess.run(
        ["git", "ls-files"],
        cwd=ROOT, capture_output=True, text=True, check=True,
    ).stdout.splitlines()
    hits = [
        path for path in tracked
        if Path(path).name in SENSITIVE_EXACT
        or any(p in path or path.endswith(p) for p in SENSITIVE_PATTERNS)
    ]
    if hits:
        return [f"被跟踪的敏感文件: {', '.join(sorted(hits))}"]
    print(f"      ok   {len(tracked)} 个跟踪文件中未发现敏感文件")
    return []


CHECKS = (
    ("模块已在架构文档中登记", check_modules_documented),
    ("敏感文件未入仓", check_no_sensitive_tracked),
)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print("项目结构自检\n")
    failed = False
    for index, (title, check) in enumerate(CHECKS, start=1):
        print(f"[{index}/{len(CHECKS)}] {title}")
        try:
            problems = check()
        except Exception as error:                      # 检查本身出错也要可见
            problems = [f"检查执行失败: {type(error).__name__}: {error}"]
        for problem in problems:
            print(f"      FAIL {problem}")
            failed = True
        print()
    print("有检查未通过" if failed else "全部通过")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
