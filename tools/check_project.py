"""项目结构自检：只包含不需要人工判断的客观规则。

用法：
    uv run python tools/check_project.py

退出码：0 = 全部通过，1 = 有检查失败。可直接接进提交前流程或 CI。

这里只放三类「客观事实」检查——环就是环、登记了就是登记了、跟踪了就
是跟踪了——因此结论没有解释空间。需要架构判断的规则（例如「核心层不
得 import 具体传输」）不放在这里，除非它们的定义已被明确写进文档。
"""

from __future__ import annotations

import ast
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


def _internal_graph(package: str) -> dict[str, set[str]]:
    """Map each module to the same-package modules it imports."""
    names = set(_modules(package))
    graph: dict[str, set[str]] = {}
    for name in sorted(names):
        source = (ROOT / package / f"{name}.py").read_text(encoding="utf-8")
        deps: set[str] = set()
        for node in ast.walk(ast.parse(source)):
            if not isinstance(node, ast.ImportFrom) or not node.module:
                continue
            if node.level:                                   # from .x import y
                deps.add(node.module.split(".")[0])
            elif node.module.startswith(package + "."):      # from pkg.x import y
                deps.add(node.module.split(".")[1])
        graph[name] = deps & names
    return graph


def _find_cycles(graph: dict[str, set[str]]) -> set[tuple[str, ...]]:
    """Return every import cycle in *graph* as a tuple of module names."""
    cycles: set[tuple[str, ...]] = set()

    def walk(node: str, path: list[str]) -> None:
        for nxt in sorted(graph.get(node, ())):
            if nxt in path:
                cycles.add(tuple(path[path.index(nxt):] + [nxt]))
            else:
                walk(nxt, path + [nxt])

    for name in sorted(graph):
        walk(name, [name])
    return cycles


def check_no_import_cycles() -> list[str]:
    """Internal imports must not form a cycle."""
    problems: list[str] = []
    for package in PACKAGES:
        cycles = _find_cycles(_internal_graph(package))
        if cycles:
            for cycle in sorted(cycles):
                problems.append(f"{package}: " + " -> ".join(cycle))
        else:
            print(f"      ok   {package}: 0 个环")
    return problems


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
    hits: list[str] = []
    for path in tracked:
        name = Path(path).name
        if name in SENSITIVE_EXACT or any(p in path or path.endswith(p) for p in SENSITIVE_PATTERNS):
            hits.append(path)
    if hits:
        return [f"被跟踪的敏感文件: {', '.join(sorted(hits))}"]
    print(f"      ok   {len(tracked)} 个跟踪文件中未发现敏感文件")
    return []


CHECKS = (
    ("内部依赖不成环", check_no_import_cycles),
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
