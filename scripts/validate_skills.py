#!/usr/bin/env python3
from __future__ import annotations

import argparse
import py_compile
import re
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path


NAME_RE = re.compile(r"^[a-z0-9][a-z0-9-]{0,63}$")


@dataclass
class Finding:
    path: Path
    message: str


def parse_frontmatter(path: Path) -> tuple[dict[str, str], list[str]]:
    lines = path.read_text(encoding="utf-8").splitlines()
    if not lines or lines[0].strip() != "---":
        return {}, ["SKILL.md must start with YAML frontmatter"]

    try:
        end = lines.index("---", 1)
    except ValueError:
        return {}, ["SKILL.md frontmatter must end with ---"]

    data: dict[str, str] = {}
    errors: list[str] = []
    for index, line in enumerate(lines[1:end], start=2):
        if not line.strip():
            continue
        if ":" not in line:
            errors.append(f"frontmatter line {index} is not key: value")
            continue
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key in data:
            errors.append(f"duplicate frontmatter key: {key}")
        data[key] = value

    return data, errors


def discover_skills(root: Path) -> list[Path]:
    return sorted(
        child
        for child in root.iterdir()
        if child.is_dir()
        and not child.name.startswith(".")
        and (child / "SKILL.md").is_file()
    )


def validate_skill(skill_dir: Path) -> list[Finding]:
    findings: list[Finding] = []
    skill_file = skill_dir / "SKILL.md"
    data, errors = parse_frontmatter(skill_file)

    for error in errors:
        findings.append(Finding(skill_file, error))

    allowed_keys = {"name", "description"}
    for key in sorted(set(data) - allowed_keys):
        findings.append(Finding(skill_file, f"unexpected frontmatter key: {key}"))

    name = data.get("name", "")
    description = data.get("description", "")

    if not name:
        findings.append(Finding(skill_file, "missing frontmatter key: name"))
    elif not NAME_RE.fullmatch(name):
        findings.append(Finding(skill_file, "name must use lowercase letters, digits, and hyphens only"))
    elif name != skill_dir.name:
        findings.append(Finding(skill_file, f"name '{name}' does not match folder '{skill_dir.name}'"))

    if not description:
        findings.append(Finding(skill_file, "missing frontmatter key: description"))
    elif len(description) < 40:
        findings.append(Finding(skill_file, "description is probably too short to trigger reliably"))

    agents_file = skill_dir / "agents" / "openai.yaml"
    if not agents_file.exists():
        findings.append(Finding(skill_dir, "missing recommended agents/openai.yaml"))

    return findings


def compile_scripts(skill_dir: Path, cache_dir: Path) -> list[Finding]:
    findings: list[Finding] = []
    for script in sorted((skill_dir / "scripts").glob("*.py")):
        relative = script.relative_to(skill_dir).with_suffix(".pyc")
        cfile = cache_dir / skill_dir.name / relative
        cfile.parent.mkdir(parents=True, exist_ok=True)
        try:
            py_compile.compile(str(script), cfile=str(cfile), doraise=True)
        except py_compile.PyCompileError as exc:
            findings.append(Finding(script, exc.msg))
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate LynnAI skill folders.")
    parser.add_argument("root", nargs="?", default=".", help="repository root")
    parser.add_argument("--compile-scripts", action="store_true", help="syntax-check Python scripts")
    args = parser.parse_args()

    root = Path(args.root).resolve()
    if not root.is_dir():
        print(f"error: root is not a directory: {root}", file=sys.stderr)
        return 2

    skills = discover_skills(root)
    if not skills:
        print(f"error: no skill folders found under {root}", file=sys.stderr)
        return 1

    findings: list[Finding] = []
    with tempfile.TemporaryDirectory(prefix="lynnai-skill-validate-") as temp_dir:
        cache_dir = Path(temp_dir)
        for skill_dir in skills:
            findings.extend(validate_skill(skill_dir))
            if args.compile_scripts:
                findings.extend(compile_scripts(skill_dir, cache_dir))

    if findings:
        print("Skill validation failed:")
        for finding in findings:
            print(f"- {finding.path.relative_to(root)}: {finding.message}")
        return 1

    print(f"OK: validated {len(skills)} skill(s)")
    for skill_dir in skills:
        print(f"- {skill_dir.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
