from __future__ import annotations

import hashlib
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Iterable
from pathlib import Path


def _cache_root() -> Path:
    override = os.environ.get("LYNNAI_SKILLS_CACHE")
    if override:
        return Path(override).expanduser()
    local_app_data = os.environ.get("LOCALAPPDATA")
    if os.name == "nt" and local_app_data:
        return Path(local_app_data) / "LynnAI" / "skills-cache"
    xdg_cache_home = os.environ.get("XDG_CACHE_HOME")
    if xdg_cache_home:
        return Path(xdg_cache_home) / "lynnai-skills"
    return Path.home() / ".cache" / "lynnai-skills"


def _all_available(import_names: Iterable[str]) -> bool:
    return all(importlib.util.find_spec(name) is not None for name in import_names)


def ensure_dependencies(skill_name: str, requirements_file: Path, import_names: Iterable[str]) -> Path | None:
    """Load cached dependencies, installing them from PyPI on the first run.

    Packages are installed under the user's cache directory. Hermes' own Python
    environment and the system interpreter are never modified.
    """

    names = tuple(import_names)
    if _all_available(names):
        return None

    requirements = requirements_file.read_bytes()
    digest = hashlib.sha256(requirements).hexdigest()[:16]
    version_root = _cache_root() / "python" / digest
    target = version_root / "site-packages"
    marker = target / ".ready"

    if marker.exists():
        sys.path.insert(0, str(target))
        importlib.invalidate_caches()
        if _all_available(names):
            return target
        sys.path.remove(str(target))
        shutil.rmtree(version_root, ignore_errors=True)

    version_root.parent.mkdir(parents=True, exist_ok=True)
    temp_dir = Path(tempfile.mkdtemp(prefix="install-", dir=version_root.parent))
    temp_target = temp_dir / "site-packages"
    temp_target.mkdir()

    print(f"[{skill_name}] Installing Python dependencies into {version_root}", file=sys.stderr)
    command = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--disable-pip-version-check",
        "--no-input",
        "--no-warn-conflicts",
        "--target",
        str(temp_target),
        "--requirement",
        str(requirements_file),
    ]
    try:
        subprocess.run(command, check=True, timeout=600)
        (temp_target / ".ready").write_text(digest, encoding="utf-8")
        try:
            temp_dir.rename(version_root)
        except OSError:
            if marker.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)
            else:
                raise
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise

    sys.path.insert(0, str(target))
    importlib.invalidate_caches()
    if not _all_available(names):
        missing = [name for name in names if importlib.util.find_spec(name) is None]
        raise RuntimeError(f"Dependency installation completed but imports are still missing: {', '.join(missing)}")
    return target
