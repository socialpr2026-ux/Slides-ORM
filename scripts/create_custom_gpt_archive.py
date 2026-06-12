"""
create_custom_gpt_archive.py

Собирает чистый архив скриптов для загрузки в Custom GPT.
Копирует только нужные файлы, исключая inputs/, outputs/, __pycache__,
.venv, .vendor_runtime и пользовательские данные.

Usage:
    python scripts/create_custom_gpt_archive.py [--no-zip]

Без --no-zip создаёт ZIP-архив в outputs/.
С --no-zip только обновляет staging-директорию.
"""

import hashlib
import json
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


REPO_ROOT = Path(__file__).resolve().parent.parent
ARCHIVE_NAME_PREFIX = "ai_orm_manager_custom_gpt"
STAGE_DIR = REPO_ROOT / "outputs"
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

# ---------------------------
# Файлы, которые НЕ нужно включать
# ---------------------------
EXCLUDE_PREFIXES = (
    ".venv",
    ".vendor_runtime",
    "__pycache__",
    ".git",
    ".gitignore",
    ".python-version",
    ".vscode",
)
EXCLUDE_NAMES = {
    ".gitignore",
    ".python-version",
    "thumbs.db",
    "desktop.ini",
}
EXCLUDE_SUFFIXES = (".pyc", ".pyo", ".so", ".dll")
EXCLUDE_DIRECTORIES = {
    "inputs",
    "outputs",
    "backups",
    "node_modules",
    ".opencode",
    ".vendor_runtime",
    ".git",
}
EXCLUDE_FILES = {
    "IMPLEMENTATION_LOG_20260602.md",
    "CONTROL_CHECKLIST_20260602.md",
}

# ---------------------------
# Файлы, которые ОБЯЗАТЕЛЬНО нужно включить (если существуют)
# ---------------------------
INCLUDE_GLOBS = [
    # Pipeline
    "ai_presentation_orm_v0_3_13/**/*.py",
    "ai_presentation_orm_v0_3_13/**/*.txt",
    "ai_presentation_orm_v0_3_13/**/*.md",
    "ai_presentation_orm_v0_3_13/run_pipeline.py",
    # Wrapper
    "scripts/report_pptx/generate_report_pptx.mjs",
    # Reference template
    "reference/**/*.pptx",
    # Current Custom GPT docs
    "AGENTS.md",
    "CUSTOM_GPT_README.md",
    "DEEPSEEK_IMPLEMENTATION_INSTRUCTIONS.md",
    "ai_orm_manager_agent_manifest.json",
    "examples/**/*.yaml",
    "examples/**/*.md",
    # Archive script itself
    "scripts/create_custom_gpt_archive.py",
]


def _should_include(rel_path: str, abs_path: Path) -> bool:
    """Проверяет, должен ли файл быть включён в архив."""
    parts = rel_path.replace("\\", "/").split("/")

    # Исключаем целые директории (на любом уровне вложенности)
    for part in parts[:-1]:  # все части, кроме имени файла
        if part in EXCLUDE_DIRECTORIES:
            return False
        if part in EXCLUDE_NAMES | EXCLUDE_FILES:
            return False

    # Исключаем по префиксу rel_path
    for prefix in EXCLUDE_PREFIXES:
        if any(prefix in p for p in parts):
            return False

    # Исключаем по имени файла
    if abs_path.name in EXCLUDE_NAMES:
        return False

    # Исключаем по суффиксу
    if abs_path.suffix in EXCLUDE_SUFFIXES:
        return False

    return True


def collect_files(root: Path) -> list[Path]:
    """Собирает только явно разрешенные файлы, рекурсивно, с фильтрацией."""
    files = []
    seen = set()
    for pattern in INCLUDE_GLOBS:
        for abs_path in sorted(root.glob(pattern)):
            if not abs_path.is_file():
                continue
            resolved = abs_path.resolve()
            if resolved in seen:
                continue
            rel_path = abs_path.relative_to(root).as_posix()
            if _should_include(rel_path, abs_path):
                files.append(abs_path)
                seen.add(resolved)
    return sorted(files)


def audit_archive_root(archive_root: Path) -> dict:
    files = [file.relative_to(archive_root).as_posix() for file in archive_root.rglob("*") if file.is_file()]
    return {
        "contains_outputs": any(path.startswith("outputs/") or "/outputs/" in path for path in files),
        "contains_inputs": any(path.startswith("inputs/") or "/inputs/" in path for path in files),
        "contains_pycache": any("__pycache__" in path for path in files),
        "contains_client_workbooks": any(path.lower().endswith((".xlsx", ".xls", ".csv")) for path in files),
        "contains_legacy_handoff": any(path in {"HANDOFF_CHATGPT_PLUS_CANVAS.md", "AI_ORM_MANAGER_AGENT_KIT.md"} for path in files),
    }


def _legacy_collect_files(root: Path) -> list[Path]:
    """Старая широкая логика сохранена только для диагностики и не используется."""
    files = []
    for abs_path in sorted(root.rglob("*")):
        if not abs_path.is_file():
            continue
        rel_path = abs_path.relative_to(root).as_posix()
        if _should_include(rel_path, abs_path):
            files.append(abs_path)
    return files


def create_staging(included_files: list[Path], root: Path, stage_dir: Path) -> Path:
    """Копирует файлы в staging-директорию."""
    archive_root = stage_dir / ARCHIVE_NAME_PREFIX
    if archive_root.exists():
        shutil.rmtree(archive_root)
    archive_root.mkdir(parents=True, exist_ok=True)

    for src in included_files:
        rel = src.relative_to(root)
        dst = archive_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)

    return archive_root


def ensure_reference_template_alias(archive_root: Path) -> None:
    """Add a stable ASCII template alias for Custom GPT upload checks."""
    reference_dir = archive_root / "reference"
    if not reference_dir.exists():
        return

    template_alias = reference_dir / "template.pptx"
    if template_alias.exists():
        return

    pptx_files = sorted(
        file for file in reference_dir.glob("*.pptx")
        if file.is_file() and file.name.lower() != "template.pptx"
    )
    if pptx_files:
        shutil.copy2(pptx_files[0], template_alias)


def create_zip(archive_root: Path, zip_path: Path) -> int:
    """Создаёт ZIP-архив из staging-директории."""
    count = 0
    with ZipFile(str(zip_path), "w", ZIP_DEFLATED) as zf:
        for file in sorted(archive_root.rglob("*")):
            if file.is_file():
                rel = file.relative_to(archive_root)
                zf.write(str(file), str(rel))
                count += 1
    return count


def write_manifest(zip_path: Path, file_count: int, archive_root: Path) -> dict:
    """Пишет манифест в .manifest.json рядом с ZIP."""
    audit = audit_archive_root(archive_root)
    manifest = {
        "archive": str(zip_path.resolve()),
        "created_at": TIMESTAMP,
        "size_bytes": zip_path.stat().st_size,
        "entry_count": file_count,
        "has_style_rules": (archive_root / "ai_presentation_orm_v0_3_13" / "ai_presentation_orm" / "style_rules.py").exists(),
        "has_slide11_builder": (archive_root / "ai_presentation_orm_v0_3_13" / "ai_presentation_orm" / "slide11_seeding_metrics_builder.py").exists(),
        "has_slide12_builder": (archive_root / "ai_presentation_orm_v0_3_13" / "ai_presentation_orm" / "slide12_final_conclusions_builder.py").exists(),
        "has_manifest": (archive_root / "ai_orm_manager_agent_manifest.json").exists(),
        "has_readme": (archive_root / "CUSTOM_GPT_README.md").exists(),
        "has_wrapper": (archive_root / "scripts" / "report_pptx" / "generate_report_pptx.mjs").exists(),
        "has_reference": (archive_root / "reference").exists() and any((archive_root / "reference").iterdir()),
        "has_reference_template_pptx": (archive_root / "reference" / "template.pptx").exists(),
        "has_regression_reference_decks": (archive_root / "reference" / "regression").exists()
        and any((archive_root / "reference" / "regression").glob("*.pptx")),
        **audit,
    }
    manifest_path = zip_path.with_suffix(".manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=4)
    return manifest


def main():
    do_zip = "--no-zip" not in sys.argv

    archive_base_dir = STAGE_DIR
    stage_dir = archive_base_dir / f"custom_gpt_archive_stage_{TIMESTAMP}"
    try:
        stage_dir.mkdir(parents=True, exist_ok=True)
    except PermissionError:
        archive_base_dir = Path(tempfile.gettempdir()) / "ai_orm_manager_archive"
        stage_dir = archive_base_dir / f"custom_gpt_archive_stage_{TIMESTAMP}"
        stage_dir.mkdir(parents=True, exist_ok=True)

    print(f"[1/4] Collecting files from {REPO_ROOT}...")
    all_files = collect_files(REPO_ROOT)
    print(f"       Found {len(all_files)} files after filtering")

    print(f"[2/4] Copying to staging: {stage_dir / ARCHIVE_NAME_PREFIX}")
    archive_root = create_staging(all_files, REPO_ROOT, stage_dir)
    ensure_reference_template_alias(archive_root)
    print(f"       Copied {len(all_files)} files")

    if do_zip:
        zip_name = f"{ARCHIVE_NAME_PREFIX}_{TIMESTAMP}.zip"
        zip_path = archive_base_dir / zip_name
        print(f"[3/4] Creating ZIP: {zip_path}")
        count = create_zip(archive_root, zip_path)
        print(f"       ZIP created: {count} files, {zip_path.stat().st_size / 1024:.0f} KB")

        print(f"[4/4] Writing manifest...")
        manifest = write_manifest(zip_path, count, archive_root)
        print(f"       Manifest: {zip_path.with_suffix('.manifest.json')}")
        print(f"\nDone! Archive ready: {zip_path}")
    else:
        print(f"\nDone! Staging directory: {stage_dir / ARCHIVE_NAME_PREFIX}")
        print("Run with --zip to create the archive.")


if __name__ == "__main__":
    main()
