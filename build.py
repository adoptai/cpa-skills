#!/usr/bin/env python3
"""
Package the CPA skills for distribution on the Adopt AI website.

Produces, for each skill in skills/:
  dist/<name>.skill        zip archive - one-click install for Claude / Cowork
  markdown/<name>.md       single self-contained file for copy-paste, with every
                           bundled script inlined as a fenced code block so a
                           copy-paste user does not end up with a broken skill

Run:  python3 build.py
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).parent
SKILLS = ROOT / "skills"
DIST = ROOT / "dist"
MD = ROOT / "markdown"

LANG = {".py": "python", ".md": "markdown", ".csv": "csv", ".json": "json",
        ".sh": "bash", ".txt": "text"}

# The Agent Skills spec defines these optional subdirectories. Anything else in a
# skill folder is not bundled — a distributed archive should never carry stray
# files, and silently shipping them is how a scratch file reaches a client.
SPEC_DIRS = {"references", "scripts", "assets"}

# Build residue that must never reach a distributed archive.
EXCLUDE_DIRS = {"__pycache__", ".pytest_cache", ".ipynb_checkpoints", "node_modules"}
EXCLUDE_SUFFIXES = {".pyc", ".pyo", ".pyd", ".so", ".DS_Store"}


def frontmatter(text: str) -> tuple[str, str]:
    """Return (description, body-after-frontmatter)."""
    if not text.startswith("---"):
        return "", text
    end = text.find("\n---", 3)
    if end == -1:
        return "", text
    fm, body = text[3:end], text[end + 4:]
    desc = ""
    lines = fm.splitlines()
    for i, ln in enumerate(lines):
        if ln.startswith("description:"):
            desc = ln.split(":", 1)[1].strip()
            for cont in lines[i + 1:]:
                if cont.startswith((" ", "\t")):
                    desc += " " + cont.strip()
                else:
                    break
            break
    return desc, body.lstrip("\n")


def build_skill(d: Path) -> dict:
    name = d.name
    skill_md = d / "SKILL.md"
    if not skill_md.exists():
        raise SystemExit(f"{name}: no SKILL.md")
    text = skill_md.read_text(encoding="utf-8")
    desc, _ = frontmatter(text)

    extras, skipped = [], []
    for p in sorted(d.rglob("*")):
        if not p.is_file() or p == skill_md or p.name.startswith("."):
            continue
        rel = p.relative_to(d)
        if set(rel.parts) & EXCLUDE_DIRS or p.suffix in EXCLUDE_SUFFIXES:
            continue                       # build residue, silently ignored
        if rel.parts[0] in SPEC_DIRS:
            extras.append(p)
        else:
            skipped.append(rel.as_posix())

    # ---- .skill archive
    DIST.mkdir(exist_ok=True)
    out = DIST / f"{name}.skill"
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(skill_md, f"{name}/SKILL.md")
        for p in extras:
            z.write(p, f"{name}/{p.relative_to(d)}")

    # ---- flat copy-paste markdown
    MD.mkdir(exist_ok=True)
    parts = [text.rstrip()]
    if extras:
        parts.append(
            "\n\n---\n\n"
            "# Appendix — bundled files\n\n"
            "If you installed the `.skill` package these files are already in place and "
            "you can ignore this appendix. If you copied the skill as text, create the "
            "files below at the paths shown, alongside your `SKILL.md`. The skill will "
            "not run without them.\n"
        )
        for p in extras:
            rel = p.relative_to(d).as_posix()
            lang = LANG.get(p.suffix, "")
            body = p.read_text(encoding="utf-8").rstrip()
            parts.append(f"\n## `{rel}`\n\n```{lang}\n{body}\n```\n")
    (MD / f"{name}.md").write_text("".join(parts) + "\n", encoding="utf-8")

    return {
        "name": name,
        "description": desc,
        "files": len(extras) + 1,
        "extras": [p.relative_to(d).as_posix() for p in extras],
        "skipped": skipped,
        "bytes": out.stat().st_size,
    }


def main() -> None:
    dirs = sorted(x for x in SKILLS.iterdir() if x.is_dir())
    if not dirs:
        raise SystemExit("No skills found in skills/")

    # Regenerate in place rather than wiping the output directories. Some
    # environments (CI containers, synced or mounted folders) disallow unlink,
    # and a build script should not fail for that reason alone.
    expected = {f"{d.name}.skill" for d in dirs} | {f"{d.name}.md" for d in dirs}
    stale = []
    for p in (DIST, MD):
        if p.exists():
            for f in p.iterdir():
                if f.is_file() and not f.name.startswith(".") \
                        and f.name not in expected:
                    try:
                        f.unlink()
                    except OSError:
                        stale.append(f"{p.name}/{f.name}")

    built = [build_skill(d) for d in dirs]

    print(f"Built {len(built)} skill(s)\n")
    skipped_any = False
    for b in built:
        print(f"  {b['name']}")
        print(f"    files {b['files']}   .skill {b['bytes']:,} bytes")
        for e in b["extras"]:
            print(f"      + {e}")
        for s in b["skipped"]:
            skipped_any = True
            print(f"      - NOT BUNDLED: {s}")
    print(f"\n  dist/      {len(built)} .skill packages")
    print(f"  markdown/  {len(built)} self-contained markdown files")
    if skipped_any:
        print(
            f"\nFiles marked NOT BUNDLED sit outside the spec directories "
            f"({', '.join(sorted(SPEC_DIRS))}/).\n"
            f"Either move them into one of those, or delete them — they are not "
            f"reaching users."
        )
    if stale:
        print(
            "\nStale output files could not be removed (no delete permission here). "
            "Delete them manually so they do not ship:"
        )
        for s in stale:
            print(f"  {s}")


if __name__ == "__main__":
    main()
