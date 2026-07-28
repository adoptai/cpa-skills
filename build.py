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

    extras = sorted(
        p for p in d.rglob("*")
        if p.is_file() and p != skill_md and not p.name.startswith(".")
    )

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
        "bytes": out.stat().st_size,
    }


def main() -> None:
    for p in (DIST, MD):
        if p.exists():
            shutil.rmtree(p)
    dirs = sorted(x for x in SKILLS.iterdir() if x.is_dir())
    if not dirs:
        raise SystemExit("No skills found in skills/")
    built = [build_skill(d) for d in dirs]

    print(f"Built {len(built)} skill(s)\n")
    for b in built:
        print(f"  {b['name']}")
        print(f"    files {b['files']}   .skill {b['bytes']:,} bytes")
        for e in b["extras"]:
            print(f"      + {e}")
    print(f"\n  dist/      {len(built)} .skill packages")
    print(f"  markdown/  {len(built)} self-contained markdown files")


if __name__ == "__main__":
    main()
