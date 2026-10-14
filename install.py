#!/usr/bin/env python3
"""Install the ssr-ppt-skill folder into one or more agents' skill directories.

  python install.py --list
  python install.py --agent claude --agent agy
  python install.py --agent cursor --project /path/to/repo    # <repo>/.agents/skills
  python install.py --agent claude --force                    # back up the old copy, then replace

Paths follow each agent's documentation as of 2026-10; check the agent's current docs if one changes.
Install a single copy per agent: several agents read more than one of these directories.
Standard library only.
"""
import argparse
import os
import shutil
import sys
import time
from pathlib import Path

SKILL = "ssr-ppt-skill"
SOURCE = Path(__file__).resolve().parent / SKILL
IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", ".build")


def user_targets(home):
    codex_home = Path(os.environ["CODEX_HOME"]) if os.environ.get("CODEX_HOME") else home / ".codex"
    return {
        "claude": (home / ".claude" / "skills", "Claude Code; also read by Grok Build"),
        "codex": (codex_home / "skills", "Codex"),
        "cursor": (home / ".cursor" / "skills", "Cursor"),
        "grok": (home / ".grok" / "skills", "Grok Build"),
        "agy": (home / ".gemini" / "antigravity-cli" / "skills", "Antigravity CLI (agy)"),
        "agy-app": (home / ".gemini" / "config" / "skills", "Antigravity app and IDE"),
        "agents": (home / ".agents" / "skills", "shared: Cursor, Grok Build, DSH skills manager"),
        "dsh": (home / ".dsh" / "skills", "DeepSeek Harness (per community dsh-skills-manager)"),
    }


def install(dest_dir, force):
    dest = dest_dir / SKILL
    if dest.exists():
        if not force:
            return f"skipped (exists; use --force to back up and replace): {dest}"
        # Keep backups outside the skills directory, or agents would load them as a second skill.
        backup = dest_dir.parent / "skill-backups" / f"{SKILL}-{time.strftime('%Y%m%d-%H%M%S')}"
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(dest), str(backup))
        note = f" (previous copy moved to {backup})"
    else:
        note = ""
    dest_dir.mkdir(parents=True, exist_ok=True)
    shutil.copytree(SOURCE, dest, ignore=IGNORE)
    return f"installed: {dest}{note}"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--agent", action="append", default=[], help="target agent; repeatable (see --list)")
    ap.add_argument("--project", type=Path, help="install into <project>/.agents/skills instead of user directories")
    ap.add_argument("--force", action="store_true", help="back up an existing copy, then replace it")
    ap.add_argument("--list", action="store_true", help="show targets and whether the skill is installed")
    ap.add_argument("--home", type=Path, default=Path.home(), help=argparse.SUPPRESS)  # for tests
    args = ap.parse_args()

    if not (SOURCE / "SKILL.md").is_file():
        sys.exit(f"skill source not found: {SOURCE}")
    targets = user_targets(args.home)

    if args.list or not (args.agent or args.project):
        for name, (path, who) in targets.items():
            state = "installed" if (path / SKILL / "SKILL.md").is_file() else "-"
            print(f"{name:<8} {state:<10} {path}  ({who})")
        if not args.list:
            print("\nChoose targets with --agent NAME (repeatable), or --project DIR.")
        return

    unknown = [a for a in args.agent if a not in targets]
    if unknown:
        sys.exit(f"unknown agent(s): {', '.join(unknown)}; choose from {', '.join(targets)}")
    if args.project:
        print(install(args.project.resolve() / ".agents" / "skills", args.force))
        if args.agent:
            print("note: --project installs once into <project>/.agents/skills; --agent was ignored")
        return
    for name in dict.fromkeys(args.agent):
        print(f"{name}: {install(targets[name][0], args.force)}")


if __name__ == "__main__":
    main()
