"""CLI for skillrunner."""

from __future__ import annotations

import argparse
import sys

from .plugin import get_registry


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="skillrunner",
        description="Skillrunner CLI zur Ausführung installierter Skills.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("list-skills", help="Verfügbare Skills auflisten")

    run_parser = sub.add_parser("run", help="Einen Skill ausführen")
    run_parser.add_argument("skill", help="Name des Skills")
    run_parser.add_argument("skill_args", nargs=argparse.REMAINDER)

    return parser


def list_skills() -> int:
    registry = get_registry()
    lines = ["Installierte Skills:"]
    for info in registry.list():
        lines.append(f"- {info.name}: {info.description}")
    print("\n".join(lines))
    return 0


def run_skill(skill_name: str, args: list[str]) -> int:
    registry = get_registry()
    try:
        skill = registry.get(skill_name)
    except KeyError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    return skill.run(args)


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    if args.command == "list-skills":
        return list_skills()
    if args.command == "run":
        return run_skill(args.skill, args.skill_args)

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
