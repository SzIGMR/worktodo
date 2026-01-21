"""Plugin registry for skillrunner."""

from __future__ import annotations

from dataclasses import dataclass
from importlib import metadata
from typing import Callable, Dict, Iterable, Protocol


class Skill(Protocol):
    """Protocol for skill plugins."""

    name: str
    description: str

    def run(self, args: list[str]) -> int:
        """Run the skill with CLI arguments."""

    def cli_help(self) -> str:
        """Return a German help string for the skill."""


@dataclass
class SkillInfo:
    name: str
    description: str
    entrypoint: Callable[[], Skill]


class SkillRegistry:
    """Registry that loads skills via entry points."""

    def __init__(self) -> None:
        self._skills: Dict[str, SkillInfo] = {}

    def load(self) -> None:
        self._skills.clear()
        entry_points = metadata.entry_points(group="skillrunner.skills")
        for ep in entry_points:
            try:
                skill_obj = ep.load()()
            except Exception as exc:  # pragma: no cover - defensive
                raise RuntimeError(f"Skill '{ep.name}' konnte nicht geladen werden: {exc}") from exc
            self._skills[ep.name] = SkillInfo(
                name=ep.name,
                description=getattr(skill_obj, "description", ""),
                entrypoint=ep.load,
            )

    def list(self) -> Iterable[SkillInfo]:
        return self._skills.values()

    def get(self, name: str) -> Skill:
        if name not in self._skills:
            raise KeyError(f"Skill '{name}' ist nicht registriert")
        return self._skills[name].entrypoint()()


_registry: SkillRegistry | None = None


def get_registry() -> SkillRegistry:
    global _registry
    if _registry is None:
        _registry = SkillRegistry()
        _registry.load()
    return _registry
