from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any


class ProfileStore:
    def __init__(self, base_dir: Path | None = None) -> None:
        self.base_dir = base_dir or Path("profiles")
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def list_profiles(self) -> list[str]:
        return sorted(path.stem for path in self.base_dir.glob("*.json"))

    def save_profile(self, name: str, data: dict[str, Any]) -> None:
        path = self._path_for_name(name)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_profile(self, name: str) -> dict[str, Any]:
        path = self._path_for_name(name)
        return json.loads(path.read_text(encoding="utf-8"))

    def _path_for_name(self, name: str) -> Path:
        normalized = re.sub(r"[^a-zA-Z0-9_-]+", "_", name.strip())
        if not normalized:
            raise ValueError("Profile name cannot be empty.")
        return self.base_dir / f"{normalized}.json"
