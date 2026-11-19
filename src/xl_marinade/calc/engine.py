"""The only module that imports ``ironcalc`` (lazily)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from xl_marinade.errors import CalcUnavailable


class Engine:
    """Thin wrapper over an IronCalc model."""

    def __init__(self, model: Any) -> None:
        self._model = model

    @classmethod
    def load(cls, path: Path) -> Engine:
        """Load a workbook into the engine.

        Raises:
            CalcUnavailable: if ``ironcalc`` is not importable.
        """
        try:
            import ironcalc  # noqa: F401
        except ImportError as exc:
            raise CalcUnavailable(
                "the calc engine is not installed; run pip install 'xl-marinade[calc]'"
            ) from exc
        raise NotImplementedError("Engine.load is implemented in a later task")
