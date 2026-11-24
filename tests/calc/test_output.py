"""Tests for the streaming policy writer: a failure never leaves a partial file."""

from __future__ import annotations

import csv
from pathlib import Path

import pytest

from xl_marinade.calc import output
from xl_marinade.calc.output import PolicyWriter


def _paths(tmp_path: Path) -> tuple[Path, Path, Path]:
    return tmp_path / "results.csv", tmp_path / "paths.csv", tmp_path / "rows.csv"


def test_header_failure_closes_and_deletes_every_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    real = csv.writer
    calls = 0

    class Boom:
        def __init__(self, fh: object, **kw: object) -> None:
            nonlocal calls
            calls += 1
            self._inner = real(fh, **kw)  # type: ignore[call-overload]
            self._fail = calls == 2

        def writerow(self, row: list[str]) -> None:
            if self._fail:
                raise OSError("disk full")
            self._inner.writerow(row)

    monkeypatch.setattr(output.csv, "writer", Boom)
    with pytest.raises(OSError, match="disk full"):
        PolicyWriter(*_paths(tmp_path))
    assert list(tmp_path.iterdir()) == []


def test_a_failing_close_still_deletes_the_other_files(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    real = output._Stream.close
    seen: list[Path] = []

    def close(self: output._Stream) -> None:
        seen.append(self.path)
        real(self)
        if len(seen) == 1:
            raise OSError("close failed")

    monkeypatch.setattr(output._Stream, "close", close)
    with pytest.raises(RuntimeError, match="stop"), PolicyWriter(*_paths(tmp_path)):
        raise RuntimeError("stop")
    assert len(seen) == 3
    assert list(tmp_path.iterdir()) == []
