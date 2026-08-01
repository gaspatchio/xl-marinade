"""Lightweight JSONL logger for LLM token usage and call purpose."""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any


def _usage_to_dict(usage: Any) -> dict[str, int]:
    if usage is None:
        return {}
    if isinstance(usage, dict):
        return {
            "prompt_tokens": int(usage.get("prompt_tokens") or 0),
            "completion_tokens": int(usage.get("completion_tokens") or 0),
            "total_tokens": int(usage.get("total_tokens") or 0),
        }
    return {
        "prompt_tokens": int(getattr(usage, "prompt_tokens", 0) or 0),
        "completion_tokens": int(getattr(usage, "completion_tokens", 0) or 0),
        "total_tokens": int(getattr(usage, "total_tokens", 0) or 0),
    }


def log_llm_usage(
    *,
    purpose: str,
    model: str,
    usage: Any,
    provider: str = "openai",
    log_path: Path | None = None,
    session_id: str | None = None,
    turn_index: int | None = None,
    extra: dict[str, Any] | None = None,
) -> None:
    path = log_path or (
        Path(os.environ["LLM_USAGE_LOG_PATH"]) if os.environ.get("LLM_USAGE_LOG_PATH") else None
    )
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)

    event: dict[str, Any] = {
        "ts": datetime.now().isoformat(),
        "purpose": purpose,
        "model": model,
        "provider": provider,
    }
    if session_id or os.environ.get("LLM_USAGE_SESSION_ID"):
        event["session_id"] = session_id or os.environ.get("LLM_USAGE_SESSION_ID")
    if turn_index is not None:
        event["turn_index"] = int(turn_index)
    event.update(_usage_to_dict(usage))
    if extra:
        event.update(extra)

    with open(path, "a") as f:
        f.write(json.dumps(event, sort_keys=True) + "\n")
