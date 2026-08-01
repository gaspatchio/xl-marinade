import json
import os
from pathlib import Path

DEFAULT_MODEL_MATRIX: dict[str, str] = {
    "default": "gpt-5.4",
    # Reasoning engine pipeline
    "question_classifier": "gpt-4.1",
    "planner": "gpt-4.1",
    "planner_repair": "gpt-4.1",
    "writer": "gpt-4.1",
    # `auditor_holistic` was removed in the Cycle 7 whole-auditor strip
    # (2026-05-06) — the LLM EvidenceAuditor it referred to is gone.
    "scope_summarizer": "gpt-4.1",
    # Semantic index tagger (classification)
    "semantic_tagger_v1_concept_assignment": "gpt-5-nano",
    # Ontology hospital enrichment (P2-T0.5-PR5 / cost-audit §6 PR-5):
    # classification + schema-fill task on short prompts. ~$0.002/proposal
    # on nano vs ~$0.10/proposal on gpt-4.1. Batch saving ~$8/hospital run.
    "ontology_enrichment": "gpt-5-nano",
    # Enrichment & orchestration
    "intent_classification": "gpt-5.4",
    "orchestrator_chat": "gpt-5.4",
    "label_enrich": "gpt-5.4",
    "structural_fix": "gpt-4.1",
    "segmentation_tag": "gpt-4.1",
    "table_group_merge": "gpt-4.1",
    "entity_model_build": "gpt-5.4",
    "semantic_index_adjudicate": "gpt-5.4",
    "llm_judge": "gpt-5.4",
    # Phase-1 §2.8a runtime quality flag (feature-flagged) — single
    # binary "may-be-incomplete" signal per answer. Uses gpt-5-nano for
    # cost reasons (~$0.002/answer at current pricing). See
    # docs/future_phases/week3-5/phase1-implementation-spec-2026-04-22.md
    # §2.8a. This is NOT the §2.8 calibrated scorer (which is undeployable
    # per Gate-1 FAIL); it is the optional feature-flagged runtime judge.
    "llm_judge_lite": "gpt-5-nano",
}

MODEL_PROFILE_ALIASES: dict[str, str] = {
    "prod": "production",
    "production": "production",
    "r_and_d": "r_and_d",
    "rnd": "r_and_d",
    "r&d": "r_and_d",
}

PROFILE_DEFAULT_MODEL: dict[str, str] = {
    "production": "gpt-5.4",
    "r_and_d": "gpt-5-mini",
}


def _load_matrix_from_path(path: Path) -> dict[str, str] | None:
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if isinstance(payload, dict):
        return {str(k): str(v) for k, v in payload.items() if v is not None}
    return None


def normalize_model_profile(profile: str | None) -> str | None:
    if profile is None:
        return None
    key = str(profile).strip().lower()
    if not key:
        return None
    return MODEL_PROFILE_ALIASES.get(key)


def _apply_profile_overlay(matrix: dict[str, str], profile: str | None) -> dict[str, str]:
    normalized = normalize_model_profile(profile)
    if normalized != "r_and_d":
        return matrix
    cheap_model = PROFILE_DEFAULT_MODEL["r_and_d"]
    # Keep the same purpose keys but force all to the R&D model for cost control.
    return dict.fromkeys(matrix.keys(), cheap_model)


def load_model_matrix(*, profile: str | None = None) -> dict[str, str]:
    env_json = os.getenv("MARINADE_MODEL_MATRIX")
    if env_json:
        try:
            payload = json.loads(env_json)
            if isinstance(payload, dict):
                matrix = {str(k): str(v) for k, v in payload.items() if v is not None}
                return _apply_profile_overlay(matrix, profile or os.getenv("OPENAI_MODEL_PROFILE"))
        except Exception:
            pass

    path_env = os.getenv("MARINADE_MODEL_MATRIX_PATH")
    if path_env:
        loaded = _load_matrix_from_path(Path(path_env))
        if loaded:
            return _apply_profile_overlay(loaded, profile or os.getenv("OPENAI_MODEL_PROFILE"))

    default_path = Path(__file__).with_name("model_matrix.json")
    loaded = _load_matrix_from_path(default_path)
    if loaded:
        return _apply_profile_overlay(loaded, profile or os.getenv("OPENAI_MODEL_PROFILE"))

    return _apply_profile_overlay(
        DEFAULT_MODEL_MATRIX.copy(),
        profile or os.getenv("OPENAI_MODEL_PROFILE"),
    )


def model_for_purpose(purpose: str, *, default_model: str) -> str:
    matrix = load_model_matrix()
    return matrix.get(purpose) or matrix.get("default") or default_model


def resolve_model_selector(
    selector: str | None,
    *,
    purpose: str = "default",
    default_model: str = "gpt-5.4",
) -> str:
    """
    Resolve a model selector into a concrete model name.

    Supported selectors:
    - explicit model names (e.g., "gpt-5.4", "gpt-4o", "gpt-3.5-turbo")
    - profile aliases ("production", "prod", "r_and_d", "rnd", "r&d")
    - empty/None -> model matrix default for purpose
    """
    raw = str(selector or "").strip()
    if not raw:
        return model_for_purpose(purpose, default_model=default_model)

    normalized_profile = normalize_model_profile(raw)
    if normalized_profile:
        matrix = load_model_matrix(profile=normalized_profile)
        return matrix.get(purpose) or matrix.get("default") or default_model

    return raw
