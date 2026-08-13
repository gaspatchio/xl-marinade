# ABOUTME: LLM enrichment service with strict JSON contracts and fixture/OpenAI providers
# ABOUTME: Implements retry-with-feedback, validation, and audit logging for Sprint 7

"""
Enrichment Service Module (Sprint 7)
======================================

This module provides a single, reliable interface for LLM enrichment decisions
with strict JSON contracts, offline fixture mode, and comprehensive error handling.

## Providers

1. **FixtureProvider**: Deterministic JSON responses from fixture files
   - Powers integration tests and offline E2E use cases
   - No network/API key required
   - Produces exact, reproducible mutations

2. **OpenAIProvider**: Real LLM calls via OpenAI API
   - Reads model from OPENAI_MODEL env var (default: gpt-5.2)
   - Records model name in summary and audit logs
   - Requires OPENAI_API_KEY

## Prompt Contracts

### 1. Label Enrichment (enrich_label_v1)
Input: Formula-forward evidence (semantic_formula, formula_r1c1, token summary, representative_formulas)
Output: set_label or override_binding mutation proposal

### 2. Structural Fix (propose_structural_fix_v1)
Input: BindingStructuralContext v1 (neighbors, occupancy, deterministic candidates)
Output: One of merge_bindings/split_binding/disable_binding or "no change"

### 3. Segmentation Tag Proposal (propose_segmentation_tags_v1)
Input: Formula evidence + limited upstream/downstream context + allowed tag enum
Output: Segmentation tag (must be in allowed set)
Call policy: Executed for EVERY active binding (not triage-gated)

## Retry Policy (AC5)

For each binding:
- Attempt up to 3 calls
- On validation failure, retry with prior error included
- After 3 failures, record and continue (don't block pipeline)
- Invalid proposals NEVER appended to mutations.json

## Audit Logging (AC6)

Optional enrichment_audit.jsonl (--llm-audit-log flag):
- Per-attempt records: binding_id, stage, provider, prompt_id/version, attempt
- Redacted request/response summaries (bounded)
- Validation pass/fail + errors
- Mutation IDs appended + timing
- Model name (when provider=openai)

Default: OFF (no functional behavior change)

See: docs/phase2_documentation_agent/backlog/sprint7/STORY_sprint7_03_enrichment_service.md
"""

import json
import logging
import os
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Literal

logger = logging.getLogger(__name__)

# Prompt versioning
PROMPT_VERSION_LABEL_ENRICHMENT = "enrich_label_v1"
PROMPT_VERSION_STRUCTURAL_FIX = "propose_structural_fix_v1"
PROMPT_VERSION_SEGMENTATION_TAG = "propose_segmentation_tags_v1"
PROMPT_VERSION_TABLE_GROUP_MERGE = "propose_table_group_merge_v1"

# Retry configuration
MAX_RETRY_ATTEMPTS = 3

# Redaction placeholder
REDACTED_VALUE = "<REDACTED>"


# ============================================================================
# Data Classes (Prompt Contracts)
# ============================================================================


@dataclass
class LabelEnrichmentRequest:
    """Input contract for label enrichment prompt."""

    binding_id: str
    current_label: str | None
    semantic_formula: str  # Normalized, numeric literals redacted
    formula_r1c1: str  # Pattern
    function_tokens: dict[str, int]  # Function name -> count
    representative_formulas: list[str]  # N small examples with A1 addresses

    # Optional context
    parent_labels: list[str] = field(default_factory=list)
    child_labels: list[str] = field(default_factory=list)


@dataclass
class LabelEnrichmentResponse:
    """Output contract for label enrichment prompt."""

    action: Literal["set_label", "override_binding", "no_change"]
    new_label: str | None = None
    reasoning: str = ""
    confidence: float = 0.0  # 0.0-1.0

    def to_mutation_params(self, binding_id: str, old_label: str | None) -> dict[str, Any] | None:
        """Convert to mutation parameters."""
        if self.action == "no_change":
            return None

        if self.action == "set_label":
            return {"binding_id": binding_id, "old": old_label, "new": self.new_label}

        if self.action == "override_binding":
            return {"binding_id": binding_id, "old_label": old_label, "new_label": self.new_label}

        return None


@dataclass
class StructuralFixCandidate:
    """A deterministic structural mutation candidate."""

    mutation_type: Literal["merge_bindings", "split_binding", "disable_binding"]
    parameters: dict[str, Any]
    rationale: str


@dataclass
class StructuralFixRequest:
    """Input contract for structural fix prompt."""

    binding_id: str
    current_label: str | None

    # Binding context
    range_str: str
    cell_count: int
    formula_pattern: str

    # Neighbor context
    neighbor_bindings: list[dict[str, Any]]  # Simplified binding summaries

    # Occupancy encoding (for merge/split decisions)
    occupancy_grid: str | None = None

    # Deterministic candidates (LLM must choose one or "no change")
    candidates: list[StructuralFixCandidate] = field(default_factory=list)


@dataclass
class StructuralFixResponse:
    """Output contract for structural fix prompt."""

    action: Literal["merge_bindings", "split_binding", "disable_binding", "no_change"]
    selected_candidate_index: int | None = None  # Index into request.candidates
    reasoning: str = ""

    def to_mutation_params(self, request: StructuralFixRequest) -> dict[str, Any] | None:
        """Convert to mutation parameters."""
        if self.action == "no_change":
            return None

        if self.selected_candidate_index is None:
            raise ValueError("selected_candidate_index required when action != no_change")

        if not (0 <= self.selected_candidate_index < len(request.candidates)):
            raise ValueError(
                f"selected_candidate_index {self.selected_candidate_index} out of range"
            )

        candidate = request.candidates[self.selected_candidate_index]
        return candidate.parameters


@dataclass
class SegmentationTagRequest:
    """Input contract for segmentation tag proposal prompt."""

    binding_id: str
    current_label: str | None

    # Formula evidence
    semantic_formula: str
    formula_r1c1: str
    function_tokens: dict[str, int]

    # Dependency context (limited)
    upstream_binding_summaries: list[dict[str, Any]] = field(default_factory=list)
    downstream_binding_summaries: list[dict[str, Any]] = field(default_factory=list)

    # Allowed tag set (enum constraint)
    allowed_tags: list[str] = field(default_factory=list)


@dataclass
class SegmentationTagResponse:
    """Output contract for segmentation tag proposal prompt."""

    tag: str  # Must be in allowed_tags
    reasoning: str = ""
    confidence: float = 0.0


@dataclass
class TableGroupMergeRequest:
    """Input contract for table group merge prompt."""

    group_id: str  # Stable deterministic ID from clustering
    bindings: list[
        dict[str, Any]
    ]  # List of {binding_id, sheet, address_a1, kind, classification, label, bbox, sample_values}
    nearby_headers: list[dict[str, Any]]  # List of {cell_address, value, number_format} (bounded)
    constraints: dict[str, Any]  # {allow_formula_merges: bool, max_merge_groups: int}


@dataclass
class TableGroupMergeResponse:
    """Output contract for table group merge prompt."""

    action: Literal["no_change", "merge"]
    merge_groups: list[dict[str, Any]] = field(
        default_factory=list
    )  # List of {binding_ids: [...], confidence: float, rationale: str}
    warnings: list[str] = field(default_factory=list)


# ============================================================================
# Validation
# ============================================================================


def validate_label_enrichment_response(response: dict[str, Any]) -> LabelEnrichmentResponse:
    """Validate and parse label enrichment response."""
    required_fields = ["action", "reasoning"]
    for field_name in required_fields:
        if field_name not in response:
            raise ValueError(f"Missing required field: {field_name}")

    action = response["action"]
    if action not in ["set_label", "override_binding", "no_change"]:
        raise ValueError(f"Invalid action: {action}")

    if action != "no_change" and "new_label" not in response:
        raise ValueError(f"new_label required when action={action}")

    if action != "no_change" and not response.get("new_label"):
        raise ValueError("new_label cannot be empty")

    confidence = response.get("confidence", 0.0)
    if not isinstance(confidence, (int, float)) or not (0.0 <= confidence <= 1.0):
        raise ValueError(f"confidence must be float in [0.0, 1.0], got {confidence}")

    return LabelEnrichmentResponse(
        action=action,
        new_label=response.get("new_label"),
        reasoning=response["reasoning"],
        confidence=float(confidence),
    )


def validate_structural_fix_response(
    response: dict[str, Any], request: StructuralFixRequest
) -> StructuralFixResponse:
    """Validate and parse structural fix response."""
    required_fields = ["action", "reasoning"]
    for field_name in required_fields:
        if field_name not in response:
            raise ValueError(f"Missing required field: {field_name}")

    action = response["action"]
    if action not in ["merge_bindings", "split_binding", "disable_binding", "no_change"]:
        raise ValueError(f"Invalid action: {action}")

    if action != "no_change":
        if "selected_candidate_index" not in response:
            raise ValueError(f"selected_candidate_index required when action={action}")

        idx = response["selected_candidate_index"]
        if not isinstance(idx, int):
            raise ValueError(f"selected_candidate_index must be int, got {type(idx)}")

        if not (0 <= idx < len(request.candidates)):
            raise ValueError(
                f"selected_candidate_index {idx} out of range [0, {len(request.candidates)})"
            )

    return StructuralFixResponse(
        action=action,
        selected_candidate_index=response.get("selected_candidate_index"),
        reasoning=response["reasoning"],
    )


def validate_segmentation_tag_response(
    response: dict[str, Any], allowed_tags: list[str]
) -> SegmentationTagResponse:
    """Validate and parse segmentation tag response."""
    required_fields = ["tag", "reasoning"]
    for field_name in required_fields:
        if field_name not in response:
            raise ValueError(f"Missing required field: {field_name}")

    tag = response["tag"]
    if tag not in allowed_tags:
        raise ValueError(f"tag '{tag}' not in allowed set: {allowed_tags}")

    confidence = response.get("confidence", 0.0)
    if not isinstance(confidence, (int, float)) or not (0.0 <= confidence <= 1.0):
        raise ValueError(f"confidence must be float in [0.0, 1.0], got {confidence}")

    return SegmentationTagResponse(
        tag=tag, reasoning=response["reasoning"], confidence=float(confidence)
    )


def validate_table_group_merge_response(response: dict[str, Any]) -> TableGroupMergeResponse:
    """Validate and parse table group merge response."""
    required_fields = ["action"]
    for field_name in required_fields:
        if field_name not in response:
            raise ValueError(f"Missing required field: {field_name}")

    action = response["action"]
    if action not in ["no_change", "merge"]:
        raise ValueError(f"Invalid action: {action}")

    merge_groups = response.get("merge_groups", [])
    if action == "merge" and not merge_groups:
        raise ValueError("merge_groups required when action=merge")

    # Validate each merge group
    for i, group in enumerate(merge_groups):
        if "binding_ids" not in group:
            raise ValueError(f"merge_groups[{i}] missing binding_ids")
        if not isinstance(group["binding_ids"], list) or len(group["binding_ids"]) < 2:
            raise ValueError(f"merge_groups[{i}].binding_ids must be list with at least 2 IDs")
        if "confidence" not in group:
            raise ValueError(f"merge_groups[{i}] missing confidence")
        confidence = group["confidence"]
        if not isinstance(confidence, (int, float)) or not (0.0 <= confidence <= 1.0):
            raise ValueError(
                f"merge_groups[{i}].confidence must be float in [0.0, 1.0], got {confidence}"
            )
        if "rationale" not in group:
            raise ValueError(f"merge_groups[{i}] missing rationale")

    warnings = response.get("warnings", [])
    if not isinstance(warnings, list):
        raise ValueError("warnings must be list")

    return TableGroupMergeResponse(action=action, merge_groups=merge_groups, warnings=warnings)


# ============================================================================
# Provider Interface
# ============================================================================


class EnrichmentProvider(ABC):
    """Abstract interface for enrichment providers."""

    @abstractmethod
    def enrich_label(self, request: LabelEnrichmentRequest) -> dict[str, Any]:
        """Call label enrichment prompt."""
        pass

    @abstractmethod
    def propose_structural_fix(self, request: StructuralFixRequest) -> dict[str, Any]:
        """Call structural fix prompt."""
        pass

    @abstractmethod
    def propose_segmentation_tag(self, request: SegmentationTagRequest) -> dict[str, Any]:
        """Call segmentation tag prompt."""
        pass

    @abstractmethod
    def propose_table_group_merge(self, request: TableGroupMergeRequest) -> dict[str, Any]:
        """Call table group merge prompt."""
        pass

    @abstractmethod
    def get_provider_name(self) -> str:
        """Return provider name for audit logs."""
        pass


# ============================================================================
# Fixture Provider (Deterministic)
# ============================================================================


class FixtureProvider(EnrichmentProvider):
    """Deterministic fixture-backed provider for testing."""

    def __init__(self, fixture_dir: Path):
        self.fixture_dir = fixture_dir
        self._load_fixtures()

    def _load_fixtures(self) -> None:
        """Load fixture files."""
        self.label_fixtures: dict[str, dict[str, Any]] = {}
        self.structural_fixtures: dict[str, dict[str, Any]] = {}
        self.segmentation_fixtures: dict[str, dict[str, Any]] = {}
        self.table_group_merge_fixtures: dict[str, dict[str, Any]] = {}

        # Load label enrichment fixtures
        label_fixture_path = self.fixture_dir / "label_enrichment_fixtures.json"
        if label_fixture_path.exists():
            with open(label_fixture_path, encoding="utf-8") as f:
                data = json.load(f)
                self.label_fixtures = {item["binding_id"]: item["response"] for item in data}

        # Load structural fix fixtures
        structural_fixture_path = self.fixture_dir / "structural_fix_fixtures.json"
        if structural_fixture_path.exists():
            with open(structural_fixture_path, encoding="utf-8") as f:
                data = json.load(f)
                self.structural_fixtures = {item["binding_id"]: item["response"] for item in data}

        # Load segmentation tag fixtures
        segmentation_fixture_path = self.fixture_dir / "segmentation_tag_fixtures.json"
        if segmentation_fixture_path.exists():
            with open(segmentation_fixture_path, encoding="utf-8") as f:
                data = json.load(f)
                self.segmentation_fixtures = {item["binding_id"]: item["response"] for item in data}

        # Load table group merge fixtures
        table_group_merge_fixture_path = self.fixture_dir / "table_group_merge_fixtures.json"
        if table_group_merge_fixture_path.exists():
            with open(table_group_merge_fixture_path, encoding="utf-8") as f:
                data = json.load(f)
                self.table_group_merge_fixtures = {
                    item["group_id"]: item["response"] for item in data
                }

    def enrich_label(self, request: LabelEnrichmentRequest) -> dict[str, Any]:
        """Return fixture response for label enrichment."""
        if request.binding_id in self.label_fixtures:
            return self.label_fixtures[request.binding_id]

        # Default: no change
        return {
            "action": "no_change",
            "reasoning": "No fixture available for this binding",
            "confidence": 0.0,
        }

    def propose_structural_fix(self, request: StructuralFixRequest) -> dict[str, Any]:
        """Return fixture response for structural fix."""
        if request.binding_id in self.structural_fixtures:
            return self.structural_fixtures[request.binding_id]

        # Default: no change
        return {"action": "no_change", "reasoning": "No fixture available for this binding"}

    def propose_segmentation_tag(self, request: SegmentationTagRequest) -> dict[str, Any]:
        """Return fixture response for segmentation tag."""
        if request.binding_id in self.segmentation_fixtures:
            return self.segmentation_fixtures[request.binding_id]

        # Default: deterministic heuristic fallback (must remain within allowed_tags)
        allowed = set(request.allowed_tags or [])
        if allowed:
            tokens = {k.upper(): int(v) for k, v in (request.function_tokens or {}).items()}
            label = (request.current_label or "").lower()

            lookup_funcs = {"VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH"}
            calc_funcs = {
                "SUM",
                "AVERAGE",
                "MIN",
                "MAX",
                "ROUND",
                "IF",
                "IFS",
                "COUNT",
                "COUNTA",
                "COUNTIF",
                "COUNTIFS",
            }

            if lookup_funcs & set(tokens.keys()) and "Lookup" in allowed:
                return {
                    "tag": "Lookup",
                    "reasoning": "Heuristic fallback: lookup-style functions detected",
                    "confidence": 0.25,
                }

            if (calc_funcs & set(tokens.keys())) and "Calculation" in allowed:
                return {
                    "tag": "Calculation",
                    "reasoning": "Heuristic fallback: calculation-style functions detected",
                    "confidence": 0.25,
                }

            if (
                any(
                    word in label
                    for word in ("result", "output", "total", "net", "gross", "profit")
                )
                and "Output" in allowed
            ):
                return {
                    "tag": "Output",
                    "reasoning": "Heuristic fallback: label suggests an output/result metric",
                    "confidence": 0.20,
                }

            if (
                any(
                    word in label for word in ("input", "assumption", "rate", "factor", "parameter")
                )
                and "Input" in allowed
            ):
                return {
                    "tag": "Input",
                    "reasoning": "Heuristic fallback: label suggests an input/assumption",
                    "confidence": 0.20,
                }

            # Stable fallback order
            for tag in ("Input", "Calculation", "Output", "Lookup"):
                if tag in allowed:
                    return {
                        "tag": tag,
                        "reasoning": "Heuristic fallback: default tag selection",
                        "confidence": 0.10,
                    }

        raise ValueError("No allowed_tags provided and no fixture available")

    def propose_table_group_merge(self, request: TableGroupMergeRequest) -> dict[str, Any]:
        """Return fixture response for table group merge."""
        if request.group_id in self.table_group_merge_fixtures:
            return self.table_group_merge_fixtures[request.group_id]

        # Default: no change
        return {
            "action": "no_change",
            "merge_groups": [],
            "warnings": ["No fixture available for this group"],
        }

    def get_provider_name(self) -> str:
        return "fixture"


# ============================================================================
# OpenAI Provider
# ============================================================================


class OpenAIProvider(EnrichmentProvider):
    """Real LLM provider using OpenAI API."""

    def __init__(self, api_key: str | None = None, model: str | None = None):
        from xl_marinade.llm.factory import resolve_api_key

        self.api_key = resolve_api_key(api_key)
        if not self.api_key:
            raise ValueError("No LLM API key set (LLM_API_KEY or OPENAI_API_KEY)")

        from xl_marinade.docs.model_matrix import load_model_matrix, resolve_model_selector

        self.model_override = None
        if model:
            self.model_override = resolve_model_selector(
                model,
                purpose="default",
                default_model="gpt-5.2",
            )

        # Read default model from env or use matrix default.
        self.model_default = resolve_model_selector(
            os.getenv("OPENAI_MODEL"),
            purpose="default",
            default_model="gpt-5.2",
        )
        self.model_matrix = load_model_matrix()

        # Construct the client through the one BYOK seam (honours LLM_BASE_URL).
        from xl_marinade.llm.factory import make_llm_client

        self.client = make_llm_client(api_key=self.api_key)

    def _model_for_purpose(self, purpose: str) -> str:
        if self.model_override:
            return self.model_override
        return (
            self.model_matrix.get(purpose) or self.model_matrix.get("default") or self.model_default
        )

    def _call_openai(self, system_prompt: str, user_prompt: str, *, purpose: str) -> dict[str, Any]:
        """Call OpenAI API and parse a JSON object response."""
        from xl_marinade.llm.llm_usage_logger import log_llm_usage

        model_name = self._model_for_purpose(purpose)
        used_fallback = False
        try:
            response = self.client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                response_format={"type": "json_object"},
                temperature=0.0,
            )
        except Exception as e:
            message = str(e)
            if "response_format" in message and "json_object" in message:
                used_fallback = True
                response = self.client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=0.0,
                )
            else:
                raise

        log_llm_usage(
            purpose=purpose,
            model=model_name,
            usage=getattr(response, "usage", None),
            extra={"response_format": "json_object", "fallback": used_fallback},
        )

        content = response.choices[0].message.content
        if not content:
            raise ValueError("Empty response from OpenAI")

        try:
            return json.loads(content)
        except json.JSONDecodeError:
            start = content.find("{")
            end = content.rfind("}")
            if start != -1 and end != -1 and end > start:
                return json.loads(content[start : end + 1])
            raise

    def enrich_label(self, request: LabelEnrichmentRequest) -> dict[str, Any]:
        """Call OpenAI for label enrichment."""
        system_prompt = """You are an expert actuarial analyst reviewing Excel model labels.
Your task is to improve ambiguous or generic labels using formula evidence.

Return JSON with:
{
  "action": "set_label" | "override_binding" | "no_change",
  "new_label": "Improved Label" (if action != no_change),
  "reasoning": "Why this label is better",
  "confidence": 0.0-1.0
}

Guidelines:
- Use actuarial terminology when appropriate
- Prefer specific over generic labels
- Base decisions on formula patterns and token usage
- If current label is already good, return "no_change"
"""

        user_prompt = f"""Current label: {request.current_label or "(none)"}

Formula evidence:
- Semantic formula: {request.semantic_formula}
- R1C1 pattern: {request.formula_r1c1}
- Function tokens: {json.dumps(request.function_tokens)}
- Representative formulas: {json.dumps(request.representative_formulas[:5])}

Parent labels: {json.dumps(request.parent_labels[:5])}
Child labels: {json.dumps(request.child_labels[:5])}

Propose an improved label or return "no_change".
"""

        return self._call_openai(system_prompt, user_prompt, purpose="label_enrich")

    def propose_structural_fix(self, request: StructuralFixRequest) -> dict[str, Any]:
        """Call OpenAI for structural fix."""
        system_prompt = """You are an expert actuarial analyst reviewing Excel model binding structure.
Your task is to select the best structural fix from a set of deterministic candidates.

Return JSON with:
{
  "action": "merge_bindings" | "split_binding" | "disable_binding" | "no_change",
  "selected_candidate_index": 0-based index (if action != no_change),
  "reasoning": "Why this fix improves the model structure"
}

Guidelines:
- You MUST choose from the provided candidates (no arbitrary ranges)
- Consider cell conservation (don't lose data)
- Prefer fixes that improve auditability
- If no candidate improves the structure, return "no_change"
"""

        candidates_str = "\n".join(
            [f"{i}. {c.mutation_type}: {c.rationale}" for i, c in enumerate(request.candidates)]
        )

        user_prompt = f"""Binding: {request.binding_id}
Label: {request.current_label or "(none)"}
Range: {request.range_str} ({request.cell_count} cells)
Formula pattern: {request.formula_pattern}

Neighbor bindings: {json.dumps(request.neighbor_bindings[:10], indent=2)}

Candidates:
{candidates_str}

Select the best candidate (by index) or return "no_change".
"""

        return self._call_openai(system_prompt, user_prompt, purpose="structural_fix")

    def propose_segmentation_tag(self, request: SegmentationTagRequest) -> dict[str, Any]:
        """Call OpenAI for segmentation tag."""
        system_prompt = """You are an expert actuarial analyst reviewing Excel model segmentation.
Your task is to assign a segmentation tag from the allowed set.

Return JSON with:
{
  "tag": "one of the allowed tags",
  "reasoning": "Why this tag fits the binding's role",
  "confidence": 0.0-1.0
}

Guidelines:
- You MUST choose from the allowed_tags list
- Consider formula patterns and dependency context
- Use actuarial domain knowledge
"""

        user_prompt = f"""Binding: {request.binding_id}
Label: {request.current_label or "(none)"}

Formula evidence:
- Semantic formula: {request.semantic_formula}
- R1C1 pattern: {request.formula_r1c1}
- Function tokens: {json.dumps(request.function_tokens)}

Upstream bindings: {json.dumps(request.upstream_binding_summaries[:5])}
Downstream bindings: {json.dumps(request.downstream_binding_summaries[:5])}

Allowed tags: {json.dumps(request.allowed_tags)}

Select the most appropriate tag.
"""

        return self._call_openai(system_prompt, user_prompt, purpose="segmentation_tag")

    def propose_table_group_merge(self, request: TableGroupMergeRequest) -> dict[str, Any]:
        """Call OpenAI for table group merge."""
        system_prompt = """You are an expert actuarial analyst reviewing Excel model table structures.
Your task is to propose merges for fragmented table components that should be treated as coherent tables.

Return JSON with:
{
  "action": "no_change" | "merge",
  "merge_groups": [
    {
      "binding_ids": ["id1", "id2", ...],
      "confidence": 0.0-1.0,
      "rationale": "Why these bindings form a coherent table"
    }
  ],
  "warnings": ["Optional warnings about ambiguous cases"]
}

Guidelines:
- Only propose merges when evidence strongly suggests the bindings are parts of the same logical table
- Consider: aligned row/column structure, consistent headers, semantic coherence
- Prefer conservative merges (high confidence) over aggressive merges
- If evidence is insufficient, return "no_change"
- Respect constraints (e.g., allow_formula_merges flag)
"""

        bindings_str = json.dumps(request.bindings[:20], indent=2)  # Limit to 20 for prompt size
        headers_str = json.dumps(request.nearby_headers[:50], indent=2)  # Limit to 50

        user_prompt = f"""Group ID: {request.group_id}

Bindings in group:
{bindings_str}

Nearby headers:
{headers_str}

Constraints:
{json.dumps(request.constraints)}

Propose merge groups or return "no_change".
"""

        return self._call_openai(system_prompt, user_prompt, purpose="table_group_merge")

    def get_provider_name(self) -> str:
        return "openai"

    def get_model_name(self) -> str:
        """Return the resolved model name."""
        return self.model_matrix.get("default") or self.model_default


# ============================================================================
# Enrichment Service (Main Interface)
# ============================================================================


@dataclass
class EnrichmentAttempt:
    """Record of a single enrichment attempt."""

    binding_id: str
    stage: str  # "label_enrichment", "structural_fix", "segmentation_tag"
    attempt_number: int
    timestamp: str
    provider: str
    prompt_id: str
    prompt_version: str

    # Request/response (redacted)
    request_summary: dict[str, Any]
    response_summary: dict[str, Any]

    # Validation
    validation_passed: bool
    validation_error: str | None = None

    # Outcome
    mutation_ids_appended: list[int] = field(default_factory=list)
    duration_ms: float = 0.0

    # Model name (only for openai provider)
    model: str | None = None


class EnrichmentService:
    """Main enrichment service with retry-with-feedback and audit logging."""

    def __init__(self, provider: EnrichmentProvider, audit_log_path: Path | None = None):
        self.provider = provider
        self.audit_log_path = audit_log_path
        self.audit_enabled = audit_log_path is not None

        # Initialize audit log file.
        #
        # Important: do not truncate here. The Sprint 7 pipeline creates one
        # EnrichmentService per stage (label/structural/segmentation), and all
        # stages should append to the same audit file for a run. Runners that
        # need a fresh log should delete the file before creating the service.
        if self.audit_enabled and self.audit_log_path:
            self.audit_log_path.parent.mkdir(parents=True, exist_ok=True)
            self.audit_log_path.touch(exist_ok=True)

    def _write_audit_record(self, attempt: EnrichmentAttempt) -> None:
        """Write audit record to JSONL file."""
        if not self.audit_enabled or not self.audit_log_path:
            return

        record = {
            "binding_id": attempt.binding_id,
            "stage": attempt.stage,
            "attempt_number": attempt.attempt_number,
            "timestamp": attempt.timestamp,
            "provider": attempt.provider,
            "prompt_id": attempt.prompt_id,
            "prompt_version": attempt.prompt_version,
            "request_summary": attempt.request_summary,
            "response_summary": attempt.response_summary,
            "validation_passed": attempt.validation_passed,
            "validation_error": attempt.validation_error,
            "mutation_ids_appended": attempt.mutation_ids_appended,
            "duration_ms": attempt.duration_ms,
        }

        # Add model name if available
        if attempt.model:
            record["model"] = attempt.model

        with open(self.audit_log_path, "a", encoding="utf-8", newline="\n") as f:
            f.write(json.dumps(record) + "\n")

    def _redact_request(self, request: Any) -> dict[str, Any]:
        """Create redacted summary of request."""
        if isinstance(request, LabelEnrichmentRequest):
            return {
                "binding_id": request.binding_id,
                "current_label": request.current_label,
                "semantic_formula": request.semantic_formula[:100],  # Truncate
                "function_tokens": request.function_tokens,
            }
        elif isinstance(request, StructuralFixRequest):
            return {
                "binding_id": request.binding_id,
                "current_label": request.current_label,
                "range_str": request.range_str,
                "cell_count": request.cell_count,
                "candidate_count": len(request.candidates),
            }
        elif isinstance(request, SegmentationTagRequest):
            return {
                "binding_id": request.binding_id,
                "current_label": request.current_label,
                "allowed_tags": request.allowed_tags,
            }
        elif isinstance(request, TableGroupMergeRequest):
            return {
                "group_id": request.group_id,
                "binding_count": len(request.bindings),
                "nearby_header_count": len(request.nearby_headers),
                "constraints": request.constraints,
            }
        else:
            return {"type": type(request).__name__}

    def _redact_response(self, response: dict[str, Any]) -> dict[str, Any]:
        """Create redacted summary of response."""
        # Limit response size
        summary = {}
        for key, value in response.items():
            if isinstance(value, str) and len(value) > 200:
                summary[key] = value[:200] + "..."
            else:
                summary[key] = value
        return summary

    def enrich_label_with_retry(
        self, request: LabelEnrichmentRequest, mutation_logger: Any
    ) -> tuple[int | None, list[str]]:
        """
        Enrich label with retry-with-feedback.

        Returns:
            (mutation_id or None, list of validation errors)
        """
        stage = "label_enrichment"
        errors: list[str] = []
        model_name = None

        if isinstance(self.provider, OpenAIProvider):
            model_name = self.provider.get_model_name()

        for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
            start_time = time.time()

            try:
                # Call provider
                raw_response = self.provider.enrich_label(request)

                # Validate response
                validated = validate_label_enrichment_response(raw_response)

                # Convert to mutation
                mutation_params = validated.to_mutation_params(
                    request.binding_id, request.current_label
                )

                duration_ms = (time.time() - start_time) * 1000

                if mutation_params is None:
                    # No change
                    if self.audit_enabled:
                        self._write_audit_record(
                            EnrichmentAttempt(
                                binding_id=request.binding_id,
                                stage=stage,
                                attempt_number=attempt,
                                timestamp=datetime.now(UTC).isoformat(),
                                provider=self.provider.get_provider_name(),
                                prompt_id="enrich_label",
                                prompt_version=PROMPT_VERSION_LABEL_ENRICHMENT,
                                request_summary=self._redact_request(request),
                                response_summary=self._redact_response(raw_response),
                                validation_passed=True,
                                mutation_ids_appended=[],
                                duration_ms=duration_ms,
                                model=model_name,
                            )
                        )
                    return (None, [])

                # Append mutation
                if validated.action == "set_label":
                    if hasattr(mutation_logger, "append_mutation"):
                        metadata: dict[str, Any] = {
                            "reasoning": validated.reasoning,
                            "knowledge_source": "llm",
                            "sprint": 7,
                            "provider": self.provider.get_provider_name(),
                            "prompt_id": "enrich_label",
                            "prompt_version": PROMPT_VERSION_LABEL_ENRICHMENT,
                            "confidence_initial": validated.confidence,
                        }
                        parameters: dict[str, Any] = {
                            "binding_id": mutation_params["binding_id"],
                            "old": mutation_params["old"],
                            "new": mutation_params["new"],
                            "confidence": validated.confidence,
                        }
                        mutation_id = mutation_logger.append_mutation(
                            "set_label",
                            parameters,
                            metadata,
                        )
                    else:
                        mutation_id = mutation_logger.set_label(
                            binding_id=mutation_params["binding_id"],
                            old_label=mutation_params["old"],
                            new_label=mutation_params["new"],
                            reasoning=validated.reasoning,
                            knowledge_source="llm",
                            confidence=validated.confidence,
                        )
                elif validated.action == "override_binding":
                    if hasattr(mutation_logger, "append_mutation"):
                        metadata = {
                            "reasoning": validated.reasoning,
                            "knowledge_source": "llm",
                            "sprint": 7,
                            "provider": self.provider.get_provider_name(),
                            "prompt_id": "enrich_label",
                            "prompt_version": PROMPT_VERSION_LABEL_ENRICHMENT,
                            "confidence_initial": validated.confidence,
                        }
                        parameters = {
                            "binding_id": mutation_params["binding_id"],
                            "old_label": mutation_params["old_label"],
                            "new_label": mutation_params["new_label"],
                        }
                        mutation_id = mutation_logger.append_mutation(
                            "override_binding",
                            parameters,
                            metadata,
                        )
                    else:
                        mutation_id = mutation_logger.override_binding(
                            binding_id=mutation_params["binding_id"],
                            old_label=mutation_params["old_label"],
                            new_label=mutation_params["new_label"],
                            actuarial_class=None,
                            reasoning=validated.reasoning,
                        )
                else:
                    raise ValueError(f"Unexpected action: {validated.action}")

                # Success
                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.binding_id,
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="enrich_label",
                            prompt_version=PROMPT_VERSION_LABEL_ENRICHMENT,
                            request_summary=self._redact_request(request),
                            response_summary=self._redact_response(raw_response),
                            validation_passed=True,
                            mutation_ids_appended=[mutation_id],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                return (mutation_id, [])

            except Exception as e:
                error_msg = str(e)
                errors.append(f"Attempt {attempt}: {error_msg}")

                duration_ms = (time.time() - start_time) * 1000

                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.binding_id,
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="enrich_label",
                            prompt_version=PROMPT_VERSION_LABEL_ENRICHMENT,
                            request_summary=self._redact_request(request),
                            response_summary={},
                            validation_passed=False,
                            validation_error=error_msg,
                            mutation_ids_appended=[],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                if attempt < MAX_RETRY_ATTEMPTS:
                    logger.warning(
                        f"Label enrichment attempt {attempt} failed for {request.binding_id}: {error_msg}. Retrying..."
                    )
                else:
                    logger.error(
                        f"Label enrichment failed after {MAX_RETRY_ATTEMPTS} attempts for {request.binding_id}"
                    )

        return (None, errors)

    def propose_structural_fix_with_retry(
        self, request: StructuralFixRequest, mutation_logger: Any
    ) -> tuple[int | None, list[str]]:
        """
        Propose structural fix with retry-with-feedback.

        Returns:
            (mutation_id or None, list of validation errors)
        """
        stage = "structural_fix"
        errors: list[str] = []
        model_name = None

        if isinstance(self.provider, OpenAIProvider):
            model_name = self.provider.get_model_name()

        for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
            start_time = time.time()

            try:
                # Call provider
                raw_response = self.provider.propose_structural_fix(request)

                # Validate response
                validated = validate_structural_fix_response(raw_response, request)

                # Convert to mutation
                mutation_params = validated.to_mutation_params(request)

                duration_ms = (time.time() - start_time) * 1000

                if mutation_params is None:
                    # No change
                    if self.audit_enabled:
                        self._write_audit_record(
                            EnrichmentAttempt(
                                binding_id=request.binding_id,
                                stage=stage,
                                attempt_number=attempt,
                                timestamp=datetime.now(UTC).isoformat(),
                                provider=self.provider.get_provider_name(),
                                prompt_id="propose_structural_fix",
                                prompt_version=PROMPT_VERSION_STRUCTURAL_FIX,
                                request_summary=self._redact_request(request),
                                response_summary=self._redact_response(raw_response),
                                validation_passed=True,
                                mutation_ids_appended=[],
                                duration_ms=duration_ms,
                                model=model_name,
                            )
                        )
                    return (None, [])

                # Append mutation (delegate to mutation_logger based on action)
                if validated.action == "merge_bindings":
                    if hasattr(mutation_logger, "append_mutation"):
                        metadata: dict[str, Any] = {
                            "reasoning": validated.reasoning,
                            "knowledge_source": "llm",
                            "sprint": 7,
                            "provider": self.provider.get_provider_name(),
                            "prompt_id": "propose_structural_fix",
                            "prompt_version": PROMPT_VERSION_STRUCTURAL_FIX,
                        }
                        mutation_id = mutation_logger.append_mutation(
                            "merge_bindings",
                            mutation_params,
                            metadata,
                        )
                    else:
                        mutation_id = mutation_logger.merge_bindings(
                            source_binding_ids=mutation_params["source_binding_ids"],
                            new_binding_id=mutation_params["new_binding_id"],
                            label=mutation_params["label"],
                            reasoning=validated.reasoning,
                        )
                elif validated.action == "split_binding":
                    if hasattr(mutation_logger, "append_mutation"):
                        metadata = {
                            "reasoning": validated.reasoning,
                            "knowledge_source": "llm",
                            "sprint": 7,
                            "provider": self.provider.get_provider_name(),
                            "prompt_id": "propose_structural_fix",
                            "prompt_version": PROMPT_VERSION_STRUCTURAL_FIX,
                        }
                        mutation_id = mutation_logger.append_mutation(
                            "split_binding",
                            mutation_params,
                            metadata,
                        )
                    else:
                        mutation_id = mutation_logger.split_binding(
                            source_binding_id=mutation_params["source_binding_id"],
                            new_bindings=mutation_params["new_bindings"],
                            reasoning=validated.reasoning,
                            knowledge_source="llm",
                        )
                elif validated.action == "disable_binding":
                    if hasattr(mutation_logger, "append_mutation"):
                        metadata = {
                            "reasoning": validated.reasoning,
                            "knowledge_source": "llm",
                            "sprint": 7,
                            "provider": self.provider.get_provider_name(),
                            "prompt_id": "propose_structural_fix",
                            "prompt_version": PROMPT_VERSION_STRUCTURAL_FIX,
                        }
                        mutation_id = mutation_logger.append_mutation(
                            "disable_binding",
                            mutation_params,
                            metadata,
                        )
                    else:
                        mutation_id = mutation_logger.disable_binding(
                            binding_id=mutation_params["binding_id"],
                            reason=mutation_params["reason"],
                        )
                else:
                    raise ValueError(f"Unexpected action: {validated.action}")

                # Success
                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.binding_id,
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="propose_structural_fix",
                            prompt_version=PROMPT_VERSION_STRUCTURAL_FIX,
                            request_summary=self._redact_request(request),
                            response_summary=self._redact_response(raw_response),
                            validation_passed=True,
                            mutation_ids_appended=[mutation_id],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                return (mutation_id, [])

            except Exception as e:
                error_msg = str(e)
                errors.append(f"Attempt {attempt}: {error_msg}")

                duration_ms = (time.time() - start_time) * 1000

                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.binding_id,
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="propose_structural_fix",
                            prompt_version=PROMPT_VERSION_STRUCTURAL_FIX,
                            request_summary=self._redact_request(request),
                            response_summary={},
                            validation_passed=False,
                            validation_error=error_msg,
                            mutation_ids_appended=[],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                if attempt < MAX_RETRY_ATTEMPTS:
                    logger.warning(
                        f"Structural fix attempt {attempt} failed for {request.binding_id}: {error_msg}. Retrying..."
                    )
                else:
                    logger.error(
                        f"Structural fix failed after {MAX_RETRY_ATTEMPTS} attempts for {request.binding_id}"
                    )

        return (None, errors)

    def propose_segmentation_tag_with_retry(
        self, request: SegmentationTagRequest
    ) -> tuple[SegmentationTagResponse | None, list[str]]:
        """
        Propose segmentation tag with retry-with-feedback.

        Returns:
            (SegmentationTagResponse or None, list of validation errors)
        """
        stage = "segmentation_tag"
        errors: list[str] = []
        model_name = None

        if isinstance(self.provider, OpenAIProvider):
            model_name = self.provider.get_model_name()

        for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
            start_time = time.time()

            try:
                # Call provider
                raw_response = self.provider.propose_segmentation_tag(request)

                # Validate response
                validated = validate_segmentation_tag_response(raw_response, request.allowed_tags)

                duration_ms = (time.time() - start_time) * 1000

                # Success
                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.binding_id,
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="propose_segmentation_tag",
                            prompt_version=PROMPT_VERSION_SEGMENTATION_TAG,
                            request_summary=self._redact_request(request),
                            response_summary=self._redact_response(raw_response),
                            validation_passed=True,
                            mutation_ids_appended=[],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                return (validated, [])

            except Exception as e:
                error_msg = str(e)
                errors.append(f"Attempt {attempt}: {error_msg}")

                duration_ms = (time.time() - start_time) * 1000

                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.binding_id,
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="propose_segmentation_tag",
                            prompt_version=PROMPT_VERSION_SEGMENTATION_TAG,
                            request_summary=self._redact_request(request),
                            response_summary={},
                            validation_passed=False,
                            validation_error=error_msg,
                            mutation_ids_appended=[],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                if attempt < MAX_RETRY_ATTEMPTS:
                    logger.warning(
                        f"Segmentation tag attempt {attempt} failed for {request.binding_id}: {error_msg}. Retrying..."
                    )
                else:
                    logger.error(
                        f"Segmentation tag failed after {MAX_RETRY_ATTEMPTS} attempts for {request.binding_id}"
                    )

        return (None, errors)

    def propose_table_group_merge_with_retry(
        self, request: TableGroupMergeRequest
    ) -> tuple[TableGroupMergeResponse | None, list[str]]:
        """
        Propose table group merge with retry-with-feedback.

        Returns:
            (TableGroupMergeResponse or None, list of validation errors)
        """
        stage = "table_group_merge"
        errors: list[str] = []
        model_name = None

        if isinstance(self.provider, OpenAIProvider):
            model_name = self.provider.get_model_name()

        for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
            start_time = time.time()

            try:
                # Call provider
                raw_response = self.provider.propose_table_group_merge(request)

                # Validate response
                validated = validate_table_group_merge_response(raw_response)

                duration_ms = (time.time() - start_time) * 1000

                # Success
                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.group_id,  # Use group_id as identifier
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="propose_table_group_merge",
                            prompt_version=PROMPT_VERSION_TABLE_GROUP_MERGE,
                            request_summary=self._redact_request(request),
                            response_summary=self._redact_response(raw_response),
                            validation_passed=True,
                            mutation_ids_appended=[],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                return (validated, [])

            except Exception as e:
                error_msg = str(e)
                errors.append(f"Attempt {attempt}: {error_msg}")

                duration_ms = (time.time() - start_time) * 1000

                if self.audit_enabled:
                    self._write_audit_record(
                        EnrichmentAttempt(
                            binding_id=request.group_id,
                            stage=stage,
                            attempt_number=attempt,
                            timestamp=datetime.now(UTC).isoformat(),
                            provider=self.provider.get_provider_name(),
                            prompt_id="propose_table_group_merge",
                            prompt_version=PROMPT_VERSION_TABLE_GROUP_MERGE,
                            request_summary=self._redact_request(request),
                            response_summary={},
                            validation_passed=False,
                            validation_error=error_msg,
                            mutation_ids_appended=[],
                            duration_ms=duration_ms,
                            model=model_name,
                        )
                    )

                if attempt < MAX_RETRY_ATTEMPTS:
                    logger.warning(
                        f"Table group merge attempt {attempt} failed for {request.group_id}: {error_msg}. Retrying..."
                    )
                else:
                    logger.error(
                        f"Table group merge failed after {MAX_RETRY_ATTEMPTS} attempts for {request.group_id}"
                    )

        return (None, errors)
