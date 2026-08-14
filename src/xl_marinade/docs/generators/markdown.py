# ABOUTME: Markdown documentation generator using Jinja2 templates
# ABOUTME: Transforms JSON spec into human-readable actuarial documentation

import json
import logging
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from xl_marinade.docs.utils.formatting import (
    format_table_value,
    format_time_series,
    format_value_for_display,
)

logger = logging.getLogger(__name__)


def _log_timing(step: str, elapsed: float, extra: str = "") -> None:
    suffix = f" ({extra})" if extra else ""
    logger.info(f"TIMING markdown_generator.{step}: {elapsed:.2f}s{suffix}")


# Optional dependency for templating
try:
    from jinja2 import Environment, FileSystemLoader, Template

    JINJA2_AVAILABLE = True
except ImportError:
    JINJA2_AVAILABLE = False
    Environment = None  # type: ignore
    FileSystemLoader = None  # type: ignore
    Template = None  # type: ignore

# Template directory
TEMPLATE_DIR = Path(__file__).parent.parent / "templates"


class MarkdownGenerator:
    """
    Generate human-readable Markdown documentation from JSON spec.

    Uses Jinja2 templates with deterministic tables and structured sections.
    Supports custom templates and includes default actuarial template.
    """

    def __init__(self, template_path: Path | str | None = None):
        """
        Initialize markdown generator.

        Args:
            template_path: Path to custom Jinja2 template (optional).
                          If None, uses default standard_actuarial.md.j2

        Raises:
            ImportError: If jinja2 not installed
            FileNotFoundError: If template file not found
        """
        if not JINJA2_AVAILABLE:
            raise ImportError(
                "jinja2 library required for markdown generation. Install with: pip install jinja2"
            )

        # Resolve template path
        if template_path is None:
            self.template_path = TEMPLATE_DIR / "standard_actuarial.md.j2"
        else:
            self.template_path = Path(template_path)

        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")

        # Setup Jinja2 environment
        if template_path is None:
            # Use default template from templates directory
            self.env = Environment(
                loader=FileSystemLoader(str(TEMPLATE_DIR)), trim_blocks=True, lstrip_blocks=True
            )
            # Register formatting filters
            self.env.filters["format_value"] = format_value_for_display
            self.env.filters["format_table_value"] = format_table_value
            self.env.filters["format_time_series"] = format_time_series

            self.template = self.env.get_template("standard_actuarial.md.j2")
        else:
            # Custom template - load directly
            template_dir = self.template_path.parent
            template_name = self.template_path.name
            self.env = Environment(
                loader=FileSystemLoader(str(template_dir)), trim_blocks=True, lstrip_blocks=True
            )
            # Register formatting filters
            self.env.filters["format_value"] = format_value_for_display
            self.env.filters["format_table_value"] = format_table_value
            self.env.filters["format_time_series"] = format_time_series

            self.template = self.env.get_template(template_name)

        logger.info(f"Initialized MarkdownGenerator with template: {self.template_path}")

    def generate(self, spec_json: dict[str, Any] | Path | str) -> str:
        """
        Generate markdown documentation from JSON spec.

        Args:
            spec_json: Either dict with spec data, or path to JSON file

        Returns:
            Generated markdown string

        Raises:
            ValueError: If spec invalid or missing required fields
        """
        start_time = time.perf_counter()
        # Load spec if path provided
        if isinstance(spec_json, Path | str):
            spec_path = Path(spec_json)
            if not spec_path.exists():
                raise FileNotFoundError(f"Spec file not found: {spec_path}")

            load_start = time.perf_counter()
            with open(spec_path, encoding="utf-8") as f:
                spec = json.load(f)
            _log_timing("load_spec", time.perf_counter() - load_start)
        else:
            spec = spec_json

        # Validate required fields
        if "metadata" not in spec:
            raise ValueError("Spec missing 'metadata' section")
        if "variables" not in spec:
            raise ValueError("Spec missing 'variables' section")

        # Prepare template context
        context_start = time.perf_counter()
        context = self._prepare_context(spec)
        _log_timing("prepare_context", time.perf_counter() - context_start)

        # Render template
        render_start = time.perf_counter()
        markdown = self.template.render(**context)
        _log_timing("render_template", time.perf_counter() - render_start)

        _log_timing("generate_total", time.perf_counter() - start_time)
        return markdown

    def _prepare_context(self, spec: dict[str, Any]) -> dict[str, Any]:
        """
        Prepare template context from spec.

        Organizes variables by actuarial class and adds metadata.

        Args:
            spec: JSON spec dict

        Returns:
            Context dict for template rendering
        """
        metadata = spec["metadata"]
        variables = spec["variables"]

        # Group variables by actuarial class
        grouped: dict[str, list[dict[str, Any]]] = {
            "Assumption": [],
            "Policyholder Data": [],
            "Calculation": [],
            "Result": [],
            "Index Lookup": [],
            "Unclassified": [],
        }

        for var in variables:
            actuarial_class = var.get("actuarial_class") or "Unclassified"
            if actuarial_class not in grouped:
                grouped[actuarial_class] = []
            grouped[actuarial_class].append(var)

        # Build context
        context = {
            "metadata": metadata,
            "variables": variables,
            "variables_by_class": grouped,
            "generation_timestamp": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
            "total_variables": len(variables),
            "assumption_count": len(grouped["Assumption"]),
            "policyholder_data_count": len(grouped["Policyholder Data"]),
            "calculation_count": len(grouped["Calculation"]),
            "result_count": len(grouped["Result"]),
            "index_lookup_count": len(grouped.get("Index Lookup", [])),
            "unclassified_count": len(grouped["Unclassified"]),
            "spec": spec,  # Pass full spec for reconciliation scope access
        }

        return context

    def generate_to_file(
        self, spec_json: dict[str, Any] | Path | str, output_path: Path | str
    ) -> None:
        """
        Generate markdown and write to file.

        Args:
            spec_json: Either dict with spec data, or path to JSON file
            output_path: Where to write markdown file
        """
        start_time = time.perf_counter()
        markdown = self.generate(spec_json)

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)

        write_start = time.perf_counter()
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(markdown)
        _log_timing("write_markdown", time.perf_counter() - write_start)

        logger.info(f"Generated markdown documentation: {output_file}")
        _log_timing("generate_to_file_total", time.perf_counter() - start_time)


def generate_markdown(spec_path: str, output_path: str, template_path: str | None = None) -> None:
    """
    Convenience function to generate markdown documentation.

    Args:
        spec_path: Path to model_spec.json file
        output_path: Where to write markdown file
        template_path: Optional custom template path
    """
    generator = MarkdownGenerator(template_path)
    generator.generate_to_file(spec_path, output_path)
