# ABOUTME: Test Workbook Generator package for creating deterministic Excel test fixtures.
# ABOUTME: This package creates programmatically-generated workbooks to validate the IR extractor.

__version__ = "0.1.0"

# Public API for Story 1: Basic workbook creation infrastructure
# Public API for Story 6: Label candidates and named ranges
from test_workbook_generator.label_candidates_builder import (
    check_label_positioning_for_binding_detection,
    create_label_candidates_section,
    create_named_ranges,
    validate_label_candidates,
    validate_named_ranges,
    verify_named_ranges_in_formulas,
)

# Public API for Story 4: Overlap region shared dependencies
from test_workbook_generator.overlap_region_builder import (
    check_circular_references,
    create_overlap_region,
    validate_dependency_chain,
)

# Public API for Story 2: Tree A financial calculations
from test_workbook_generator.tree_a_builder import create_tree_a

# Public API for Story 3: Tree B statistical analysis
from test_workbook_generator.tree_b_builder import (
    create_lookup_table_data,
    create_tree_b,
)

# Public API for Story 5: Volatile functions implementation
from test_workbook_generator.volatile_functions_builder import (
    check_volatile_function_syntax,
    check_volatile_resolution,
    create_volatile_functions_section,
    validate_volatile_functions,
)
from test_workbook_generator.workbook_builder import (
    create_workbook,
    save_workbook,
    write_cell_formula,
    write_cell_value,
)

__all__ = [
    "create_workbook",
    "save_workbook",
    "write_cell_value",
    "write_cell_formula",
    "create_tree_a",
    "create_tree_b",
    "create_lookup_table_data",
    "create_overlap_region",
    "validate_dependency_chain",
    "check_circular_references",
    "create_volatile_functions_section",
    "validate_volatile_functions",
    "check_volatile_resolution",
    "check_volatile_function_syntax",
    "create_label_candidates_section",
    "create_named_ranges",
    "validate_label_candidates",
    "validate_named_ranges",
    "verify_named_ranges_in_formulas",
    "check_label_positioning_for_binding_detection",
    "__version__",
]
