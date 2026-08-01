# ABOUTME: Tests for the IR-inference / workbook-edit split in diff output —
# CT.IR_INFERENCE_TYPES membership, the summary roll-up, and the
# label_evidence_changed summary key (previously uncounted).

from xl_marinade.core.ir_diff import change_types as CT
from xl_marinade.core.ir_diff.model import Change
from xl_marinade.core.ir_diff.pipeline import _build_summary


def _c(ctype):
    return Change(type=ctype)


def test_inference_types_cover_exactly_the_three_semantic_categories():
    expected = frozenset(
        {
            CT.TABLE_CANDIDATE_ADDED,
            CT.TABLE_CANDIDATE_REMOVED,
            CT.TABLE_CANDIDATE_CHANGED,
            CT.BINDING_LABEL_EVIDENCE_CHANGED,
            CT.TIME_INDEX_CANDIDATE_ADDED,
            CT.TIME_INDEX_CANDIDATE_REMOVED,
            CT.TIME_INDEX_CANDIDATE_CHANGED,
            CT.BINDING_TIME_ANNOTATION_ADDED,
            CT.BINDING_TIME_ANNOTATION_REMOVED,
            CT.BINDING_TIME_ANNOTATION_CHANGED,
        }
    )
    assert expected == CT.IR_INFERENCE_TYPES


def test_workbook_edit_types_are_not_inference():
    for t in (
        CT.VALUE_CHANGED,
        CT.FORMULA_CHANGED,
        CT.BINDING_ADDED,
        CT.BINDING_LABEL_CHANGED,
        CT.BINDING_EDGE_ADDED,
        CT.NAME_ADDED,
        CT.SHEET_RENAMED,
    ):
        assert t not in CT.IR_INFERENCE_TYPES


def test_summary_rolls_up_inference_and_counts_label_evidence():
    changes = [
        _c(CT.VALUE_CHANGED),
        _c(CT.VALUE_CHANGED),
        _c(CT.FORMULA_CHANGED),
        _c(CT.TABLE_CANDIDATE_ADDED),
        _c(CT.TABLE_CANDIDATE_CHANGED),
        _c(CT.BINDING_LABEL_EVIDENCE_CHANGED),
        _c(CT.BINDING_TIME_ANNOTATION_ADDED),
    ]
    s = _build_summary(changes)
    assert s["total_changes"] == 7
    assert s["ir_inference_changes"] == 4
    assert s["tables_changed"] == 2
    assert s["label_evidence_changed"] == 1
    assert s["time_annotations_changed"] == 1
    assert s["cells_value_changed"] == 2
    assert s["cells_formula_changed"] == 1


def test_summary_inference_zero_when_no_semantic_changes():
    s = _build_summary([_c(CT.VALUE_CHANGED)])
    assert s["ir_inference_changes"] == 0
    assert s["label_evidence_changed"] == 0
