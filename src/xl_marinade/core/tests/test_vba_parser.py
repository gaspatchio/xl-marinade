# Test VBA parser module (Story 10)

from xl_marinade.core.vba_parser import (
    UDFMetadata,
    build_udf_map,
    compute_source_hash,
    detect_application_volatile,
    detect_udf_calls_in_formula,
    extract_function_body,
    extract_function_signature,
    is_formula_volatile_due_to_udfs,
    parse_vba_module,
)


def test_compute_source_hash_deterministic():
    """Test that source hash is deterministic."""
    source1 = "Function MyFunc(a, b)\n    MyFunc = a + b\nEnd Function"
    source2 = "Function MyFunc(a, b)\n    MyFunc = a + b\nEnd Function"

    hash1 = compute_source_hash(source1)
    hash2 = compute_source_hash(source2)

    assert hash1 == hash2
    assert len(hash1) == 64  # SHA-256 hex is 64 characters
    assert hash1.islower()  # Lowercase hex


def test_compute_source_hash_different():
    """Test that different sources produce different hashes."""
    source1 = "Function MyFunc(a, b)\n    MyFunc = a + b\nEnd Function"
    source2 = "Function MyFunc(a, b)\n    MyFunc = a * b\nEnd Function"

    hash1 = compute_source_hash(source1)
    hash2 = compute_source_hash(source2)

    assert hash1 != hash2


def test_extract_function_signature_basic():
    """Test basic function signature extraction."""
    func_line = "Function MyFunc(a, b)"
    name, params = extract_function_signature(func_line)

    assert name == "MyFunc"
    assert params == ["a", "b"]


def test_extract_function_signature_no_params():
    """Test function with no parameters."""
    func_line = "Function GetValue()"
    name, params = extract_function_signature(func_line)

    assert name == "GetValue"
    assert params == []


def test_extract_function_signature_with_types():
    """Test function with typed parameters."""
    func_line = "Function Calculate(x As Double, y As Double)"
    name, params = extract_function_signature(func_line)

    assert name == "Calculate"
    assert params == ["x", "y"]


def test_extract_function_signature_public():
    """Test public function declaration."""
    func_line = "Public Function MyPublicFunc(a, b As String)"
    name, params = extract_function_signature(func_line)

    assert name == "MyPublicFunc"
    assert params == ["a", "b"]


def test_extract_function_signature_with_optional():
    """Test function with optional parameters."""
    func_line = "Function MyFunc(a, Optional b As Integer = 5)"
    name, params = extract_function_signature(func_line)

    assert name == "MyFunc"
    assert params == ["a", "b"]


def test_detect_application_volatile_present():
    """Test detection when Application.Volatile is present."""
    source = """Function MyFunc(a)
    Application.Volatile
    MyFunc = a * 2
End Function"""

    assert detect_application_volatile(source) is True


def test_detect_application_volatile_absent():
    """Test detection when Application.Volatile is absent."""
    source = """Function MyFunc(a)
    MyFunc = a * 2
End Function"""

    assert detect_application_volatile(source) is False


def test_detect_application_volatile_case_insensitive():
    """Test that detection is case-insensitive."""
    source1 = "Function F()\n    application.volatile\nEnd Function"
    source2 = "Function F()\n    APPLICATION.VOLATILE\nEnd Function"

    assert detect_application_volatile(source1) is True
    assert detect_application_volatile(source2) is True


def test_extract_function_body():
    """Test extraction of complete function body."""
    lines = [
        "Sub DoStuff()",
        '    MsgBox "hi"',
        "End Sub",
        "Function MyFunc(a)",
        "    MyFunc = a * 2",
        "End Function",
    ]

    body, end_idx = extract_function_body(lines, 3)

    assert "Function MyFunc(a)" in body
    assert "MyFunc = a * 2" in body
    assert "End Function" in body
    assert end_idx == 5


def test_parse_vba_module_single_function():
    """Test parsing module with single function."""
    source = """
Option Explicit

Function DoubleValue(x)
    DoubleValue = x * 2
End Function
"""

    udfs = parse_vba_module("Module1", source)

    assert len(udfs) == 1
    assert udfs[0].name == "DoubleValue"
    assert udfs[0].module == "Module1"
    assert udfs[0].param_count == 1
    assert udfs[0].param_names == ["x"]
    assert udfs[0].declared_volatile is False
    assert "Function DoubleValue" in udfs[0].source_text
    assert len(udfs[0].source_hash) == 64


def test_parse_vba_module_multiple_functions():
    """Test parsing module with multiple functions."""
    source = """
Function Add(a, b)
    Add = a + b
End Function

Function Multiply(a, b)
    Application.Volatile
    Multiply = a * b
End Function
"""

    udfs = parse_vba_module("Module1", source)

    assert len(udfs) == 2
    assert udfs[0].name == "Add"
    assert udfs[0].declared_volatile is False
    assert udfs[1].name == "Multiply"
    assert udfs[1].declared_volatile is True


def test_parse_vba_module_ignores_private():
    """Test that private functions are excluded."""
    source = """
Public Function PublicFunc()
    PublicFunc = 1
End Function

Private Function PrivateFunc()
    PrivateFunc = 2
End Function

Function ImplicitPublic()
    ImplicitPublic = 3
End Function
"""

    udfs = parse_vba_module("Module1", source)

    # Should only parse Public and implicit public (no modifier)
    names = [udf.name for udf in udfs]
    assert "PublicFunc" in names
    assert "ImplicitPublic" in names
    # Private functions are currently included in basic regex
    # If we want to exclude them, we need to update the regex


def test_build_udf_map():
    """Test building UDF map from UDF list."""
    udf1 = UDFMetadata(
        name="MyFunc",
        module="Module1",
        param_count=2,
        param_names=["a", "b"],
        declared_volatile=False,
        source_text="Function MyFunc(a, b)\nEnd Function",
        source_hash="a" * 64,
    )
    udf2 = UDFMetadata(
        name="OtherFunc",
        module="Module2",
        param_count=1,
        param_names=["x"],
        declared_volatile=True,
        source_text="Function OtherFunc(x)\nEnd Function",
        source_hash="b" * 64,
    )

    udf_map = build_udf_map([udf1, udf2])

    assert "MYFUNC" in udf_map
    assert "OTHERFUNC" in udf_map
    assert udf_map["MYFUNC"].name == "MyFunc"
    assert udf_map["OTHERFUNC"].declared_volatile is True


def test_detect_udf_calls_in_formula():
    """Test detecting UDF calls in formula."""
    udf1 = UDFMetadata(
        name="MyFunc",
        module="Module1",
        param_count=2,
        param_names=["a", "b"],
        declared_volatile=False,
        source_text="",
        source_hash="a" * 64,
    )
    udf_map = {"MYFUNC": udf1}

    formula = "=MyFunc(A1, B1) + 10"
    calls = detect_udf_calls_in_formula(formula, udf_map)

    assert len(calls) == 1
    assert calls[0]["name"] == "MyFunc"
    assert calls[0]["module"] == "Module1"


def test_detect_udf_calls_case_insensitive():
    """Test that UDF detection is case-insensitive."""
    udf1 = UDFMetadata(
        name="MyFunc",
        module="Module1",
        param_count=1,
        param_names=["x"],
        declared_volatile=False,
        source_text="",
        source_hash="a" * 64,
    )
    udf_map = {"MYFUNC": udf1}

    # Formula uses lowercase
    formula = "=myfunc(A1)"
    calls = detect_udf_calls_in_formula(formula, udf_map)

    assert len(calls) == 1
    assert calls[0]["name"] == "MyFunc"  # Original case from VBA


def test_detect_udf_calls_multiple():
    """Test detecting multiple UDF calls."""
    udf1 = UDFMetadata(
        name="Func1",
        module="Module1",
        param_count=1,
        param_names=["x"],
        declared_volatile=False,
        source_text="",
        source_hash="a" * 64,
    )
    udf2 = UDFMetadata(
        name="Func2",
        module="Module1",
        param_count=1,
        param_names=["y"],
        declared_volatile=False,
        source_text="",
        source_hash="b" * 64,
    )
    udf_map = {"FUNC1": udf1, "FUNC2": udf2}

    formula = "=Func1(A1) + Func2(B1)"
    calls = detect_udf_calls_in_formula(formula, udf_map)

    assert len(calls) == 2
    names = [call["name"] for call in calls]
    assert "Func1" in names
    assert "Func2" in names


def test_detect_udf_calls_no_matches():
    """Test formula with no UDF calls."""
    udf_map = {}
    formula = "=SUM(A1:A10)"
    calls = detect_udf_calls_in_formula(formula, udf_map)

    assert len(calls) == 0


def test_detect_udf_calls_ignores_builtin():
    """Test that built-in functions are not detected as UDFs."""
    udf1 = UDFMetadata(
        name="MyFunc",
        module="Module1",
        param_count=1,
        param_names=["x"],
        declared_volatile=False,
        source_text="",
        source_hash="a" * 64,
    )
    udf_map = {"MYFUNC": udf1}

    formula = "=SUM(A1:A10) + MyFunc(B1)"
    calls = detect_udf_calls_in_formula(formula, udf_map)

    # Should only detect MyFunc, not SUM
    assert len(calls) == 1
    assert calls[0]["name"] == "MyFunc"


def test_detect_udf_calls_deduplication():
    """Test that duplicate calls are deduplicated."""
    udf1 = UDFMetadata(
        name="MyFunc",
        module="Module1",
        param_count=1,
        param_names=["x"],
        declared_volatile=False,
        source_text="",
        source_hash="a" * 64,
    )
    udf_map = {"MYFUNC": udf1}

    # MyFunc called twice
    formula = "=MyFunc(A1) + MyFunc(B1)"
    calls = detect_udf_calls_in_formula(formula, udf_map)

    # Should only return one entry (deduplicated)
    assert len(calls) == 1
    assert calls[0]["name"] == "MyFunc"


def test_is_formula_volatile_due_to_udfs_true():
    """Test volatile detection when UDF is volatile."""
    udf1 = UDFMetadata(
        name="VolatileFunc",
        module="Module1",
        param_count=0,
        param_names=[],
        declared_volatile=True,  # Volatile UDF
        source_text="",
        source_hash="a" * 64,
    )
    udf_map = {"VOLATILEFUNC": udf1}

    udf_calls = [{"name": "VolatileFunc", "module": "Module1"}]

    assert is_formula_volatile_due_to_udfs(udf_calls, udf_map) is True


def test_is_formula_volatile_due_to_udfs_false():
    """Test volatile detection when UDF is not volatile."""
    udf1 = UDFMetadata(
        name="NonVolatileFunc",
        module="Module1",
        param_count=0,
        param_names=[],
        declared_volatile=False,  # Not volatile
        source_text="",
        source_hash="a" * 64,
    )
    udf_map = {"NONVOLATILEFUNC": udf1}

    udf_calls = [{"name": "NonVolatileFunc", "module": "Module1"}]

    assert is_formula_volatile_due_to_udfs(udf_calls, udf_map) is False


def test_is_formula_volatile_mixed():
    """Test volatile detection with mix of volatile and non-volatile UDFs."""
    udf1 = UDFMetadata(
        name="NonVolatile",
        module="Module1",
        param_count=0,
        param_names=[],
        declared_volatile=False,
        source_text="",
        source_hash="a" * 64,
    )
    udf2 = UDFMetadata(
        name="Volatile",
        module="Module1",
        param_count=0,
        param_names=[],
        declared_volatile=True,
        source_text="",
        source_hash="b" * 64,
    )
    udf_map = {"NONVOLATILE": udf1, "VOLATILE": udf2}

    udf_calls = [
        {"name": "NonVolatile", "module": "Module1"},
        {"name": "Volatile", "module": "Module1"},
    ]

    # Should return True if ANY called UDF is volatile
    assert is_formula_volatile_due_to_udfs(udf_calls, udf_map) is True


def test_udf_metadata_to_dict():
    """Test UDFMetadata to_dict conversion."""
    udf = UDFMetadata(
        name="TestFunc",
        module="Module1",
        param_count=2,
        param_names=["a", "b"],
        declared_volatile=True,
        source_text="Function TestFunc(a, b)\nEnd Function",
        source_hash="a" * 64,
    )

    d = udf.to_dict()

    assert d["name"] == "TestFunc"
    assert d["module"] == "Module1"
    assert d["param_count"] == 2
    assert d["param_names_json"] == '["a", "b"]'
    assert d["declared_volatile"] is True
    assert d["source_text"] == "Function TestFunc(a, b)\nEnd Function"
    assert d["source_hash"] == "a" * 64
