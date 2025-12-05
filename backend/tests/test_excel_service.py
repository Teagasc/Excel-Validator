from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.append(str(ROOT))

from app.services.excel_service import (  # noqa: E402
    detect_column_type,
    identify_duplicates,
    validate_rows,
)


def test_detect_column_type_prefers_integer_when_majority_numeric():
    values = [1, "2", "03", None, ""]
    assert detect_column_type(values) == "integer"


def test_validate_rows_marks_invalid_cells():
    rows = [{"rowId": 0, "values": {"Age": "abc", "Name": "Sam"}}]
    expected_types = {"Age": "integer", "Name": "string"}
    errors, duplicate_groups = validate_rows(rows, expected_types)
    assert len(errors) == 1
    assert errors[0]["column"] == "Age"
    assert duplicate_groups == []


def test_identify_duplicates_groups_matching_rows():
    rows = [
        {"rowId": 0, "values": {"A": 1, "B": "x"}},
        {"rowId": 1, "values": {"A": 1, "B": "x"}},
        {"rowId": 2, "values": {"A": 2, "B": "y"}},
    ]
    groups = identify_duplicates(rows, ["A", "B"])
    assert groups == [[0, 1]]

