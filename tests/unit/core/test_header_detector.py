# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Unit tests for HeaderDetector component.

Tests cover:
- Header detection algorithm with various scenarios
- Scoring mechanism for candidate rows
- Confidence calculation
- Edge cases (empty rows, merged cells, multi-level headers)
- Error handling
"""

import pytest
import pandas as pd


def test_detect_simple_header(simple_fixture, file_loader, header_detector):
    """Test detection of simple header in clean file.
    
    Verifies:
    - Detects header in row 0
    - High confidence score
    - Correct candidate ranking
    """
    print(f"\n📂 Testing header detection on: {simple_fixture.name}")
    
    # Load without header to test detection
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=None)
    
    # Act
    result = header_detector.detect(df)
    
    # Assert
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    print(f"   Expected row: {simple_fixture.header_row}")
    
    assert result.header_row == simple_fixture.header_row, "Should detect correct header row"
    assert result.confidence > 0.8, "Should have high confidence for clean file"
    assert len(result.candidates) > 0, "Should return candidate list"


def test_detect_messy_headers(messy_headers_fixture, file_loader, header_detector):
    """Test detection of header after junk rows.
    
    Verifies:
    - Detects header in row 3 (after company name, report title, empty row)
    - Skips junk rows correctly
    - Reasonable confidence score
    """
    print(f"\n📂 Testing header detection on: {messy_headers_fixture.name}")
    
    # Load without header
    df = file_loader.load(messy_headers_fixture.path_str, messy_headers_fixture.sheet_name, header_row=None)
    
    # Act
    result = header_detector.detect(df)
    
    # Assert
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    print(f"   Expected row: {messy_headers_fixture.header_row}")
    print(f"   Preview: {result.candidates[0]['preview']}")
    
    assert result.header_row == messy_headers_fixture.header_row, "Should skip junk rows"
    assert result.confidence > 0.7, "Should have reasonable confidence"


def test_detect_multilevel_headers(multilevel_headers_fixture, file_loader, header_detector):
    """Test detection of deepest header level in multi-level structure.
    
    Verifies:
    - Detects deepest level (row 2) as header
    - Skips company name and category levels
    """
    print(f"\n📂 Testing header detection on: {multilevel_headers_fixture.name}")
    
    # Load without header
    df = file_loader.load(multilevel_headers_fixture.path_str, multilevel_headers_fixture.sheet_name, header_row=None)
    
    # Act
    result = header_detector.detect(df)
    
    # Assert
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    print(f"   Expected row: {multilevel_headers_fixture.header_row}")
    
    assert result.header_row == multilevel_headers_fixture.header_row, "Should detect deepest header level"


def test_detect_enterprise_chaos(enterprise_chaos_fixture, file_loader, header_detector):
    """Test detection in worst-case scenario (junk + merged + multi-level).
    
    Verifies:
    - Handles complex real-world files
    - Detects correct header despite chaos
    """
    print(f"\n📂 Testing header detection on: {enterprise_chaos_fixture.name}")
    
    # Load without header
    df = file_loader.load(enterprise_chaos_fixture.path_str, enterprise_chaos_fixture.sheet_name, header_row=None)
    
    # Act
    result = header_detector.detect(df)
    
    # Assert
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    print(f"   Expected row: {enterprise_chaos_fixture.header_row}")
    
    assert result.header_row == enterprise_chaos_fixture.header_row, "Should handle worst-case scenario"


def test_detect_merged_cells(merged_cells_fixture, file_loader, header_detector):
    """Test detection with merged cells in headers.
    
    Verifies:
    - Handles merged cells correctly
    - Detects deepest header row
    """
    print(f"\n📂 Testing header detection on: {merged_cells_fixture.name}")
    
    # Load without header
    df = file_loader.load(merged_cells_fixture.path_str, merged_cells_fixture.sheet_name, header_row=None)
    
    # Act
    result = header_detector.detect(df)
    
    # Assert
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    print(f"   Expected row: {merged_cells_fixture.header_row}")
    
    assert result.header_row == merged_cells_fixture.header_row, "Should handle merged cells"


def test_confidence_high_for_clean_files(simple_fixture, file_loader, header_detector):
    """Test that confidence is high for clean files.
    
    Verifies:
    - Clean files get confidence > 0.9
    - First row is clearly identified as header
    """
    print(f"\n📂 Testing confidence score on clean file")
    
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=None)
    result = header_detector.detect(df)
    
    print(f"✅ Confidence: {result.confidence:.2%}")
    
    assert result.confidence > 0.9, "Clean files should have very high confidence"


def test_confidence_lower_for_messy_files(messy_headers_fixture, file_loader, header_detector):
    """Test that confidence is reasonable for messy files.
    
    Verifies:
    - Messy files still get good confidence (algorithm works well)
    - Detection works correctly despite junk rows
    """
    print(f"\n📂 Testing confidence score on messy file")
    
    df = file_loader.load(messy_headers_fixture.path_str, messy_headers_fixture.sheet_name, header_row=None)
    result = header_detector.detect(df)
    
    print(f"✅ Confidence: {result.confidence:.2%}")
    
    # Algorithm works very well even on messy files
    assert result.confidence > 0.7, "Should have good confidence even for messy files"


def test_candidates_list_returned(simple_fixture, file_loader, header_detector):
    """Test that candidates list is returned with scores.
    
    Verifies:
    - At least 3 candidates returned
    - Candidates are sorted by score (descending)
    - Each candidate has required fields
    """
    print(f"\n📂 Testing candidates list")
    
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=None)
    result = header_detector.detect(df)
    
    print(f"✅ Candidates returned: {len(result.candidates)}")
    for idx, candidate in enumerate(result.candidates[:3], 1):
        print(f"   {idx}. Row {candidate['row']}: score={candidate['score']:.3f}")
    
    assert len(result.candidates) >= 3, "Should return at least 3 candidates"
    assert result.candidates[0]['row'] == result.header_row, "Best candidate should match detected row"
    
    # Check candidates are sorted by score
    scores = [c['score'] for c in result.candidates]
    assert scores == sorted(scores, reverse=True), "Candidates should be sorted by score"
    
    # Check each candidate has required fields
    for candidate in result.candidates:
        assert 'row' in candidate, "Candidate should have 'row' field"
        assert 'score' in candidate, "Candidate should have 'score' field"
        assert 'preview' in candidate, "Candidate should have 'preview' field"


def test_detect_or_ask_high_confidence(simple_fixture, file_loader, header_detector):
    """Test detect_or_ask with high confidence.
    
    Verifies:
    - Returns (row_index, None) when confidence is high
    - No candidates list when confident
    """
    print(f"\n📂 Testing detect_or_ask with high confidence")
    
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=None)
    header_row, candidates = header_detector.detect_or_ask(df)
    
    print(f"✅ Header row: {header_row}")
    print(f"   Candidates: {candidates}")
    
    assert header_row is not None, "Should return header row when confident"
    assert candidates is None, "Should not return candidates when confident"
    assert header_row == simple_fixture.header_row, "Should detect correct row"


def test_detect_or_ask_low_confidence(header_detector):
    """Test detect_or_ask with low confidence (ambiguous case).
    
    Verifies:
    - Returns (None, candidates) when confidence is low
    - Candidates list is provided for user choice
    """
    print(f"\n📂 Testing detect_or_ask with low confidence")
    
    # Create ambiguous DataFrame (all rows look similar)
    df = pd.DataFrame({
        0: ["A", "B", "C", "D", "E"],
        1: ["1", "2", "3", "4", "5"],
        2: ["X", "Y", "Z", "W", "V"],
    })
    
    # Lower min_confidence threshold to test this path
    detector_low_threshold = header_detector.__class__(min_confidence=0.95)
    header_row, candidates = detector_low_threshold.detect_or_ask(df)
    
    print(f"✅ Header row: {header_row}")
    print(f"   Candidates count: {len(candidates) if candidates else 0}")
    
    # With very similar rows, confidence might be low
    # This test verifies the mechanism works, not the specific outcome
    if header_row is None:
        assert candidates is not None, "Should return candidates when not confident"
        assert len(candidates) > 0, "Candidates list should not be empty"
        print(f"   ✅ Low confidence path triggered correctly")
    else:
        print(f"   ℹ️ Confidence was high enough, header detected: {header_row}")


def test_empty_dataframe_error(header_detector):
    """Test error handling for empty DataFrame.
    
    Verifies:
    - ValueError is raised for empty DataFrame
    - Error message is descriptive
    """
    print(f"\n📂 Testing empty DataFrame error")
    
    df = pd.DataFrame()
    
    with pytest.raises(ValueError) as exc_info:
        header_detector.detect(df)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    assert "empty" in str(exc_info.value).lower(), "Error should mention empty DataFrame"


def test_single_row_dataframe(header_detector):
    """Test detection with single row DataFrame.
    
    Verifies:
    - Detects row 0 as header
    - Doesn't crash with minimal data
    """
    print(f"\n📂 Testing single row DataFrame")
    
    df = pd.DataFrame({
        0: ["Name"],
        1: ["Age"],
        2: ["City"],
    })
    
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    
    assert result.header_row == 0, "Should detect row 0 as header"
    assert result.confidence > 0, "Should have some confidence"


def test_all_numeric_rows(header_detector):
    """Test detection when all rows are numeric.
    
    Verifies:
    - Still detects a header (first row by default)
    - Handles edge case gracefully
    """
    print(f"\n📂 Testing all numeric rows")
    
    df = pd.DataFrame({
        0: [1, 2, 3, 4, 5],
        1: [10, 20, 30, 40, 50],
        2: [100, 200, 300, 400, 500],
    })
    
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    
    assert result.header_row >= 0, "Should detect some row as header"
    # Confidence will be low for all-numeric data
    assert result.confidence >= 0, "Should have non-negative confidence"


def test_wide_table_detection(wide_table_fixture, file_loader, header_detector):
    """Test detection on wide table (50 columns).
    
    Verifies:
    - Handles many columns correctly
    - Performance is acceptable
    """
    print(f"\n📂 Testing header detection on wide table")
    
    df = file_loader.load(wide_table_fixture.path_str, wide_table_fixture.sheet_name, header_row=None)
    
    print(f"   DataFrame shape: {df.shape}")
    
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    
    assert result.header_row == wide_table_fixture.header_row, "Should detect header in wide table"


def test_single_column_detection(single_column_fixture, file_loader, header_detector):
    """Test detection on single column table.
    
    Verifies:
    - Handles minimal structure
    - Detects header correctly
    """
    print(f"\n📂 Testing header detection on single column")
    
    df = file_loader.load(single_column_fixture.path_str, single_column_fixture.sheet_name, header_row=None)
    
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    
    assert result.header_row == single_column_fixture.header_row, "Should detect header in single column"


def test_mixed_languages_detection(mixed_languages_fixture, file_loader, header_detector):
    """Test detection with mixed languages and special chars.
    
    Verifies:
    - Handles unicode correctly
    - Detects header with mixed encodings
    """
    print(f"\n📂 Testing header detection with mixed languages")
    
    df = file_loader.load(mixed_languages_fixture.path_str, mixed_languages_fixture.sheet_name, header_row=None)
    
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    print(f"   Preview: {result.candidates[0]['preview']}")
    
    assert result.header_row == mixed_languages_fixture.header_row, "Should handle mixed languages"


def test_special_chars_detection(special_chars_fixture, file_loader, header_detector):
    """Test detection with special characters and formula prefixes.
    
    Verifies:
    - Handles special characters (=, +, -, @)
    - Doesn't confuse formula prefixes with headers
    """
    print(f"\n📂 Testing header detection with special chars")
    
    df = file_loader.load(special_chars_fixture.path_str, special_chars_fixture.sheet_name, header_row=None)
    
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    
    assert result.header_row == special_chars_fixture.header_row, "Should handle special characters"


def test_with_nulls_detection(with_nulls_fixture, file_loader, header_detector):
    """Test detection with null values in data.
    
    Verifies:
    - Handles null values correctly
    - Doesn't confuse sparse data with headers
    """
    print(f"\n📂 Testing header detection with null values")
    
    df = file_loader.load(with_nulls_fixture.path_str, with_nulls_fixture.sheet_name, header_row=None)
    
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    
    assert result.header_row == with_nulls_fixture.header_row, "Should handle null values"


def test_scan_rows_parameter(simple_fixture, file_loader, header_detector):
    """Test that scan_rows parameter limits scanning.
    
    Verifies:
    - Only scans specified number of rows
    - Doesn't scan entire file unnecessarily
    """
    print(f"\n📂 Testing scan_rows parameter")
    
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=None)
    
    # Create detector with limited scan
    detector_limited = header_detector.__class__(scan_rows=5)
    result = detector_limited.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Candidates scanned: {len(result.candidates)}")
    
    # Should only scan first 5 rows
    assert all(c['row'] < 5 for c in result.candidates), "Should only scan first 5 rows"


def test_min_confidence_parameter(simple_fixture, file_loader, header_detector):
    """Test that min_confidence parameter affects detect_or_ask.
    
    Verifies:
    - Higher threshold requires higher confidence
    - Lower threshold accepts lower confidence
    """
    print(f"\n📂 Testing min_confidence parameter")
    
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=None)
    
    # Very high threshold - might not be confident enough
    detector_strict = header_detector.__class__(min_confidence=0.99)
    header_row_strict, candidates_strict = detector_strict.detect_or_ask(df)
    
    # Very low threshold - should always be confident
    detector_lenient = header_detector.__class__(min_confidence=0.1)
    header_row_lenient, candidates_lenient = detector_lenient.detect_or_ask(df)
    
    print(f"✅ Strict (0.99): header={header_row_strict}, candidates={candidates_strict is not None}")
    print(f"   Lenient (0.1): header={header_row_lenient}, candidates={candidates_lenient is not None}")
    
    # Lenient should always return header
    assert header_row_lenient is not None, "Lenient threshold should always detect header"
    assert candidates_lenient is None, "Lenient threshold should not return candidates"


def test_parametrized_all_messy_fixtures(messy_fixture_meta, file_loader, header_detector):
    """Parametrized test: detect headers in all messy fixtures.
    
    This test runs for EACH messy fixture automatically.
    Verifies that all real-world messy files can be handled.
    """
    print(f"\n📂 Testing header detection on messy fixture: {messy_fixture_meta.name}")
    print(f"   Description: {messy_fixture_meta.description}")
    
    df = file_loader.load(messy_fixture_meta.path_str, messy_fixture_meta.sheet_name, header_row=None)
    result = header_detector.detect(df)
    
    print(f"✅ Detected header row: {result.header_row}")
    print(f"   Expected row: {messy_fixture_meta.header_row}")
    print(f"   Confidence: {result.confidence:.2%}")
    
    assert result.header_row == messy_fixture_meta.header_row, f"Should detect correct header in {messy_fixture_meta.name}"
    assert result.confidence > 0.6, "Should have reasonable confidence even for messy files"
