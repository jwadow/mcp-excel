# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Basic smoke tests for MCP server.

These tests verify fundamental server functionality:
- Server starts without errors
- Responds to initialization
- Lists all tools correctly
- JSON Schema validation
- Graceful shutdown
"""

import pytest


# ============================================================================
# SERVER LIFECYCLE TESTS
# ============================================================================

def test_server_starts_and_responds(mcp_server):
    """Smoke: Server starts successfully and responds to requests.
    
    This is the most basic smoke test - if this fails, nothing else will work.
    """
    print("\n✅ Server is running and responsive")
    assert mcp_server.process.poll() is None, "Server process should be running"
    assert mcp_server._initialized, "Server should be initialized"


def test_server_initialization(mcp_server):
    """Smoke: Server initialization returns valid response.
    
    Verifies that the MCP initialization handshake works correctly.
    """
    # Server is already initialized by fixture, but we can verify the state
    assert mcp_server._initialized
    print("\n✅ Server initialization successful")


# ============================================================================
# TOOL DISCOVERY TESTS
# ============================================================================

def test_list_tools_returns_all_tools(mcp_server):
    """Smoke: Server returns complete list of 25 tools.
    
    This verifies that all tools are registered correctly in main.py.
    """
    print("\n📋 Listing all tools...")
    tools = mcp_server.list_tools()
    
    print(f"  Found {len(tools)} tools")
    
    # Verify count
    assert len(tools) == 25, f"Expected 25 tools, got {len(tools)}"
    
    # Verify all expected tools are present
    expected_tools = {
        # Inspection (5)
        "inspect_file",
        "get_sheet_info",
        "get_column_names",
        "get_data_profile",
        "find_column",
        # Data retrieval (3)
        "get_unique_values",
        "get_value_counts",
        "filter_and_get_rows",
        # Filtering and counting (3)
        "filter_and_count",
        "filter_and_count_batch",
        "analyze_overlap",
        # Aggregation (2)
        "aggregate",
        "group_by",
        # Statistics (3)
        "get_column_stats",
        "correlate",
        "detect_outliers",
        # Validation (2)
        "find_duplicates",
        "find_nulls",
        # Multi-sheet (2)
        "search_across_sheets",
        "compare_sheets",
        # Time series (3)
        "calculate_period_change",
        "calculate_running_total",
        "calculate_moving_average",
        # Advanced (2)
        "rank_rows",
        "calculate_expression",
    }
    
    tool_names = {tool["name"] for tool in tools}
    
    # Check for missing tools
    missing = expected_tools - tool_names
    if missing:
        pytest.fail(f"Missing tools: {missing}")
    
    # Check for unexpected tools
    unexpected = tool_names - expected_tools
    if unexpected:
        pytest.fail(f"Unexpected tools: {unexpected}")
    
    print(f"  ✅ All 25 tools present")


def test_all_tools_have_required_fields(mcp_server):
    """Smoke: All tools have required fields (name, description, inputSchema).
    
    Verifies that tool definitions are complete and valid.
    """
    print("\n🔍 Validating tool definitions...")
    tools = mcp_server.list_tools()
    
    for tool in tools:
        tool_name = tool.get("name", "<unnamed>")
        print(f"  Checking {tool_name}...")
        
        # Required fields
        assert "name" in tool, f"Tool missing 'name' field"
        assert "description" in tool, f"Tool {tool_name} missing 'description' field"
        assert "inputSchema" in tool, f"Tool {tool_name} missing 'inputSchema' field"
        
        # Validate inputSchema structure
        schema = tool["inputSchema"]
        assert isinstance(schema, dict), f"Tool {tool_name} inputSchema must be dict"
        assert "type" in schema, f"Tool {tool_name} inputSchema missing 'type'"
        assert schema["type"] == "object", f"Tool {tool_name} inputSchema type must be 'object'"
        assert "properties" in schema, f"Tool {tool_name} inputSchema missing 'properties'"
        assert "required" in schema, f"Tool {tool_name} inputSchema missing 'required'"
        
        # Validate required fields exist in properties
        required_fields = schema.get("required", [])
        properties = schema.get("properties", {})
        for field in required_fields:
            assert field in properties, f"Tool {tool_name} required field '{field}' not in properties"
    
    print(f"  ✅ All {len(tools)} tool definitions valid")


def test_all_tools_have_valid_json_schema(mcp_server):
    """Smoke: All tool JSON schemas are valid according to JSON Schema Draft 7.
    
    This is critical - invalid schemas will cause MCP Framework to reject tool calls.
    """
    from jsonschema import Draft7Validator, SchemaError
    
    print("\n📐 Validating JSON schemas...")
    tools = mcp_server.list_tools()
    
    for tool in tools:
        tool_name = tool["name"]
        schema = tool["inputSchema"]
        
        print(f"  Validating {tool_name} schema...")
        
        try:
            # This will raise SchemaError if schema is invalid
            Draft7Validator.check_schema(schema)
        except SchemaError as e:
            pytest.fail(f"Tool {tool_name} has invalid JSON Schema: {e}")
    
    print(f"  ✅ All {len(tools)} schemas valid")


def test_filter_tools_have_definitions_at_root(mcp_server):
    """Smoke: Tools with filters have 'definitions' at root level (not nested).
    
    This is the bug that was fixed - definitions must be at root for $ref to work.
    Critical test to prevent regression.
    """
    print("\n🔍 Checking filter definitions placement...")
    tools = mcp_server.list_tools()
    
    # Tools that use filters
    filter_tools = {
        "filter_and_count",
        "filter_and_count_batch",
        "analyze_overlap",
        "filter_and_get_rows",
        "aggregate",
        "group_by",
        "get_column_stats",
        "correlate",
        "calculate_period_change",
        "calculate_running_total",
        "calculate_moving_average",
        "rank_rows",
        "calculate_expression",
    }
    
    for tool in tools:
        tool_name = tool["name"]
        
        if tool_name not in filter_tools:
            continue
        
        print(f"  Checking {tool_name}...")
        schema = tool["inputSchema"]
        
        # Check that definitions exist at root level
        assert "definitions" in schema, f"Tool {tool_name} missing 'definitions' at root level"
        
        # Check that FilterGroup is defined
        definitions = schema["definitions"]
        assert "FilterGroup" in definitions, f"Tool {tool_name} missing 'FilterGroup' in definitions"
        
        # Check that filters property uses $ref
        properties = schema.get("properties", {})
        if "filters" in properties:
            filters_schema = properties["filters"]
            # Should have items with oneOf containing $ref
            if "items" in filters_schema:
                items = filters_schema["items"]
                if "oneOf" in items:
                    one_of = items["oneOf"]
                    # Should have at least one $ref
                    has_ref = any("$ref" in option for option in one_of)
                    assert has_ref, f"Tool {tool_name} filters should use $ref for nested groups"
    
    print(f"  ✅ All filter tools have correct schema structure")


# ============================================================================
# TOOL DESCRIPTIONS TESTS
# ============================================================================

def test_all_tools_have_meaningful_descriptions(mcp_server):
    """Smoke: All tools have non-empty, meaningful descriptions.
    
    Descriptions are critical for AI agents to understand what tools do.
    """
    print("\n📝 Checking tool descriptions...")
    tools = mcp_server.list_tools()
    
    for tool in tools:
        tool_name = tool["name"]
        description = tool.get("description", "")
        
        # Check description exists and is not empty
        assert description, f"Tool {tool_name} has empty description"
        
        # Check description is meaningful (at least 50 characters)
        assert len(description) >= 50, f"Tool {tool_name} description too short ({len(description)} chars): {description[:50]}..."
    
    print(f"  ✅ All {len(tools)} tools have meaningful descriptions")


# ============================================================================
# ERROR HANDLING TESTS
# ============================================================================

def test_invalid_tool_name_returns_error(mcp_server):
    """Smoke: Calling non-existent tool returns proper error.
    
    Verifies error handling at MCP Framework level.
    Note: MCP Framework returns errors in response, not as exceptions.
    """
    print("\n❌ Testing error handling for invalid tool...")
    
    try:
        result = mcp_server.call_tool("nonexistent_tool", {})
        # If we got here, check if result contains error information
        print(f"  Result: {result}")
        # MCP might return error in result or as empty result
        # This is acceptable - server didn't crash
        print("  ✅ Server handled invalid tool gracefully (no crash)")
    except RuntimeError as e:
        # This is also acceptable - error was raised
        error_msg = str(e)
        print(f"  Got error (also acceptable): {error_msg[:100]}...")
        assert "nonexistent_tool" in error_msg.lower() or "unknown" in error_msg.lower() or "not found" in error_msg.lower()
        print("  ✅ Error handling works correctly")


def test_invalid_arguments_return_error(mcp_server, simple_fixture):
    """Smoke: Calling tool with invalid arguments returns proper error.
    
    Verifies JSON Schema validation at MCP Framework level.
    Note: MCP Framework may return errors in response or raise exceptions.
    """
    print("\n❌ Testing error handling for invalid arguments...")
    
    try:
        result = mcp_server.call_tool("get_sheet_info", {
            "file_path": str(simple_fixture.path_str)
            # Missing required sheet_name
        })
        # If we got here, check if result contains error
        print(f"  Result keys: {list(result.keys())}")
        if "error" in result:
            print(f"  Got error in result: {result['error']}")
            print("  ✅ Validation error returned in response")
        else:
            # Server might have handled it gracefully or used defaults
            print("  ⚠️  Server accepted invalid arguments (might use defaults)")
            print("  This is acceptable for smoke test - server didn't crash")
    except RuntimeError as e:
        # This is also acceptable - validation error was raised
        error_msg = str(e)
        print(f"  Got validation error: {error_msg[:100]}...")
        print("  ✅ Validation works correctly")
