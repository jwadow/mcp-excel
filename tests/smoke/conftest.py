# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Pytest fixtures for smoke tests.

This module provides infrastructure for testing the MCP server through real JSON-RPC protocol.
The server is started as a separate process and communicates via STDIO, exactly like a real AI agent would.
"""

import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Generator

import pytest

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "src"))

from tests.fixtures.registry import get_fixture, FixtureMetadata


class MCPServerProcess:
    """Wrapper for MCP server process with JSON-RPC communication."""

    def __init__(self, process: subprocess.Popen):
        """Initialize MCP server process wrapper.
        
        Args:
            process: Running MCP server process
        """
        self.process = process
        self.request_id = 0
        self._initialized = False

    def send_request(self, method: str, params: dict[str, Any] | None = None, timeout: float = 30.0) -> dict[str, Any]:
        """Send JSON-RPC request to MCP server.
        
        Args:
            method: JSON-RPC method name (e.g., "tools/list", "tools/call")
            params: Method parameters (optional)
            timeout: Timeout in seconds (default: 30)
            
        Returns:
            Parsed JSON-RPC response
            
        Raises:
            RuntimeError: If server doesn't respond or returns error
            TimeoutError: If request times out
        """
        # Check if process is still running
        if self.process.poll() is not None:
            stderr = self.process.stderr.read() if self.process.stderr else ""
            raise RuntimeError(f"MCP server process terminated unexpectedly. Stderr: {stderr}")

        # Build JSON-RPC request
        self.request_id += 1
        request = {
            "jsonrpc": "2.0",
            "id": self.request_id,
            "method": method,
        }
        if params is not None:
            request["params"] = params

        # Send request (one line of JSON)
        request_json = json.dumps(request) + "\n"
        print(f"\n→ Sending request: {method}")
        print(f"  Request ID: {self.request_id}")
        if params:
            params_str = json.dumps(params, indent=2, ensure_ascii=False)
            if len(params_str) > 200:
                params_str = params_str[:200] + "..."
            print(f"  Params: {params_str}")
        
        try:
            self.process.stdin.write(request_json)
            self.process.stdin.flush()
        except BrokenPipeError:
            stderr = self.process.stderr.read() if self.process.stderr else ""
            raise RuntimeError(f"Failed to write to server stdin (broken pipe). Stderr: {stderr}")

        # Read response with timeout
        start_time = time.time()
        response_line = None
        
        while time.time() - start_time < timeout:
            # Check if process died
            if self.process.poll() is not None:
                stderr = self.process.stderr.read() if self.process.stderr else ""
                raise RuntimeError(f"MCP server died while waiting for response. Stderr: {stderr}")
            
            # Try to read line
            try:
                response_line = self.process.stdout.readline()
                if response_line:
                    break
            except Exception as e:
                raise RuntimeError(f"Error reading from server stdout: {e}")
            
            time.sleep(0.01)  # Small sleep to avoid busy-waiting

        if not response_line:
            raise TimeoutError(f"Server didn't respond within {timeout}s")

        # Parse response
        try:
            response = json.loads(response_line)
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Invalid JSON response from server: {response_line[:200]}. Error: {e}")

        print(f"← Received response for request {self.request_id}")
        print(f"  Response keys: {list(response.keys())}")
        
        # Check for JSON-RPC error
        if "error" in response:
            error = response["error"]
            print(f"  ❌ ERROR: {error}")
            raise RuntimeError(f"Server returned error: {json.dumps(error, indent=2)}")

        return response

    def initialize(self) -> dict[str, Any]:
        """Initialize MCP server (required before any other calls).
        
        Returns:
            Server capabilities and metadata
        """
        if self._initialized:
            return {"status": "already_initialized"}
        
        print("\n🔧 Initializing MCP server...")
        response = self.send_request(
            "initialize",
            {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "smoke-test-client",
                    "version": "1.0.0"
                }
            }
        )
        
        self._initialized = True
        print("✅ Server initialized successfully")
        return response

    def list_tools(self) -> list[dict[str, Any]]:
        """Get list of available tools from server.
        
        Returns:
            List of tool definitions
        """
        response = self.send_request("tools/list", {})
        return response.get("result", {}).get("tools", [])

    def call_tool(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        """Call a tool on the server.
        
        Args:
            name: Tool name
            arguments: Tool arguments
            
        Returns:
            Tool response (parsed from JSON)
        """
        response = self.send_request(
            "tools/call",
            {
                "name": name,
                "arguments": arguments
            }
        )
        
        # Extract result from MCP response
        result = response.get("result", {})
        
        # MCP wraps tool responses in "content" array with TextContent
        if "content" in result and isinstance(result["content"], list):
            # Get first content item (should be TextContent with JSON)
            if len(result["content"]) > 0:
                content_item = result["content"][0]
                if content_item.get("type") == "text":
                    # Parse JSON from text content
                    text = content_item.get("text", "{}")
                    try:
                        return json.loads(text)
                    except json.JSONDecodeError:
                        return {"raw_text": text}
        
        return result

    def shutdown(self) -> None:
        """Gracefully shutdown the server."""
        if self.process.poll() is None:  # Still running
            print("\n🛑 Shutting down MCP server...")
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
                print("✅ Server terminated gracefully")
            except subprocess.TimeoutExpired:
                print("⚠️  Server didn't terminate, killing...")
                self.process.kill()
                self.process.wait()
                print("✅ Server killed")


@pytest.fixture(scope="module")
def mcp_server() -> Generator[MCPServerProcess, None, None]:
    """Start real MCP server process and provide communication interface.
    
    This fixture:
    - Starts MCP server as subprocess
    - Initializes JSON-RPC communication
    - Provides helper methods for sending requests
    - Ensures proper cleanup on test completion
    
    Scope: module (one server for all tests in a file for performance)
    
    Yields:
        MCPServerProcess instance for communication with server
    """
    print("\n" + "="*80)
    print("🚀 Starting MCP Excel Server for smoke tests...")
    print("="*80)
    
    # Start server process
    process = subprocess.Popen(
        [sys.executable, "-m", "mcp_excel.main"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        bufsize=1,  # Line buffered - critical for Windows
        cwd=str(Path(__file__).parent.parent.parent)  # Project root
    )
    
    # Give server time to start
    time.sleep(1)
    
    # Check if server started successfully
    if process.poll() is not None:
        stderr = process.stderr.read()
        raise RuntimeError(f"MCP server failed to start. Exit code: {process.returncode}. Stderr:\n{stderr}")
    
    print(f"✅ Server process started (PID: {process.pid})")
    
    # Wrap in helper class
    server = MCPServerProcess(process)
    
    # Initialize server
    try:
        server.initialize()
    except Exception as e:
        # Cleanup on initialization failure
        server.shutdown()
        raise RuntimeError(f"Failed to initialize MCP server: {e}")
    
    # Yield to tests
    yield server
    
    # Cleanup
    print("\n" + "="*80)
    print("🧹 Cleaning up MCP server...")
    print("="*80)
    server.shutdown()


@pytest.fixture
def mcp_call_tool(mcp_server: MCPServerProcess):
    """Convenience fixture for calling tools.
    
    Usage:
        def test_something(mcp_call_tool):
            result = mcp_call_tool("inspect_file", {"file_path": "..."})
            assert result is not None
    """
    def call(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        return mcp_server.call_tool(name, arguments)
    return call


# ============================================================================
# FIXTURE METADATA (reuse from main conftest.py)
# ============================================================================

@pytest.fixture
def simple_fixture() -> FixtureMetadata:
    """Simple table with Cyrillic data."""
    return get_fixture("simple")


@pytest.fixture
def with_dates_fixture() -> FixtureMetadata:
    """Table with datetime columns."""
    return get_fixture("with_dates")


@pytest.fixture
def numeric_types_fixture() -> FixtureMetadata:
    """Table with different numeric types."""
    return get_fixture("numeric_types")


@pytest.fixture
def multi_sheet_fixture() -> FixtureMetadata:
    """File with 3 sheets (Products, Clients, Orders)."""
    return get_fixture("multi_sheet")


@pytest.fixture
def with_nulls_fixture() -> FixtureMetadata:
    """Table with null/empty values."""
    return get_fixture("with_nulls")


@pytest.fixture
def with_duplicates_fixture() -> FixtureMetadata:
    """Table with duplicate rows."""
    return get_fixture("with_duplicates")
