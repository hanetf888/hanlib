"""
Pytest configuration and fixtures for graph_excel_handler tests.
"""

import os
import sys
import pytest
from unittest.mock import MagicMock, patch

# Add lib directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lib"))


@pytest.fixture
def mock_credentials():
    """Provide mock Azure AD credentials for testing."""
    return {
        "client_id": "test-client-id",
        "client_secret": "test-client-secret",
        "tenant_id": "test-tenant-id",
    }


@pytest.fixture
def mock_file_ids():
    """Provide mock SharePoint file identifiers."""
    return {
        "site_id": "contoso.sharepoint.com,abc123,def456",
        "drive_id": "b!abc123def456",
        "item_id": "01ABC123DEF456",
    }


@pytest.fixture
def mock_session_id():
    """Provide a mock workbook session ID."""
    return "session-id-12345"


@pytest.fixture
def mock_access_token():
    """Provide a mock access token."""
    return "mock-access-token-xyz"


@pytest.fixture
def handler(mock_credentials):
    """Create a GraphExcelHandler instance with mock credentials."""
    from graph_excel_handler import GraphExcelHandler

    return GraphExcelHandler(
        client_id=mock_credentials["client_id"],
        client_secret=mock_credentials["client_secret"],
        tenant_id=mock_credentials["tenant_id"],
    )


@pytest.fixture
def mock_msal_app(mock_access_token):
    """Mock the MSAL ConfidentialClientApplication."""
    with patch("graph_excel_handler.ConfidentialClientApplication") as mock_class:
        mock_app = MagicMock()
        mock_app.acquire_token_for_client.return_value = {
            "access_token": mock_access_token
        }
        mock_class.return_value = mock_app
        yield mock_class


@pytest.fixture
def mock_requests():
    """Mock the requests module for HTTP calls."""
    with patch("graph_excel_handler.requests") as mock_req:
        yield mock_req


# Integration test fixtures (require real credentials)


@pytest.fixture
def integration_credentials():
    """
    Load real credentials from environment variables for integration tests.

    Set these environment variables to run integration tests:
    - GRAPH_CLIENT_ID
    - GRAPH_CLIENT_SECRET
    - GRAPH_TENANT_ID
    - GRAPH_TEST_SITE_URL
    - GRAPH_TEST_FILE_PATH
    """
    client_id = os.environ.get("GRAPH_CLIENT_ID")
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET")
    tenant_id = os.environ.get("GRAPH_TENANT_ID")

    if not all([client_id, client_secret, tenant_id]):
        pytest.skip("Integration test credentials not configured")

    return {
        "client_id": client_id,
        "client_secret": client_secret,
        "tenant_id": tenant_id,
    }


@pytest.fixture
def integration_test_file():
    """
    Provide test file location for integration tests.

    Set these environment variables:
    - GRAPH_TEST_SITE_URL: Full SharePoint site URL
    - GRAPH_TEST_FILE_PATH: Path to test Excel file in site
    """
    site_url = os.environ.get("GRAPH_TEST_SITE_URL")
    file_path = os.environ.get("GRAPH_TEST_FILE_PATH")

    if not all([site_url, file_path]):
        pytest.skip("Integration test file location not configured")

    return {
        "site_url": site_url,
        "file_path": file_path,
    }
