"""
Tests for GraphExcelHandler - Microsoft Graph API Excel operations.
"""

import pytest
from unittest.mock import MagicMock, patch

from graph_excel_handler import GraphExcelHandler


class TestAuthentication:
    """Tests for authentication and token acquisition."""

    def test_init_stores_credentials(self, mock_credentials):
        """Test that credentials are stored on initialization."""
        handler = GraphExcelHandler(
            client_id=mock_credentials["client_id"],
            client_secret=mock_credentials["client_secret"],
            tenant_id=mock_credentials["tenant_id"],
        )

        assert handler.client_id == mock_credentials["client_id"]
        assert handler.client_secret == mock_credentials["client_secret"]
        assert handler.tenant_id == mock_credentials["tenant_id"]

    def test_get_access_token_success(self, handler, mock_msal_app, mock_access_token):
        """Test successful token acquisition."""
        token = handler._get_access_token()

        assert token == mock_access_token
        mock_msal_app.return_value.acquire_token_for_client.assert_called_once()

    def test_get_access_token_failure(self, handler):
        """Test token acquisition failure handling."""
        with patch("graph_excel_handler.ConfidentialClientApplication") as mock_class:
            mock_app = MagicMock()
            mock_app.acquire_token_for_client.return_value = {
                "error": "invalid_client",
                "error_description": "Invalid client secret",
            }
            mock_class.return_value = mock_app

            token = handler._get_access_token()

            assert token is None

    def test_get_headers_includes_token(self, handler, mock_msal_app, mock_access_token):
        """Test that headers include the authorization token."""
        headers = handler._get_headers()

        assert "Authorization" in headers
        assert headers["Authorization"] == f"Bearer {mock_access_token}"
        assert headers["Content-Type"] == "application/json"

    def test_get_headers_includes_session_id(self, handler, mock_msal_app, mock_session_id):
        """Test that headers include session ID when provided."""
        headers = handler._get_headers(session_id=mock_session_id)

        assert "workbook-session-id" in headers
        assert headers["workbook-session-id"] == mock_session_id


class TestSessionManagement:
    """Tests for workbook session creation and closure."""

    def test_create_session_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful session creation."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"id": mock_session_id}
        mock_response.raise_for_status = MagicMock()
        mock_requests.post.return_value = mock_response

        result = handler.create_session(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            persist_changes=True,
        )

        assert result["success"] is True
        assert result["data"]["session_id"] == mock_session_id

    def test_create_session_with_persist_false(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test session creation with persist_changes=False."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"id": mock_session_id}
        mock_response.raise_for_status = MagicMock()
        mock_requests.post.return_value = mock_response

        result = handler.create_session(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            persist_changes=False,
        )

        assert result["success"] is True
        # Verify the payload included persist_changes=False
        call_args = mock_requests.post.call_args
        assert call_args[1]["json"]["persistChanges"] is False

    def test_close_session_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful session closure."""
        mock_response = MagicMock()
        mock_response.status_code = 204
        mock_response.raise_for_status = MagicMock()
        mock_requests.post.return_value = mock_response

        result = handler.close_session(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
        )

        assert result["success"] is True
        assert result["data"]["closed"] is True


class TestReadOperations:
    """Tests for reading data from Excel files."""

    def test_read_range_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful range reading."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "values": [["A1", "B1"], ["A2", "B2"]],
            "formulas": [["", ""], ["", ""]],
            "address": "Sheet1!A1:B2",
            "rowCount": 2,
            "columnCount": 2,
        }
        mock_response.raise_for_status = MagicMock()
        mock_requests.get.return_value = mock_response

        result = handler.read_range(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
            worksheet_name="Sheet1",
            range_address="A1:B2",
        )

        assert result["success"] is True
        assert result["data"]["values"] == [["A1", "B1"], ["A2", "B2"]]
        assert result["data"]["row_count"] == 2
        assert result["data"]["column_count"] == 2

    def test_read_worksheet_names_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful worksheet names retrieval."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "value": [
                {"name": "Sheet1", "id": "1"},
                {"name": "Sheet2", "id": "2"},
                {"name": "Data", "id": "3"},
            ]
        }
        mock_response.raise_for_status = MagicMock()
        mock_requests.get.return_value = mock_response

        result = handler.read_worksheet_names(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
        )

        assert result["success"] is True
        assert result["data"]["worksheets"] == ["Sheet1", "Sheet2", "Data"]

    def test_get_used_range_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful used range retrieval."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "address": "Sheet1!A1:D10",
            "values": [["Header1", "Header2", "Header3", "Header4"]],
            "rowCount": 10,
            "columnCount": 4,
        }
        mock_response.raise_for_status = MagicMock()
        mock_requests.get.return_value = mock_response

        result = handler.get_used_range(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
            worksheet_name="Sheet1",
        )

        assert result["success"] is True
        assert result["data"]["address"] == "Sheet1!A1:D10"
        assert result["data"]["row_count"] == 10
        assert result["data"]["column_count"] == 4


class TestWriteOperations:
    """Tests for writing data to Excel files."""

    def test_write_range_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful range writing."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "address": "Sheet1!A1:B2",
            "rowCount": 2,
            "columnCount": 2,
        }
        mock_response.raise_for_status = MagicMock()
        mock_requests.patch.return_value = mock_response

        values = [["New1", "New2"], ["New3", "New4"]]
        result = handler.write_range(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
            worksheet_name="Sheet1",
            range_address="A1:B2",
            values=values,
        )

        assert result["success"] is True
        assert result["data"]["address"] == "Sheet1!A1:B2"
        # Verify the payload
        call_args = mock_requests.patch.call_args
        assert call_args[1]["json"]["values"] == values

    def test_write_cell_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful single cell writing."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "address": "Sheet1!A1",
            "rowCount": 1,
            "columnCount": 1,
        }
        mock_response.raise_for_status = MagicMock()
        mock_requests.patch.return_value = mock_response

        result = handler.write_cell(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
            worksheet_name="Sheet1",
            cell_address="A1",
            value="Test Value",
        )

        assert result["success"] is True
        # Verify the value is wrapped in 2D list
        call_args = mock_requests.patch.call_args
        assert call_args[1]["json"]["values"] == [["Test Value"]]

    def test_write_cell_with_formula(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test writing a formula to a cell."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "address": "Sheet1!C1",
            "rowCount": 1,
            "columnCount": 1,
        }
        mock_response.raise_for_status = MagicMock()
        mock_requests.patch.return_value = mock_response

        result = handler.write_cell(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
            worksheet_name="Sheet1",
            cell_address="C1",
            value="=SUM(A1:B1)",
        )

        assert result["success"] is True


class TestFormulaOperations:
    """Tests for formula recalculation."""

    def test_recalculate_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful formula recalculation."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.raise_for_status = MagicMock()
        mock_requests.post.return_value = mock_response

        result = handler.recalculate(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
        )

        assert result["success"] is True
        assert result["data"]["recalculated"] is True
        # Verify calculationType was set to Full
        call_args = mock_requests.post.call_args
        assert call_args[1]["json"]["calculationType"] == "Full"


class TestDataRefresh:
    """Tests for data connection refresh."""

    def test_refresh_all_data_connections_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful data connections refresh."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.raise_for_status = MagicMock()
        mock_requests.post.return_value = mock_response

        result = handler.refresh_all_data_connections(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
        )

        assert result["success"] is True
        assert result["data"]["refreshed"] is True


class TestOfficeScripts:
    """Tests for Office Scripts execution."""

    def test_run_office_script_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids
    ):
        """Test successful Office Script execution."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"value": "Script completed successfully"}
        mock_response.raise_for_status = MagicMock()
        mock_requests.post.return_value = mock_response

        result = handler.run_office_script(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            script_id="script-id-123",
        )

        assert result["success"] is True
        assert "result" in result["data"]
        # Verify beta endpoint is used
        call_args = mock_requests.post.call_args
        assert "beta" in call_args[0][0]


class TestFileIdResolution:
    """Tests for resolving SharePoint paths to IDs."""

    def test_get_file_ids_success(self, handler, mock_msal_app, mock_requests):
        """Test successful file ID resolution."""
        # Mock site info response
        site_response = MagicMock()
        site_response.status_code = 200
        site_response.json.return_value = {
            "id": "contoso.sharepoint.com,abc123,def456"
        }
        site_response.raise_for_status = MagicMock()

        # Mock drives response
        drives_response = MagicMock()
        drives_response.status_code = 200
        drives_response.json.return_value = {
            "value": [
                {"name": "Documents", "id": "drive-id-123"},
                {"name": "Site Assets", "id": "drive-id-456"},
            ]
        }
        drives_response.raise_for_status = MagicMock()

        # Mock item response
        item_response = MagicMock()
        item_response.status_code = 200
        item_response.json.return_value = {
            "id": "item-id-789",
            "name": "test.xlsx",
            "webUrl": "https://contoso.sharepoint.com/sites/test/Documents/test.xlsx",
        }
        item_response.raise_for_status = MagicMock()

        mock_requests.get.side_effect = [site_response, drives_response, item_response]

        result = handler.get_file_ids(
            site_url="https://contoso.sharepoint.com/sites/test",
            file_path="Documents/test.xlsx",
        )

        assert result["success"] is True
        assert result["data"]["site_id"] == "contoso.sharepoint.com,abc123,def456"
        assert result["data"]["drive_id"] == "drive-id-123"
        assert result["data"]["item_id"] == "item-id-789"
        assert result["data"]["file_name"] == "test.xlsx"


class TestErrorHandling:
    """Tests for error handling and response structure."""

    def test_error_response_structure(self, handler, mock_msal_app, mock_requests, mock_file_ids):
        """Test that error responses have the correct structure."""
        mock_response = MagicMock()
        mock_response.status_code = 404
        mock_response.raise_for_status.side_effect = Exception("Not found")
        mock_requests.post.return_value = mock_response

        result = handler.create_session(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
        )

        assert result["success"] is False
        assert "exception_type" in result
        assert "exception_message" in result
        assert "traceback" in result
        assert result["exception_type"] == "Exception"
        assert "Not found" in result["exception_message"]

    def test_rate_limit_handling(
        self, handler, mock_msal_app, mock_requests, mock_file_ids
    ):
        """Test that rate limiting is properly detected and reported."""
        mock_response = MagicMock()
        mock_response.status_code = 429
        mock_response.headers = {"Retry-After": "30"}
        mock_requests.post.return_value = mock_response

        result = handler.create_session(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
        )

        assert result["success"] is False
        assert "Rate limited" in result["exception_message"]
        assert "30" in result["exception_message"]

    def test_authentication_error_handling(self, handler, mock_requests, mock_file_ids):
        """Test handling of authentication errors."""
        with patch("graph_excel_handler.ConfidentialClientApplication") as mock_class:
            mock_app = MagicMock()
            mock_app.acquire_token_for_client.return_value = {
                "error": "invalid_grant",
                "error_description": "Invalid credentials",
            }
            mock_class.return_value = mock_app

            # This should fail because token is None
            mock_response = MagicMock()
            mock_response.status_code = 401
            mock_response.raise_for_status.side_effect = Exception("Unauthorized")
            mock_requests.post.return_value = mock_response

            result = handler.create_session(
                site_id=mock_file_ids["site_id"],
                drive_id=mock_file_ids["drive_id"],
                item_id=mock_file_ids["item_id"],
            )

            assert result["success"] is False


class TestSaveWorkbook:
    """Tests for workbook save operations."""

    def test_save_workbook_success(
        self, handler, mock_msal_app, mock_requests, mock_file_ids, mock_session_id
    ):
        """Test successful workbook save (close and reopen session)."""
        # Mock close session response
        close_response = MagicMock()
        close_response.status_code = 204
        close_response.raise_for_status = MagicMock()

        # Mock create session response
        create_response = MagicMock()
        create_response.status_code = 200
        create_response.json.return_value = {"id": "new-session-id"}
        create_response.raise_for_status = MagicMock()

        mock_requests.post.side_effect = [close_response, create_response]

        result = handler.save_workbook(
            site_id=mock_file_ids["site_id"],
            drive_id=mock_file_ids["drive_id"],
            item_id=mock_file_ids["item_id"],
            session_id=mock_session_id,
        )

        assert result["success"] is True
        assert result["data"]["saved"] is True
        assert result["data"]["new_session_id"] == "new-session-id"


# Integration Tests (require real credentials)


@pytest.mark.integration
class TestIntegration:
    """Integration tests that run against real SharePoint."""

    def test_full_workflow(self, integration_credentials, integration_test_file):
        """Test complete workflow: authenticate, create session, read/write, close."""
        handler = GraphExcelHandler(
            client_id=integration_credentials["client_id"],
            client_secret=integration_credentials["client_secret"],
            tenant_id=integration_credentials["tenant_id"],
        )

        # Get file IDs
        file_ids = handler.get_file_ids(
            site_url=integration_test_file["site_url"],
            file_path=integration_test_file["file_path"],
        )
        assert file_ids["success"], f"Failed to get file IDs: {file_ids}"

        site_id = file_ids["data"]["site_id"]
        drive_id = file_ids["data"]["drive_id"]
        item_id = file_ids["data"]["item_id"]

        # Create session
        session = handler.create_session(site_id, drive_id, item_id)
        assert session["success"], f"Failed to create session: {session}"
        session_id = session["data"]["session_id"]

        try:
            # Read worksheet names
            worksheets = handler.read_worksheet_names(
                site_id, drive_id, item_id, session_id
            )
            assert worksheets["success"], f"Failed to read worksheets: {worksheets}"
            assert len(worksheets["data"]["worksheets"]) > 0

            # Write to a cell
            write_result = handler.write_cell(
                site_id,
                drive_id,
                item_id,
                session_id,
                worksheet_name=worksheets["data"]["worksheets"][0],
                cell_address="A1",
                value="Integration Test",
            )
            assert write_result["success"], f"Failed to write cell: {write_result}"

            # Read back the value
            read_result = handler.read_range(
                site_id,
                drive_id,
                item_id,
                session_id,
                worksheet_name=worksheets["data"]["worksheets"][0],
                range_address="A1",
            )
            assert read_result["success"], f"Failed to read range: {read_result}"
            assert read_result["data"]["values"][0][0] == "Integration Test"

        finally:
            # Always close session
            close_result = handler.close_session(site_id, drive_id, item_id, session_id)
            assert close_result["success"], f"Failed to close session: {close_result}"
