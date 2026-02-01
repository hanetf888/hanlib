"""
Microsoft Graph API handler for Excel operations on SharePoint.
Supports co-authoring without file corruption by using Graph API sessions
instead of direct file manipulation with openpyxl.
"""

import logging
import traceback
from typing import Any, Optional
from urllib.parse import urlparse, quote

import requests
from msal import ConfidentialClientApplication

logger = logging.getLogger(__name__)


class GraphExcelHandler:
    """
    Microsoft Graph API handler for Excel operations on SharePoint.
    Supports co-authoring without file corruption.
    """

    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        """
        Initialize with Azure AD app credentials.

        Args:
            client_id: Azure AD application (client) ID
            client_secret: Azure AD application client secret
            tenant_id: Azure AD tenant ID
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self._token_cache = None
        self._msal_app = None

    def _get_msal_app(self) -> ConfidentialClientApplication:
        """Get or create the MSAL application instance."""
        if self._msal_app is None:
            self._msal_app = ConfidentialClientApplication(
                self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                client_credential=self.client_secret,
            )
        return self._msal_app

    def _get_access_token(self) -> Optional[str]:
        """
        Acquire an access token using client credentials flow.

        Returns:
            Access token string or None if acquisition fails
        """
        app = self._get_msal_app()
        result = app.acquire_token_for_client(scopes=self.GRAPH_SCOPE)

        if "access_token" in result:
            return result["access_token"]

        logger.error(f"Failed to acquire token: {result.get('error_description', 'Unknown error')}")
        return None

    def _get_headers(self, session_id: Optional[str] = None) -> dict:
        """
        Get HTTP headers for Graph API requests.

        Args:
            session_id: Optional workbook session ID for session-based operations

        Returns:
            Dictionary of HTTP headers
        """
        token = self._get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        if session_id:
            headers["workbook-session-id"] = session_id
        return headers

    def _build_workbook_url(self, site_id: str, drive_id: str, item_id: str) -> str:
        """Build the base URL for workbook operations."""
        return f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/workbook"

    def _make_error_response(self, exception: Exception) -> dict:
        """
        Create a standardized error response dictionary.

        Args:
            exception: The exception that occurred

        Returns:
            Error response dictionary with success=False
        """
        return {
            "success": False,
            "exception_type": type(exception).__name__,
            "exception_message": str(exception),
            "traceback": traceback.format_exc(),
        }

    def _make_success_response(self, data: Any) -> dict:
        """
        Create a standardized success response dictionary.

        Args:
            data: The data to include in the response

        Returns:
            Success response dictionary with success=True
        """
        return {"success": True, "data": data}

    def _handle_rate_limit(self, response: requests.Response) -> bool:
        """
        Check for rate limiting and log retry-after header if present.

        Args:
            response: The HTTP response object

        Returns:
            True if rate limited, False otherwise
        """
        if response.status_code == 429:
            retry_after = response.headers.get("Retry-After", "unknown")
            logger.warning(f"Rate limited. Retry after: {retry_after} seconds")
            return True
        return False

    # -------------------------------------------------------------------------
    # Session Management
    # -------------------------------------------------------------------------

    def create_session(
        self, site_id: str, drive_id: str, item_id: str, persist_changes: bool = True
    ) -> dict:
        """
        Create a workbook session for Excel operations.

        Sessions allow multiple operations to be grouped together and support
        co-authoring scenarios. Sessions timeout after ~5 minutes of inactivity.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            persist_changes: If True, changes are persisted to the workbook.
                           If False, creates a non-persistent session for read-only.

        Returns:
            Success dict with session_id in data, or error dict
        """
        try:
            url = f"{self._build_workbook_url(site_id, drive_id, item_id)}/createSession"
            headers = self._get_headers()
            payload = {"persistChanges": persist_changes}

            response = requests.post(url, headers=headers, json=payload)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            result = response.json()

            return self._make_success_response({"session_id": result.get("id")})

        except Exception as e:
            logger.error(f"Failed to create session: {e}")
            return self._make_error_response(e)

    def close_session(
        self, site_id: str, drive_id: str, item_id: str, session_id: str
    ) -> dict:
        """
        Close a workbook session.

        This commits any pending changes and releases the session.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Session ID to close

        Returns:
            Success dict or error dict
        """
        try:
            url = f"{self._build_workbook_url(site_id, drive_id, item_id)}/closeSession"
            headers = self._get_headers(session_id)

            response = requests.post(url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()

            return self._make_success_response({"closed": True})

        except Exception as e:
            logger.error(f"Failed to close session: {e}")
            return self._make_error_response(e)

    # -------------------------------------------------------------------------
    # Read Operations
    # -------------------------------------------------------------------------

    def read_range(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        session_id: str,
        worksheet_name: str,
        range_address: str,
    ) -> dict:
        """
        Read values from a range of cells.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID
            worksheet_name: Name of the worksheet
            range_address: Range address (e.g., "A1:C10")

        Returns:
            Success dict with values (2D list) in data, or error dict
        """
        try:
            worksheet_encoded = quote(worksheet_name, safe="")
            url = (
                f"{self._build_workbook_url(site_id, drive_id, item_id)}"
                f"/worksheets/{worksheet_encoded}/range(address='{range_address}')"
            )
            headers = self._get_headers(session_id)

            response = requests.get(url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            result = response.json()

            return self._make_success_response({
                "values": result.get("values", []),
                "formulas": result.get("formulas", []),
                "address": result.get("address"),
                "row_count": result.get("rowCount"),
                "column_count": result.get("columnCount"),
            })

        except Exception as e:
            logger.error(f"Failed to read range: {e}")
            return self._make_error_response(e)

    def read_worksheet_names(
        self, site_id: str, drive_id: str, item_id: str, session_id: str
    ) -> dict:
        """
        Get a list of all worksheet names in the workbook.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID

        Returns:
            Success dict with list of worksheet names in data, or error dict
        """
        try:
            url = f"{self._build_workbook_url(site_id, drive_id, item_id)}/worksheets"
            headers = self._get_headers(session_id)

            response = requests.get(url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            result = response.json()

            worksheets = [ws.get("name") for ws in result.get("value", [])]

            return self._make_success_response({"worksheets": worksheets})

        except Exception as e:
            logger.error(f"Failed to read worksheet names: {e}")
            return self._make_error_response(e)

    def get_used_range(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        session_id: str,
        worksheet_name: str,
    ) -> dict:
        """
        Get the used range of a worksheet (the smallest range encompassing all cells with data).

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID
            worksheet_name: Name of the worksheet

        Returns:
            Success dict with range info in data, or error dict
        """
        try:
            worksheet_encoded = quote(worksheet_name, safe="")
            url = (
                f"{self._build_workbook_url(site_id, drive_id, item_id)}"
                f"/worksheets/{worksheet_encoded}/usedRange"
            )
            headers = self._get_headers(session_id)

            response = requests.get(url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            result = response.json()

            return self._make_success_response({
                "address": result.get("address"),
                "values": result.get("values", []),
                "row_count": result.get("rowCount"),
                "column_count": result.get("columnCount"),
            })

        except Exception as e:
            logger.error(f"Failed to get used range: {e}")
            return self._make_error_response(e)

    # -------------------------------------------------------------------------
    # Write Operations
    # -------------------------------------------------------------------------

    def write_range(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        session_id: str,
        worksheet_name: str,
        range_address: str,
        values: list,
    ) -> dict:
        """
        Write values to a range of cells.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID
            worksheet_name: Name of the worksheet
            range_address: Range address (e.g., "A1:C3")
            values: 2D list of values to write (rows x columns)

        Returns:
            Success dict with updated range info, or error dict
        """
        try:
            worksheet_encoded = quote(worksheet_name, safe="")
            url = (
                f"{self._build_workbook_url(site_id, drive_id, item_id)}"
                f"/worksheets/{worksheet_encoded}/range(address='{range_address}')"
            )
            headers = self._get_headers(session_id)
            payload = {"values": values}

            response = requests.patch(url, headers=headers, json=payload)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            result = response.json()

            return self._make_success_response({
                "address": result.get("address"),
                "row_count": result.get("rowCount"),
                "column_count": result.get("columnCount"),
            })

        except Exception as e:
            logger.error(f"Failed to write range: {e}")
            return self._make_error_response(e)

    def write_cell(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        session_id: str,
        worksheet_name: str,
        cell_address: str,
        value: Any,
    ) -> dict:
        """
        Write a value to a single cell.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID
            worksheet_name: Name of the worksheet
            cell_address: Cell address (e.g., "A1")
            value: Value to write (string, number, boolean, or formula starting with =)

        Returns:
            Success dict with updated cell info, or error dict
        """
        # Use write_range with a single value wrapped in 2D list
        return self.write_range(
            site_id=site_id,
            drive_id=drive_id,
            item_id=item_id,
            session_id=session_id,
            worksheet_name=worksheet_name,
            range_address=cell_address,
            values=[[value]],
        )

    # -------------------------------------------------------------------------
    # Formula Operations
    # -------------------------------------------------------------------------

    def recalculate(
        self, site_id: str, drive_id: str, item_id: str, session_id: str
    ) -> dict:
        """
        Trigger recalculation of all formulas in the workbook.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID

        Returns:
            Success dict or error dict
        """
        try:
            url = f"{self._build_workbook_url(site_id, drive_id, item_id)}/application/calculate"
            headers = self._get_headers(session_id)
            payload = {"calculationType": "Full"}

            response = requests.post(url, headers=headers, json=payload)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()

            return self._make_success_response({"recalculated": True})

        except Exception as e:
            logger.error(f"Failed to recalculate: {e}")
            return self._make_error_response(e)

    # -------------------------------------------------------------------------
    # Data Refresh
    # -------------------------------------------------------------------------

    def refresh_all_data_connections(
        self, site_id: str, drive_id: str, item_id: str, session_id: str
    ) -> dict:
        """
        Refresh all external data connections in the workbook.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID

        Returns:
            Success dict or error dict
        """
        try:
            url = f"{self._build_workbook_url(site_id, drive_id, item_id)}/refreshAllDataConnections"
            headers = self._get_headers(session_id)

            response = requests.post(url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()

            return self._make_success_response({"refreshed": True})

        except Exception as e:
            logger.error(f"Failed to refresh data connections: {e}")
            return self._make_error_response(e)

    # -------------------------------------------------------------------------
    # Office Scripts
    # -------------------------------------------------------------------------

    def run_office_script(
        self, site_id: str, drive_id: str, item_id: str, script_id: str
    ) -> dict:
        """
        Run an Office Script on the workbook.

        Office Scripts are TypeScript-based automation scripts that replace VBA macros
        for cloud-based Excel files. Scripts must be pre-created in SharePoint/OneDrive.

        Note: This uses the Graph API beta endpoint as Office Scripts API is in preview.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            script_id: ID of the Office Script to run

        Returns:
            Success dict with script result, or error dict
        """
        try:
            # Office Scripts use beta endpoint
            url = (
                f"https://graph.microsoft.com/beta/sites/{site_id}/drives/{drive_id}"
                f"/items/{item_id}/workbook/scripts/{script_id}/run"
            )
            headers = self._get_headers()

            response = requests.post(url, headers=headers, json={})

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            result = response.json()

            return self._make_success_response({"result": result})

        except Exception as e:
            logger.error(f"Failed to run Office Script: {e}")
            return self._make_error_response(e)

    # -------------------------------------------------------------------------
    # Save Operations
    # -------------------------------------------------------------------------

    def save_workbook(
        self, site_id: str, drive_id: str, item_id: str, session_id: str
    ) -> dict:
        """
        Explicitly save the workbook.

        Note: Changes are automatically saved when closing a persistent session.
        This method can be used for intermediate saves during long operations.

        Args:
            site_id: SharePoint site ID
            drive_id: Drive ID containing the file
            item_id: Item ID of the Excel file
            session_id: Active session ID

        Returns:
            Success dict or error dict
        """
        try:
            # The Graph API doesn't have an explicit save endpoint for workbooks
            # in persistent sessions. Changes are saved on session close.
            # We can force a save by doing a refresh session or by closing and reopening.
            # For now, we'll close and reopen the session to force a save.

            close_result = self.close_session(site_id, drive_id, item_id, session_id)
            if not close_result.get("success"):
                return close_result

            # Reopen session
            new_session_result = self.create_session(
                site_id, drive_id, item_id, persist_changes=True
            )
            if not new_session_result.get("success"):
                return new_session_result

            return self._make_success_response({
                "saved": True,
                "new_session_id": new_session_result["data"]["session_id"],
            })

        except Exception as e:
            logger.error(f"Failed to save workbook: {e}")
            return self._make_error_response(e)

    # -------------------------------------------------------------------------
    # Helper: Find File by Path
    # -------------------------------------------------------------------------

    def get_file_ids(self, site_url: str, file_path: str) -> dict:
        """
        Get the site_id, drive_id, and item_id for a file given its SharePoint path.

        Args:
            site_url: SharePoint site URL (e.g., "https://contoso.sharepoint.com/sites/MySite")
            file_path: Path to the file within the site's document library
                      (e.g., "Documents/Reports/sales.xlsx")

        Returns:
            Success dict with site_id, drive_id, and item_id, or error dict
        """
        try:
            # Parse the site URL
            parsed = urlparse(site_url)
            hostname = parsed.netloc
            site_path = parsed.path

            # Get site ID
            site_url_encoded = f"{hostname}:{site_path}"
            site_api_url = f"{self.GRAPH_BASE_URL}/sites/{site_url_encoded}"
            headers = self._get_headers()

            response = requests.get(site_api_url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            site_info = response.json()
            site_id = site_info.get("id")

            # Parse file path to get drive name and item path
            path_parts = file_path.strip("/").split("/", 1)
            if len(path_parts) < 2:
                drive_name = path_parts[0]
                item_path = ""
            else:
                drive_name = path_parts[0]
                item_path = path_parts[1]

            # Get drive ID
            drives_url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives"
            response = requests.get(drives_url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            drives = response.json().get("value", [])

            drive_id = None
            for drive in drives:
                if drive.get("name", "").lower() == drive_name.lower():
                    drive_id = drive.get("id")
                    break

            if not drive_id:
                # Try with default "Documents" library if not found
                for drive in drives:
                    if drive.get("name", "").lower() == "documents":
                        drive_id = drive.get("id")
                        item_path = file_path.strip("/")
                        break

            if not drive_id:
                return self._make_error_response(
                    ValueError(f"Drive not found for path: {file_path}")
                )

            # Get item ID
            if item_path:
                item_url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{item_path}"
            else:
                item_url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root"

            response = requests.get(item_url, headers=headers)

            if self._handle_rate_limit(response):
                return self._make_error_response(
                    Exception(f"Rate limited. Retry-After: {response.headers.get('Retry-After')}")
                )

            response.raise_for_status()
            item_info = response.json()
            item_id = item_info.get("id")

            return self._make_success_response({
                "site_id": site_id,
                "drive_id": drive_id,
                "item_id": item_id,
                "file_name": item_info.get("name"),
                "web_url": item_info.get("webUrl"),
            })

        except Exception as e:
            logger.error(f"Failed to get file IDs: {e}")
            return self._make_error_response(e)
