"""
sheets_service.py — Google Sheets operations with Calibri font, size 13, center alignment
BRANCH: DADAR - Sales Only
Handles: Not Sent, actual data
Formatting: Dark black text (#000000), All borders on data cells (first 12 columns only)
UPDATED: Fixed write_batch() to properly clear "Not Sent" status when data is written
UPDATED: Supports writing "Not Sent" status for late Sales submissions
UPDATED: Added retry logic with exponential backoff for connection failures (503 errors)
UPDATED: Added timeout to prevent hanging
UPDATED: Removed HR references (Sales only branch)
UPDATED: Pre-populates employees with current date when creating new sheet
UPDATED: Borders applied ONLY to first 12 columns (A to L) - the actual data table
UPDATED: Added extensive debug logging
"""

import logging
import json
import os
import re
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import gspread
from google.oauth2 import service_account

from config import (
    DATE_IN_SUBJECT_FORMAT,
    SALES_COLUMN_MAPPING,
    SALES_EMPLOYEES,
    SALES_HEADERS,
    SALES_SPREADSHEET_ID,
    SHEET_NAME_FORMAT,
)
from error_handler import with_retry

logger = logging.getLogger(__name__)
_SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Maximum columns for data table (A to L = 12 columns)
MAX_DATA_COLUMNS = 12


def _get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS", "")
    if creds_json:
        logger.info("Loading from GOOGLE_CREDENTIALS env var")
        return service_account.Credentials.from_service_account_info(json.loads(creds_json), scopes=_SCOPES)

    if os.path.exists("credentials.json"):
        logger.info("Loading from credentials.json file")
        return service_account.Credentials.from_service_account_file("credentials.json", scopes=_SCOPES)

    raise Exception("No valid credentials found")


def _get_gspread_client() -> gspread.Client:
    client = gspread.authorize(_get_credentials())
    # Set timeout to prevent hanging on slow connections
    client.timeout = 30
    return client


def _connect_with_retry(max_retries: int = 5, initial_delay: int = 2) -> Tuple[gspread.Client, object]:
    """
    Attempt to connect to Google Sheets with retry logic for 503 errors.
    
    Args:
        max_retries: Maximum number of retry attempts
        initial_delay: Initial delay in seconds (doubles each retry)
    
    Returns:
        Tuple of (client, sales_ss)
    """
    last_error = None
    
    for attempt in range(max_retries):
        try:
            logger.info(f"🔄 Connecting to Google Sheets (attempt {attempt + 1}/{max_retries})...")
            
            client = _get_gspread_client()
            
            # Try to open the Sales spreadsheet
            sales_ss = client.open_by_key(SALES_SPREADSHEET_ID)
            
            # Verify connection by fetching spreadsheet title
            sales_title = sales_ss.title
            
            logger.info(f"✅ Successfully connected to Sales spreadsheet: '{sales_title}'")
            
            return client, sales_ss
            
        except Exception as e:
            last_error = e
            error_msg = str(e)
            
            # Check if this is a 503 error or connection issue
            is_503 = "503" in error_msg or "unavailable" in error_msg.lower()
            is_connection = "connection" in error_msg.lower() or "timeout" in error_msg.lower()
            
            if is_503 or is_connection:
                if attempt < max_retries - 1:
                    delay = initial_delay * (2 ** attempt)  # Exponential backoff
                    logger.warning(f"⚠️ Google Sheets {'503 (unavailable)' if is_503 else 'connection'} error (attempt {attempt + 1})")
                    logger.info(f"⏳ Retrying in {delay} seconds...")
                    time.sleep(delay)
                    continue
                else:
                    logger.error(f"❌ Failed to connect after {max_retries} attempts")
                    break
            else:
                # Non-retryable error
                logger.error(f"❌ Non-retryable error: {e}")
                break
    
    # If we get here, all retries failed
    raise Exception(f"Failed to connect to Google Sheets after {max_retries} attempts. Last error: {last_error}")


class SheetsService:
    def __init__(self):
        # Use retry logic for connection
        client, sales_ss = _connect_with_retry()
        
        self._sales_ss = sales_ss
        self._client = client
        
        logger.info("✅ Connected to Sales spreadsheet")
        self._emp_cache: Dict[Tuple[str, str], Dict[str, int]] = {}
        self._ws_cache: Dict[Tuple[str, str], gspread.Worksheet] = {}
        self._worksheet_data_cache: Dict[Tuple[str, str], List[List[str]]] = {}
        self._cache_timestamp: Dict[Tuple[str, str], datetime] = {}
        self._cache_ttl_seconds = 60
        self._date_marked_not_sent: set = set()

    def _spreadsheet(self, department: str):
        return self._sales_ss

    @staticmethod
    def sheet_name(date_str: str) -> str:
        return datetime.strptime(date_str, DATE_IN_SUBJECT_FORMAT).strftime(SHEET_NAME_FORMAT)
    
    def get_column_mapping(self, department: str) -> Dict:
        return SALES_COLUMN_MAPPING

    def _get_cached_worksheet_data(self, department: str, date_str: str) -> List[List[str]]:
        key = (department, date_str)
        now = datetime.now()
        
        if key in self._worksheet_data_cache and key in self._cache_timestamp:
            if (now - self._cache_timestamp[key]).seconds < self._cache_ttl_seconds:
                return self._worksheet_data_cache[key]
        
        ws = self._get_worksheet(department, date_str)
        data = ws.get_all_values()
        self._worksheet_data_cache[key] = data
        self._cache_timestamp[key] = now
        return data
    
    def _invalidate_cache(self, department: str, date_str: str) -> None:
        key = (department, date_str)
        if key in self._worksheet_data_cache:
            del self._worksheet_data_cache[key]
        if key in self._cache_timestamp:
            del self._cache_timestamp[key]

    def _apply_table_formatting(self, ws: gspread.Worksheet) -> None:
        """
        Apply formatting (font, alignment, borders) to data table cells.
        Borders applied ONLY to first 12 columns (A to L).
        """
        try:
            all_values = ws.get_all_values()
            if not all_values:
                logger.info("🔍 No values found in sheet for formatting")
                return
            
            max_row = len(all_values)
            # Only use first 12 columns for borders (A to L)
            max_col = min(MAX_DATA_COLUMNS, max(len(row) for row in all_values) if all_values else len(SALES_HEADERS))
            
            logger.info(f"🔍 Formatting: max_row={max_row}, max_col={max_col}")
            
            # Ensure we have at least headers
            if max_row < 1:
                return
            
            # Add some buffer for future data (format up to 500 rows)
            max_row = min(max_row + 50, 500)
            
            # Create range for entire table (A1 to L{max_row})
            range_str = f"A1:{gspread.utils.rowcol_to_a1(max_row, max_col)}"
            
            sheet_id = ws.id
            
            requests = [{
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": max_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": max_col
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontFamily": "Calibri",
                                "fontSize": 13,
                                "foregroundColor": {
                                    "red": 0.0,
                                    "green": 0.0,
                                    "blue": 0.0
                                }
                            },
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "borders": {
                                "top": {"style": "SOLID"},
                                "bottom": {"style": "SOLID"},
                                "left": {"style": "SOLID"},
                                "right": {"style": "SOLID"}
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.borders"
                }
            }]
            
            ws.spreadsheet.batch_update({"requests": requests})
            logger.info(f"✅ Applied table formatting to '{ws.title}' - {max_row} rows x {max_col} columns (A to L)")
            
        except Exception as e:
            logger.warning(f"Table formatting failed for {ws.title}: {e}")
            import traceback
            traceback.print_exc()

    def _apply_formatting_to_range(self, ws: gspread.Worksheet, range_str: str = None) -> None:
        """
        Apply formatting to a specific range. If no range is provided, format the entire table.
        Borders applied ONLY to first 12 columns.
        """
        try:
            if range_str is None:
                self._apply_table_formatting(ws)
                return
            
            # Parse the range
            sheet_id = ws.id
            
            # Get all values to determine max rows
            all_values = ws.get_all_values()
            max_row = len(all_values)
            # Only use first 12 columns
            max_col = min(MAX_DATA_COLUMNS, max(len(row) for row in all_values) if all_values else len(SALES_HEADERS))
            
            # Parse range to get start and end
            start_row_num = 1
            end_row_num = max_row
            start_col_num = 1
            end_col_num = max_col
            
            if ":" in range_str:
                start_cell, end_cell = range_str.split(":")
                
                # Parse start row
                start_row_match = re.search(r'(\d+)$', start_cell)
                if start_row_match:
                    start_row_num = int(start_row_match.group(1))
                
                # Parse end row
                end_row_match = re.search(r'(\d+)$', end_cell)
                if end_row_match:
                    end_row_num = int(end_row_match.group(1))
                
                # Parse start column
                start_col_letters = ''.join(filter(str.isalpha, start_cell))
                start_col_num = self._col_letter_to_index(start_col_letters) if start_col_letters else 1
                
                # Parse end column
                end_col_letters = ''.join(filter(str.isalpha, end_cell))
                end_col_num = self._col_letter_to_index(end_col_letters) if end_col_letters else max_col
            else:
                # Single cell
                row_match = re.search(r'(\d+)$', range_str)
                if row_match:
                    start_row_num = int(row_match.group(1))
                    end_row_num = start_row_num
                
                col_letters = ''.join(filter(str.isalpha, range_str))
                if col_letters:
                    start_col_num = self._col_letter_to_index(col_letters)
                    end_col_num = start_col_num
            
            # Ensure we don't exceed 12 columns
            end_col_num = min(end_col_num, MAX_DATA_COLUMNS)
            if start_col_num > MAX_DATA_COLUMNS:
                # If formatting column beyond 12, skip (no borders needed)
                return
            
            # Ensure we don't go beyond data limits
            end_row_num = min(end_row_num, max_row + 10)
            
            requests = [{
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row_num - 1,
                        "endRowIndex": end_row_num,
                        "startColumnIndex": start_col_num - 1,
                        "endColumnIndex": end_col_num
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontFamily": "Calibri",
                                "fontSize": 13,
                                "foregroundColor": {
                                    "red": 0.0,
                                    "green": 0.0,
                                    "blue": 0.0
                                }
                            },
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "borders": {
                                "top": {"style": "SOLID"},
                                "bottom": {"style": "SOLID"},
                                "left": {"style": "SOLID"},
                                "right": {"style": "SOLID"}
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.borders"
                }
            }]
            
            ws.spreadsheet.batch_update({"requests": requests})
            logger.info(f"✅ Applied formatting to '{ws.title}' - Range: {range_str} (within A-L)")
            
        except Exception as e:
            logger.warning(f"Formatting failed for {ws.title}: {e}")

    @staticmethod
    def _col_letter_to_index(col_letter: str) -> int:
        """Convert column letter to index (1-based)"""
        index = 0
        for char in col_letter:
            index = index * 26 + (ord(char.upper()) - ord('A') + 1)
        return index

    def ensure_status_column(self, department: str, date_str: str) -> None:
        if department != "Sales":
            return
        
        try:
            ws = self._get_worksheet(department, date_str)
            headers = ws.row_values(1)
            logger.info(f"🔍 Current headers: {headers}")
            
            if "Report Status" not in headers:
                last_col = len(headers) + 1
                # Only add if within 12 columns
                if last_col <= MAX_DATA_COLUMNS:
                    ws.update_cell(1, last_col, "Report Status")
                    logger.info(f"✅ Added 'Report Status' column at position {last_col}")
                    # Format the entire table after adding column
                    self._apply_table_formatting(ws)
                else:
                    logger.warning(f"⚠️ Cannot add 'Report Status' column - would exceed {MAX_DATA_COLUMNS} columns")
            else:
                logger.info("✅ 'Report Status' column already exists")
                # Ensure table formatting is applied (first 12 columns only)
                self._apply_table_formatting(ws)
        except Exception as e:
            logger.error(f"❌ Could not ensure status column: {e}")
            import traceback
            traceback.print_exc()

    def _get_worksheet(self, department: str, date_str: str) -> gspread.Worksheet:
        key = (department, date_str)
        if key in self._ws_cache:
            logger.info(f"🔍 Using cached worksheet for {department}/{date_str}")
            return self._ws_cache[key]

        ss = self._spreadsheet(department)
        name = self.sheet_name(date_str)
        logger.info(f"🔍 Looking for worksheet: {name}")
        
        try:
            ws = ss.worksheet(name)
            logger.info(f"✅ Found existing worksheet: {name}")
        except gspread.WorksheetNotFound:
            logger.info(f"🆕 Worksheet '{name}' not found. Creating new one...")
            ws = self._create_worksheet(ss, name, department)

        self._ws_cache[key] = ws
        return ws

    def _create_worksheet(self, ss: gspread.Spreadsheet, name: str, department: str) -> gspread.Worksheet:
        """Create a new worksheet with all employees pre-populated with current date"""
        employees = SALES_EMPLOYEES
        headers = SALES_HEADERS
        date_str = datetime.now().strftime("%d-%m-%Y")
        
        logger.info(f"🆕 Creating new worksheet: {name}")
        logger.info(f"👥 Employees: {employees}")
        logger.info(f"📋 Headers: {headers}")
        
        ws = ss.add_worksheet(title=name, rows=str(len(employees) * 35 + 10), cols="20")
        
        ws.update("A1", [headers])
        logger.info("✅ Headers written")
        
        # Add employees WITH the current date
        rows = []
        for emp in employees:
            rows.append([date_str, emp])
        ws.update(f"A2:B{len(employees) + 1}", rows)
        logger.info(f"✅ Added {len(rows)} employees with date {date_str}")
        
        # Apply table formatting (first 12 columns only)
        self._apply_table_formatting(ws)
        
        logger.info(f"✅ Created sheet '{name}' for {department} with {len(employees)} employees (borders: A-L)")
        return ws

    def mark_all_as_not_sent(self, department: str, date_str: str) -> None:
        """Mark employees as 'Not Sent' ONLY if they have NO data in other columns"""
        if department != "Sales":
            return
        
        date_key = f"{department}_{date_str}"
        if date_key in self._date_marked_not_sent:
            logger.info(f"⏭️ Already marked 'Not Sent' for {department} on {date_str}")
            return
        
        logger.info(f"📝 Starting mark_all_as_not_sent for {department} on {date_str}")
        
        try:
            ws = self._get_worksheet(department, date_str)
            
            self.ensure_status_column(department, date_str)
            
            headers = ws.row_values(1)
            status_col = None
            for i, header in enumerate(headers, start=1):
                if header == "Report Status":
                    status_col = i
                    break
            
            if status_col is None:
                status_col = len(headers) + 1
                if status_col <= MAX_DATA_COLUMNS:
                    ws.update_cell(1, status_col, "Report Status")
                    self._apply_table_formatting(ws)
                    logger.info(f"✅ Created 'Report Status' column at {status_col}")
                else:
                    logger.warning(f"⚠️ Cannot add 'Report Status' column - would exceed {MAX_DATA_COLUMNS} columns")
                    return
            
            all_values = self._get_cached_worksheet_data(department, date_str)
            logger.info(f"🔍 Found {len(all_values)} rows in sheet")
            
            updates = []
            not_sent_count = 0
            
            for i, row in enumerate(all_values[1:], start=2):
                row_date = row[0].strip() if len(row) > 0 else ""
                if row_date == date_str:
                    has_data = False
                    for col_idx, val in enumerate(row[2:], start=3):
                        if col_idx <= MAX_DATA_COLUMNS and val and val.strip() not in ["", "0", "00:00:00"]:
                            has_data = True
                            break
                    
                    if not has_data and status_col <= MAX_DATA_COLUMNS:
                        col_letter = gspread.utils.rowcol_to_a1(i, status_col).rstrip("0123456789")
                        range_str = f"{col_letter}{i}"
                        updates.append({"range": range_str, "values": [["Not Sent"]]})
                        not_sent_count += 1
            
            if updates:
                logger.info(f"📝 Marking {not_sent_count} employees as 'Not Sent'")
                for update in updates:
                    ws.update(update["range"], update["values"], value_input_option="USER_ENTERED")
                # Format the table after updates (first 12 columns only)
                self._apply_table_formatting(ws)
                self._date_marked_not_sent.add(date_key)
                self._invalidate_cache(department, date_str)
                logger.info(f"✅ Marked {not_sent_count} employees as 'Not Sent' for {date_str}")
            else:
                logger.info(f"ℹ️ No employees to mark as 'Not Sent' for {date_str}")
                
        except Exception as e:
            logger.error(f"❌ Failed to mark 'Not Sent' for {department} on {date_str}: {e}")
            import traceback
            traceback.print_exc()

    @with_retry()
    def ensure_date_for_all_employees(self, department: str, date_str: str) -> None:
        """Add date rows for employees who don't have a row with that date"""
        logger.info(f"📝 ensure_date_for_all_employees: {department} on {date_str}")
        
        employees = SALES_EMPLOYEES
        logger.info(f"👥 Employees to check: {employees}")
        
        try:
            ws = self._get_worksheet(department, date_str)
            
            all_values = self._get_cached_worksheet_data(department, date_str)
            logger.info(f"🔍 Found {len(all_values)} rows in sheet")
            
            has_date = set()
            for row in all_values[1:]:
                if row and row[0].strip() == date_str:
                    name = row[1].strip() if len(row) > 1 else ""
                    if name:
                        has_date.add(name)
            
            logger.info(f"🔍 Employees already have date {date_str}: {has_date}")
            
            missing_employees = [emp for emp in employees if emp not in has_date]
            logger.info(f"🔍 Missing employees: {missing_employees}")
            
            updates = []
            for emp in missing_employees:
                row_num = len(all_values) + 1 + len(updates)
                range_str = f"A{row_num}:B{row_num}"
                updates.append({"range": range_str, "values": [[date_str, emp]]})
            
            if updates:
                logger.info(f"📝 Adding {len(updates)} missing employees")
                ws.batch_update(updates, value_input_option="USER_ENTERED")
                # Format the table after adding rows (first 12 columns only)
                self._apply_table_formatting(ws)
                self._invalidate_cache(department, date_str)
                logger.info(f"✅ Added date for {len(updates)} employees in {department}")
            else:
                logger.info(f"ℹ️ All employees already have rows for {date_str}")
                
        except Exception as e:
            logger.error(f"❌ Failed to ensure employees for {department} on {date_str}: {e}")
            import traceback
            traceback.print_exc()

    @with_retry()
    def find_employee_row(self, department: str, date_str: str, employee_name: str) -> Optional[int]:
        all_values = self._get_cached_worksheet_data(department, date_str)
        for i, row in enumerate(all_values[1:], start=2):
            if row and row[0].strip() == date_str and row[1].strip().lower() == employee_name.lower():
                return i
        return None

    @with_retry()
    def write_batch(self, department: str, date_str: str, updates: List[Tuple[int, Dict]]) -> None:
        if not updates:
            return
        
        logger.info(f"📝 write_batch: {department} on {date_str} - {len(updates)} entries")
        
        try:
            ws = self._get_worksheet(department, date_str)
            mapping = SALES_COLUMN_MAPPING
            
            headers = ws.row_values(1)
            status_col = None
            for i, header in enumerate(headers, start=1):
                if header == "Report Status":
                    status_col = i
                    break
            
            if status_col is None and department == "Sales":
                status_col = len(headers) + 1
                if status_col <= MAX_DATA_COLUMNS:
                    ws.update_cell(1, status_col, "Report Status")
                    self._apply_table_formatting(ws)
                    logger.info(f"Created Report Status column at column {status_col}")
                else:
                    logger.warning(f"⚠️ Cannot create 'Report Status' column - would exceed {MAX_DATA_COLUMNS} columns")
                    status_col = None
            
            batch_requests = []
            
            for row_number, data in updates:
                logger.info(f"🔍 Writing row {row_number}: {data}")
                cell_updates = {}
                
                for field, col in mapping.items():
                    if field in ("Date", "Employee Name"):
                        continue
                    if field == "Report Status":
                        continue
                    if col > MAX_DATA_COLUMNS:
                        # Skip columns beyond 12
                        continue
                    val = data.get(field, 0 if field != "Duration" else "00:00:00")
                    cell_updates[col] = val
                
                if status_col and status_col <= MAX_DATA_COLUMNS:
                    report_status = data.get("report_status", "")
                    cell_updates[status_col] = report_status
                    logger.info(f"🔍 Setting status for row {row_number} to '{report_status}'")
                
                if not cell_updates:
                    continue
                
                cols = sorted(cell_updates)
                # Ensure we don't exceed 12 columns
                cols = [c for c in cols if c <= MAX_DATA_COLUMNS]
                if not cols:
                    continue
                    
                min_col, max_col = cols[0], cols[-1]
                row_values = [cell_updates.get(c, "") for c in range(min_col, max_col + 1)]
                col_start = gspread.utils.rowcol_to_a1(row_number, min_col).rstrip("0123456789")
                col_end = gspread.utils.rowcol_to_a1(row_number, max_col).rstrip("0123456789")
                range_str = f"{col_start}{row_number}:{col_end}{row_number}"
                batch_requests.append({"range": range_str, "values": [row_values]})
            
            if batch_requests:
                logger.info(f"📝 Writing {len(batch_requests)} batch requests")
                ws.batch_update(batch_requests, value_input_option="USER_ENTERED")
                # Format the table after batch writes (first 12 columns only)
                self._apply_table_formatting(ws)
                self._invalidate_cache(department, date_str)
                logger.info(f"✅ Batch wrote {len(batch_requests)} rows to {ws.title} (borders: A-L)")
            else:
                logger.info("ℹ️ No batch requests to write")
                
        except Exception as e:
            logger.error(f"❌ Failed to write batch: {e}")
            import traceback
            traceback.print_exc()

    def mark_not_sent(self, department: str, date_str: str) -> None:
        if department != "Sales":
            return
        self.mark_all_as_not_sent(department, date_str)

    def mark_invalid_report(self, department: str, date_str: str, employee_name: str) -> None:
        if department != "Sales":
            return
        
        try:
            ws = self._get_worksheet(department, date_str)
            
            headers = ws.row_values(1)
            status_col = None
            for i, header in enumerate(headers, start=1):
                if header == "Report Status":
                    status_col = i
                    break
            
            if status_col is None:
                status_col = len(headers) + 1
                if status_col <= MAX_DATA_COLUMNS:
                    ws.update_cell(1, status_col, "Report Status")
                    self._apply_table_formatting(ws)
                else:
                    logger.warning(f"⚠️ Cannot add 'Report Status' column - would exceed {MAX_DATA_COLUMNS} columns")
                    return
            
            all_values = self._get_cached_worksheet_data(department, date_str)
            row_num = None
            for i, row in enumerate(all_values[1:], start=2):
                row_date = row[0].strip() if len(row) > 0 else ""
                row_name = row[1].strip() if len(row) > 1 else ""
                if row_date == date_str and row_name.lower() == employee_name.lower():
                    row_num = i
                    break

            if not row_num:
                logger.warning(f"Row not found for {employee_name} on {date_str}")
                return

            if status_col <= MAX_DATA_COLUMNS:
                col_letter = gspread.utils.rowcol_to_a1(row_num, status_col).rstrip("0123456789")
                range_str = f"{col_letter}{row_num}"
                ws.update(range_str, [["Invalid"]], value_input_option="USER_ENTERED")
                self._apply_table_formatting(ws)
                self._invalidate_cache(department, date_str)
                logger.info(f"✅ Marked {employee_name} as 'Invalid' for {date_str}")
                
        except Exception as e:
            logger.error(f"❌ Failed to mark invalid: {e}")
            import traceback
            traceback.print_exc()

    def mark_quota_error(self, department: str, date_str: str, employee_name: str) -> None:
        if department != "Sales":
            return
        
        try:
            ws = self._get_worksheet(department, date_str)
            
            headers = ws.row_values(1)
            status_col = None
            for i, header in enumerate(headers, start=1):
                if header == "Report Status":
                    status_col = i
                    break
            
            if status_col is None:
                status_col = len(headers) + 1
                if status_col <= MAX_DATA_COLUMNS:
                    ws.update_cell(1, status_col, "Report Status")
                    self._apply_table_formatting(ws)
                else:
                    logger.warning(f"⚠️ Cannot add 'Report Status' column - would exceed {MAX_DATA_COLUMNS} columns")
                    return
            
            all_values = self._get_cached_worksheet_data(department, date_str)
            row_num = None
            for i, row in enumerate(all_values[1:], start=2):
                row_date = row[0].strip() if len(row) > 0 else ""
                row_name = row[1].strip() if len(row) > 1 else ""
                if row_date == date_str and row_name.lower() == employee_name.lower():
                    row_num = i
                    break

            if not row_num:
                logger.warning(f"Row not found for {employee_name} on {date_str}")
                return

            if status_col <= MAX_DATA_COLUMNS:
                col_letter = gspread.utils.rowcol_to_a1(row_num, status_col).rstrip("0123456789")
                range_str = f"{col_letter}{row_num}"
                ws.update(range_str, [["Quota Error"]], value_input_option="USER_ENTERED")
                self._apply_table_formatting(ws)
                self._invalidate_cache(department, date_str)
                logger.info(f"✅ Marked {employee_name} as 'Quota Error' for {date_str}")
                
        except Exception as e:
            logger.error(f"❌ Failed to mark quota error: {e}")
            import traceback
            traceback.print_exc()

    def list_sheets(self, department: str) -> List[str]:
        return [ws.title for ws in self._spreadsheet(department).worksheets()]

    def get_employees_for_date(self, department: str, date_str: str) -> Dict[str, int]:
        all_values = self._get_cached_worksheet_data(department, date_str)
        emp_rows = {}
        for i, row in enumerate(all_values[1:], start=2):
            row_date = row[0].strip() if len(row) > 0 else ""
            row_name = row[1].strip() if len(row) > 1 else ""
            if row_date == date_str and row_name:
                emp_rows[row_name] = i
        return emp_rows
