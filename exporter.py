"""
Excel exporter module using openpyxl.
Generates styled workbooks with dropdown data validations, clickable hyperlinks,
frozen header row, auto-filters, and clean typography.
"""
from pathlib import Path
from typing import List, Dict, Any, Union, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from config import DEFAULT_EXCEL_PATH, EXCEL_DROPDOWNS
from db import DatabaseManager


# Modern Color Palette & Typography
HEADER_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")  # Slate 800
HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)

ROW_FONT = Font(name="Segoe UI", size=10, color="0F172A")
LINK_FONT = Font(name="Segoe UI", size=10, color="0284C7", underline="single")  # Sky 600
ROW_EVEN_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")  # Slate 50
ROW_ODD_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

BORDER_THIN = Border(
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
    bottom=Side(style="thin", color="E2E8F0"),
)
HEADER_BORDER = Border(
    left=Side(style="thin", color="334155"),
    right=Side(style="thin", color="334155"),
    top=Side(style="medium", color="0F172A"),
    bottom=Side(style="medium", color="0F172A"),
)


def export_leads_to_excel(
    leads: Union[List[Dict[str, Any]], DatabaseManager, str, Path],
    output_path: Union[str, Path] = DEFAULT_EXCEL_PATH,
    max_validation_rows: int = 1500,
) -> Path:
    """
    Export scraped leads to a professionally formatted Excel spreadsheet with:
    - Columns: Business Name, Phone, Google Maps URL, Business Type, Call Status, Interest, Send
    - Dropdowns on every data row for Call Status, Interest, Send
    - Clickable Google Maps URLs
    - Frozen top row
    - Auto-filter on all columns
    - Auto-adjusted column widths
    """
    target_path = Path(output_path)
    
    # Resolve input leads
    if isinstance(leads, (str, Path)):
        db = DatabaseManager(leads)
        lead_records = db.get_all_leads()
    elif isinstance(leads, DatabaseManager):
        lead_records = leads.get_all_leads()
    else:
        lead_records = list(leads)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Google Maps Leads"

    # Ensure grid lines are visible in Excel
    ws.views.sheetView[0].showGridLines = True

    # Define Columns
    headers = [
        "Business Name",
        "Phone",
        "Google Maps URL",
        "Business Type",
        "Call Status",
        "Interest",
        "Send",
        "Lead Quality",
    ]
    ws.append(headers)

    # Style Header Row (Row 1)
    ws.row_dimensions[1].height = 28
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = HEADER_BORDER

    # Populate Data Rows
    current_row = 2
    for item in lead_records:
        name = item.get("business_name") or ""
        phone = item.get("phone") or ""
        # Prefer short CID URL if available, fallback to full Maps URL
        url = item.get("short_cid_url") or item.get("google_maps_url") or ""
        b_type = item.get("category") or ""
        call_status = item.get("call_status") or ""
        interest = item.get("interest") or ""
        send = item.get("send") or ""
        lead_quality = item.get("lead_quality") or ""

        row_values = [name, phone, url, b_type, call_status, interest, send, lead_quality]
        ws.append(row_values)

        ws.row_dimensions[current_row].height = 20
        fill = ROW_EVEN_FILL if current_row % 2 == 0 else ROW_ODD_FILL

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.border = BORDER_THIN
            cell.fill = fill

            if col_idx == 3 and url:
                # Clickable Hyperlink for Google Maps URL
                cell.hyperlink = url
                cell.font = LINK_FONT
                cell.alignment = Alignment(vertical="center", horizontal="left")
            elif col_idx in (2, 5, 6, 7, 8):
                # Centered alignment for Phone, Status, Interest, Send, Lead Quality
                cell.font = ROW_FONT
                cell.alignment = Alignment(vertical="center", horizontal="center")
            else:
                cell.font = ROW_FONT
                cell.alignment = Alignment(vertical="center", horizontal="left")

        current_row += 1

    last_data_row = max(current_row - 1, 1)
    validation_end_row = max(last_data_row, max_validation_rows)

    # Configure Data Validation Dropdowns
    # 1. Call Status: answered, unanswered, call back
    call_status_opts = '"{}"'.format(",".join(EXCEL_DROPDOWNS["call_status"]))
    dv_call = DataValidation(
        type="list",
        formula1=call_status_opts,
        allow_blank=True,
        showDropDown=None,
    )
    dv_call.error = "Please select a valid Call Status from the dropdown list."
    dv_call.errorTitle = "Invalid Call Status"
    dv_call.prompt = "Choose call outcome"
    dv_call.promptTitle = "Call Status"
    ws.add_data_validation(dv_call)
    dv_call.add(f"E2:E{validation_end_row}")

    # 2. Interest: interested, not interested
    interest_opts = '"{}"'.format(",".join(EXCEL_DROPDOWNS["interest"]))
    dv_interest = DataValidation(
        type="list",
        formula1=interest_opts,
        allow_blank=True,
        showDropDown=None,
    )
    dv_interest.error = "Please select a valid Interest value from the dropdown list."
    dv_interest.errorTitle = "Invalid Interest Option"
    dv_interest.prompt = "Choose lead interest"
    dv_interest.promptTitle = "Interest"
    ws.add_data_validation(dv_interest)
    dv_interest.add(f"F2:F{validation_end_row}")

    # 3. Send: send, not send
    send_opts = '"{}"'.format(",".join(EXCEL_DROPDOWNS["send"]))
    dv_send = DataValidation(
        type="list",
        formula1=send_opts,
        allow_blank=True,
        showDropDown=None,
    )
    dv_send.error = "Please select 'send' or 'not send'."
    dv_send.errorTitle = "Invalid Send Option"
    dv_send.prompt = "Choose send status"
    dv_send.promptTitle = "Send"
    ws.add_data_validation(dv_send)
    dv_send.add(f"G2:G{validation_end_row}")

    # 4. Lead Quality: quality lead, not quality lead
    quality_opts = '"{}"'.format(",".join(EXCEL_DROPDOWNS["lead_quality"]))
    dv_quality = DataValidation(
        type="list",
        formula1=quality_opts,
        allow_blank=True,
        showDropDown=None,
    )
    dv_quality.error = "Please select 'quality lead' or 'not quality lead'."
    dv_quality.errorTitle = "Invalid Quality Option"
    dv_quality.prompt = "Choose lead quality"
    dv_quality.promptTitle = "Lead Quality"
    ws.add_data_validation(dv_quality)
    dv_quality.add(f"H2:H{validation_end_row}")

    # Freeze Top Row (A2 freeze pane)
    ws.freeze_panes = "A2"

    # Add Auto-Filters on Header
    ws.auto_filter.ref = f"A1:G{last_data_row}"

    # Auto-adjust column widths with generous padding
    min_widths = {
        "A": 32,  # Business Name
        "B": 20,  # Phone
        "C": 42,  # Google Maps URL
        "D": 26,  # Business Type
        "E": 18,  # Call Status
        "F": 18,  # Interest
        "G": 16,  # Send
    }

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = str(cell.value or "")
            if len(val) > max_len:
                max_len = len(val)
        computed_width = max(max_len + 4, min_widths.get(col_letter, 15))
        # Cap excessive width for neat view
        ws.column_dimensions[col_letter].width = min(computed_width, 60)

    # Save to disk safely with fallback if file is open in Excel
    try:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(target_path)
    except PermissionError:
        alt_path = target_path.with_name(f"{target_path.stem}_latest{target_path.suffix}")
        try:
            wb.save(alt_path)
            return alt_path
        except Exception:
            pass
    except Exception:
        pass
    return target_path


def generate_excel_bytes(
    leads: Union[List[Dict[str, Any]], DatabaseManager, str, Path]
) -> bytes:
    """
    Generate an Excel workbook directly in memory as bytes.
    Ideal for direct browser downloads without Windows file locking conflicts.
    """
    import io

    # Resolve input leads
    if isinstance(leads, (str, Path)):
        db = DatabaseManager(leads)
        lead_records = db.get_all_leads()
    elif isinstance(leads, DatabaseManager):
        lead_records = leads.get_all_leads()
    else:
        lead_records = list(leads)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Google Maps Leads"
    ws.views.sheetView[0].showGridLines = True

    headers = [
        "Business Name",
        "Phone",
        "Google Maps URL",
        "Business Type",
        "Call Status",
        "Interest",
        "Send",
        "Lead Quality",
    ]
    ws.append(headers)

    ws.row_dimensions[1].height = 28
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = HEADER_BORDER

    current_row = 2
    for item in lead_records:
        name = item.get("business_name") or ""
        phone = item.get("phone") or ""
        maps_url = item.get("short_cid_url") or item.get("google_maps_url") or ""
        b_type = item.get("category") or ""
        call_status = item.get("call_status") or ""
        interest = item.get("interest") or ""
        send = item.get("send") or ""
        quality = item.get("lead_quality") or ""

        row_data = [name, phone, maps_url, b_type, call_status, interest, send, quality]
        ws.append(row_data)

        ws.row_dimensions[current_row].height = 20
        fill_to_use = ROW_EVEN_FILL if current_row % 2 == 0 else ROW_ODD_FILL

        for col_idx in range(1, len(row_data) + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = ROW_FONT
            cell.fill = fill_to_use
            cell.border = BORDER_THIN

            if col_idx in (2, 5, 6, 7, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            if col_idx == 3 and maps_url and maps_url.startswith("http"):
                cell.hyperlink = maps_url
                cell.font = LINK_FONT

        current_row += 1

    last_data_row = max(current_row - 1, 2)
    validation_end_row = max(last_data_row, 1500)

    dv_call = DataValidation(
        type="list",
        formula1=f'"{",".join(EXCEL_DROPDOWNS["call_status"])}"',
        allow_blank=True,
    )
    ws.add_data_validation(dv_call)
    dv_call.add(f"E2:E{validation_end_row}")

    dv_interest = DataValidation(
        type="list",
        formula1=f'"{",".join(EXCEL_DROPDOWNS["interest"])}"',
        allow_blank=True,
    )
    ws.add_data_validation(dv_interest)
    dv_interest.add(f"F2:F{validation_end_row}")

    dv_send = DataValidation(
        type="list",
        formula1=f'"{",".join(EXCEL_DROPDOWNS["send"])}"',
        allow_blank=True,
    )
    ws.add_data_validation(dv_send)
    dv_send.add(f"G2:G{validation_end_row}")

    dv_quality = DataValidation(
        type="list",
        formula1=f'"{",".join(EXCEL_DROPDOWNS["lead_quality"])}"',
        allow_blank=True,
    )
    ws.add_data_validation(dv_quality)
    dv_quality.add(f"H2:H{validation_end_row}")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{last_data_row}"

    min_widths = {
        "A": 32, "B": 20, "C": 42, "D": 26, "E": 18, "F": 18, "G": 16, "H": 20
    }
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        computed_width = max(max_len + 4, min_widths.get(col_letter, 15))
        ws.column_dimensions[col_letter].width = min(computed_width, 60)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
