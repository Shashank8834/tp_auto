"""
Parser for Annexure 3 - Tested Party Margins (TNMM) Excel file.
Extracts financial data and connected persons to auto-fill the form.
"""

import openpyxl
from typing import Any


def parse_annexure3(file_path: str) -> dict[str, Any]:
    """
    Parse Annexure 3 Excel file and extract:
    - Financial data (operating revenue, costs, expenses, salaries)
    - Connected persons (partners with designations, remuneration, roles)
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    # Read all cell values into a grid for flexible parsing
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))

    financials = {}
    connected_persons = []

    # Parse by scanning for known labels in column A (index 0)
    for i, row in enumerate(rows):
        label = str(row[0] or "").strip().lower() if row[0] else ""
        value = row[1]

        if "operating revenue" in label and "total" not in label:
            # Could be the line item or the total — grab the first numeric one
            if value and isinstance(value, (int, float)):
                financials.setdefault("operating_revenue", value)

        elif "total operating revenue" in label:
            if value and isinstance(value, (int, float)):
                financials["operating_revenue"] = value

        elif "cost of sales" in label:
            if value and isinstance(value, (int, float)):
                financials["cost_of_sales"] = value

        elif "admin" in label and "general" in label and "expense" in label:
            if value and isinstance(value, (int, float)):
                financials["admin_expenses"] = value

        elif "other expense" in label:
            if value and isinstance(value, (int, float)):
                financials["other_expenses"] = value

        elif "staff salary" in label or "staff salaries" in label:
            if value and isinstance(value, (int, float)):
                financials["staff_salary"] = value

        elif "partner" in label and "salar" in label:
            if value and isinstance(value, (int, float)):
                financials["partner_salaries"] = value

    # Parse connected persons from the partner details table (columns D-G, typically rows 12-14)
    # Look for rows where column D has a person name (Mr./Mrs./Ms.) and column E has a designation
    for i, row in enumerate(rows):
        # Ensure enough columns exist
        if len(row) < 6:
            continue

        name = str(row[3] or "").strip() if row[3] else ""
        designation = str(row[4] or "").strip() if row[4] else ""
        remuneration = row[5]
        roles = str(row[6] or "").strip() if len(row) > 6 and row[6] else ""

        # Detect person rows: has a name with title prefix or designation contains known titles
        if name and designation and remuneration and isinstance(remuneration, (int, float)):
            if any(prefix in name for prefix in ["Mr.", "Mrs.", "Ms.", "Dr.", "Prof."]) or \
               any(d in designation.lower() for d in ["partner", "director", "manager", "ceo", "cfo", "coo"]):
                connected_persons.append({
                    "name": name,
                    "designation": designation,
                    "remuneration": str(remuneration),
                    "roles": roles if roles and roles != "-" else "",
                })

    wb.close()

    # Default missing financial fields to 0
    for field in ["operating_revenue", "cost_of_sales", "admin_expenses", "other_expenses", "staff_salary", "partner_salaries"]:
        financials.setdefault(field, 0)

    return {
        "financials": financials,
        "connected_persons": connected_persons,
    }
