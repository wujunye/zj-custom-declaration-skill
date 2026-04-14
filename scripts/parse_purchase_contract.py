#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Parse Chinese purchase contract Excel files (.xls or .xlsx) and output structured JSON.
"""

import json
import argparse
import sys
import re
from pathlib import Path

try:
    import xlrd
except ImportError:
    print("Error: xlrd not found. Install it with: pip install xlrd")
    sys.exit(1)

try:
    from openpyxl import load_workbook
except ImportError:
    print("Error: openpyxl not found. Install it with: pip install openpyxl")
    sys.exit(1)


def extract_city_from_supplier(supplier_name):
    """
    Extract city from supplier name using common patterns.
    Examples: '义乌市XXXX有限公司' -> '义乌', 'XX县YYY有限公司' -> 'XX县'
    """
    # Pattern: X市 or XX市
    match = re.search(r'(\w+市)', supplier_name)
    if match:
        return match.group(1).replace('市', '')

    # Pattern: XX县
    match = re.search(r'(\w+县)', supplier_name)
    if match:
        return match.group(1)

    return ""


def parse_package_size(size_str):
    """
    Parse package size string like '43*45*54cm' or '43*45*54' into [43, 45, 54].
    Returns list of integers or None if parsing fails.
    """
    if not size_str or not isinstance(size_str, str):
        return None

    # Remove 'cm' suffix if present
    size_str = size_str.rstrip('cm').strip()

    # Split by * and convert to floats, then to ints
    try:
        parts = [int(float(x.strip())) for x in size_str.split('*')]
        return parts
    except (ValueError, AttributeError):
        return None


def load_excel_sheet(file_path):
    """
    Load Excel sheet from .xls or .xlsx file.
    Returns (sheet, is_openpyxl) tuple.
    """
    file_path = Path(file_path)

    if file_path.suffix.lower() == '.xls':
        # Use xlrd for .xls files
        workbook = xlrd.open_workbook(str(file_path))
        sheet = workbook.sheet_by_index(0)
        return sheet, False
    elif file_path.suffix.lower() == '.xlsx':
        # Use openpyxl for .xlsx files
        workbook = load_workbook(str(file_path))
        sheet = workbook.active
        return sheet, True
    else:
        raise ValueError(f"Unsupported file format: {file_path.suffix}")


def get_cell_value(sheet, row, col, is_openpyxl):
    """
    Get cell value from sheet, handling both xlrd and openpyxl.
    """
    if is_openpyxl:
        cell = sheet.cell(row + 1, col + 1)
        return cell.value
    else:
        return sheet.cell_value(row, col)


def find_header_row(sheet, is_openpyxl):
    """
    Find the header row by looking for '产品名称' in column 0.
    Returns row index or -1 if not found.
    """
    nrows = sheet.max_row if is_openpyxl else sheet.nrows

    for row_idx in range(nrows):
        cell_value = get_cell_value(sheet, row_idx, 0, is_openpyxl)
        if cell_value and isinstance(cell_value, str):
            if '产品名称' in cell_value:
                return row_idx

    return -1


def parse_product_name(name_str):
    """
    Parse product name from purchase contract.
    If the cell contains a newline, take the first line (Chinese name only).
    English name comes from the knowledge base, not from the contract.
    Returns name_cn string.
    """
    if not name_str:
        return ""

    # Take the first line only (Chinese name); ignore anything after newline
    parts = str(name_str).split('\n')
    return parts[0].strip()


def parse_purchase_contract(input_file, output_file=None):
    """
    Parse a Chinese purchase contract Excel file and output JSON.
    """
    try:
        input_path = Path(input_file)
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        # Load sheet
        sheet, is_openpyxl = load_excel_sheet(input_file)

        # Find header row
        header_row = find_header_row(sheet, is_openpyxl)
        if header_row == -1:
            raise ValueError("Could not find header row with '产品名称'")

        # Parse contract metadata (rows 0-3)
        contract_no = str(get_cell_value(sheet, 1, 8, is_openpyxl) or "").strip()
        date = str(get_cell_value(sheet, 2, 3, is_openpyxl) or "").strip()
        supplier_name = str(get_cell_value(sheet, 3, 3, is_openpyxl) or "").strip()
        buyer_name = str(get_cell_value(sheet, 3, 8, is_openpyxl) or "").strip()

        # Extract supplier city
        supplier_city = extract_city_from_supplier(supplier_name)

        # Parse items (from header_row + 1 until "小写合计")
        items = []
        nrows = sheet.max_row if is_openpyxl else sheet.nrows

        item_start = header_row + 1
        for row_idx in range(item_start, nrows):
            # Check if this is the "小写合计" row
            first_cell = str(get_cell_value(sheet, row_idx, 0, is_openpyxl) or "").strip()
            if '小写合计' in first_cell or '合计' in first_cell:
                break

            # Get all cell values for this row
            name = get_cell_value(sheet, row_idx, 0, is_openpyxl)
            spec = get_cell_value(sheet, row_idx, 2, is_openpyxl)
            fba_sku = get_cell_value(sheet, row_idx, 3, is_openpyxl)
            unit = get_cell_value(sheet, row_idx, 4, is_openpyxl)
            quantity = get_cell_value(sheet, row_idx, 5, is_openpyxl)
            packing_rate = get_cell_value(sheet, row_idx, 6, is_openpyxl)
            unit_price = get_cell_value(sheet, row_idx, 7, is_openpyxl)
            package_size = get_cell_value(sheet, row_idx, 8, is_openpyxl)
            net_weight = get_cell_value(sheet, row_idx, 9, is_openpyxl)
            gross_weight = get_cell_value(sheet, row_idx, 10, is_openpyxl)
            total_amount = get_cell_value(sheet, row_idx, 11, is_openpyxl)

            # Skip empty rows
            if not name:
                continue

            # Parse product name (Chinese only; English comes from knowledge base)
            name_cn = parse_product_name(name)

            # Parse package size
            package_size_cm = parse_package_size(package_size)

            # Build item object
            item = {
                "name_cn": name_cn,
                "spec": str(spec).strip() if spec else "",
                "fba_sku": str(fba_sku).strip() if fba_sku else "",
                "unit": str(unit).strip() if unit else "",
                "quantity": int(quantity) if quantity and isinstance(quantity, (int, float)) else 0,
                "packing_rate": int(packing_rate) if packing_rate and isinstance(packing_rate, (int, float)) else 0,
                "unit_price_with_tax": float(unit_price) if unit_price and isinstance(unit_price, (int, float)) else 0.0,
                "package_size_cm": package_size_cm,
                "net_weight_kg": float(net_weight) if net_weight and isinstance(net_weight, (int, float)) else 0.0,
                "gross_weight_kg": float(gross_weight) if gross_weight and isinstance(gross_weight, (int, float)) else 0.0,
                "total_amount": float(total_amount) if total_amount and isinstance(total_amount, (int, float)) else 0.0,
            }

            items.append(item)

        # Find and parse grand total (look for "共计" row)
        grand_total = 0.0
        for row_idx in range(nrows):
            first_cell = str(get_cell_value(sheet, row_idx, 0, is_openpyxl) or "").strip()
            if '共计' in first_cell:
                grand_total_value = get_cell_value(sheet, row_idx, 11, is_openpyxl)
                if grand_total_value and isinstance(grand_total_value, (int, float)):
                    grand_total = float(grand_total_value)
                break

        # Build result JSON
        result = {
            "contract_no": contract_no,
            "date": date,
            "supplier": {
                "name": supplier_name,
                "city": supplier_city,
            },
            "buyer": {
                "name": buyer_name,
            },
            "items": items,
            "grand_total": grand_total,
        }

        return result

    except Exception as e:
        print(f"Error parsing file: {e}", file=sys.stderr)
        raise


def main():
    parser = argparse.ArgumentParser(
        description="Parse Chinese purchase contract Excel files and output structured JSON"
    )
    parser.add_argument(
        "input_file",
        help="Path to input Excel file (.xls or .xlsx)"
    )
    parser.add_argument(
        "--output",
        default="/tmp/purchase_contract.json",
        help="Path to output JSON file (default: /tmp/purchase_contract.json)"
    )

    args = parser.parse_args()

    try:
        result = parse_purchase_contract(args.input_file)

        # Write to output file
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        print(f"Success! Output written to: {output_path}")
        print(json.dumps(result, ensure_ascii=False, indent=2))

    except Exception as e:
        print(f"Fatal error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
