#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple


def extract_warehouse_code_from_filename(filename: str) -> Optional[str]:
    """Extract warehouse code from filename like FBA199DN1ZLP-MDW2.pdf -> MDW2"""
    match = re.search(r'-([A-Z0-9]{3,4})\.pdf$', filename)
    return match.group(1) if match else None


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from PDF using pdftotext -layout"""
    try:
        result = subprocess.run(
            ['pdftotext', '-layout', pdf_path, '-'],
            capture_output=True,
            text=True,
            timeout=30
        )
        if result.returncode != 0:
            print(f"Warning: pdftotext failed for {pdf_path}: {result.stderr}", file=sys.stderr)
            return ""
        return result.stdout
    except FileNotFoundError:
        print("Error: pdftotext not found. Please install poppler-utils.", file=sys.stderr)
        sys.exit(1)
    except subprocess.TimeoutExpired:
        print(f"Warning: pdftotext timeout for {pdf_path}", file=sys.stderr)
        return ""
    except Exception as e:
        print(f"Warning: Error extracting text from {pdf_path}: {e}", file=sys.stderr)
        return ""


def parse_total_boxes(text: str) -> int:
    """Extract total number of boxes from pattern '共 N 个纸箱'"""
    match = re.search(r'共\s*(\d+)\s*个纸箱', text)
    return int(match.group(1)) if match else 0


def parse_warehouse_code(text: str) -> Optional[str]:
    """Extract warehouse code from text (3-4 char code after FBA: line)"""
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if 'FBA:' in line:
            # Next non-empty line should be the warehouse code
            for j in range(i + 1, min(i + 5, len(lines))):
                next_line = lines[j].strip()
                if next_line and re.match(r'^[A-Z0-9]{3,4}$', next_line):
                    return next_line
    return None


def parse_address(text: str) -> str:
    """Extract US address from the text"""
    lines = text.split('\n')
    address_lines = []
    in_address = False

    for i, line in enumerate(lines):
        if 'FBA:' in line:
            # Skip the FBA line and warehouse code, then collect address lines
            # Look for lines with address content (contains street, city, state, zip patterns)
            for j in range(i + 1, min(i + 10, len(lines))):
                curr_line = lines[j].strip()

                # Skip warehouse codes, MDW2 etc
                if re.match(r'^[A-Z0-9]{3,4}$', curr_line):
                    continue

                # Look for US states or common address patterns
                if (curr_line and not curr_line.startswith('FBA') and
                    not '：' in curr_line and
                    not '发货地' in curr_line and
                    not '目的地' in curr_line and
                    not '纸箱编号' in curr_line and
                    len(curr_line) > 2):

                    # Check if it looks like an address line
                    if (re.search(r'\b(DR|ST|AVE|RD|BLVD|CT|PL|LN|PKWY)\b', curr_line) or
                        re.search(r'[A-Z]{2}\s+\d{5}', curr_line) or
                        re.search(r'^\d+\s+[A-Z]', curr_line)):
                        address_lines.append(curr_line)

                if re.search(r'[A-Z]{2}\s+\d{5}', curr_line):
                    break

    return ', '.join(address_lines)


def parse_pages(text: str) -> List[Dict]:
    """Parse text into per-page (per-box) data"""
    pages = []

    # Split by box number pattern "纸箱编号 N，共 M 个纸箱"
    page_pattern = r'纸箱编号\s*(\d+)，共\s*(\d+)\s*个纸箱'

    # Find all page breaks
    page_boundaries = []
    for match in re.finditer(page_pattern, text):
        page_boundaries.append((match.start(), int(match.group(1)), int(match.group(2))))

    if not page_boundaries:
        # Single page or no clear boundaries, process entire text as one
        page_boundaries = [(0, 1, 1)]

    for idx, (start_pos, box_num, total_boxes) in enumerate(page_boundaries):
        # Get text until next boundary or end
        if idx + 1 < len(page_boundaries):
            end_pos = page_boundaries[idx + 1][0]
            page_text = text[start_pos:end_pos]
        else:
            page_text = text[start_pos:]

        # Extract SKU
        sku_match = re.search(r'Single SKU\s*\n\s*([A-Z0-9\-]+)', page_text)
        sku = sku_match.group(1).strip() if sku_match else None

        # Extract quantity per box
        qty_match = re.search(r'数量\s*(\d+)', page_text)
        qty_per_box = int(qty_match.group(1)) if qty_match else 0

        if sku and qty_per_box > 0:
            pages.append({
                'box_number': box_num,
                'sku': sku,
                'qty_per_box': qty_per_box
            })

    return pages


def aggregate_sku_breakdown(pages: List[Dict]) -> List[Dict]:
    """Group consecutive boxes with same SKU and qty_per_box into one entry"""
    if not pages:
        return []

    breakdown = []
    current_sku = None
    current_qty = None
    box_count = 0

    for page in pages:
        sku = page['sku']
        qty = page['qty_per_box']

        if sku == current_sku and qty == current_qty:
            # Same SKU and qty, increment count
            box_count += 1
        else:
            # Different SKU or qty, save previous and start new
            if current_sku is not None:
                breakdown.append({
                    'sku': current_sku,
                    'boxes': box_count,
                    'qty_per_box': current_qty,
                    'total_qty': box_count * current_qty
                })
            current_sku = sku
            current_qty = qty
            box_count = 1

    # Don't forget the last entry
    if current_sku is not None:
        breakdown.append({
            'sku': current_sku,
            'boxes': box_count,
            'qty_per_box': current_qty,
            'total_qty': box_count * current_qty
        })

    return breakdown


def parse_pdf(pdf_path: str) -> Optional[Dict]:
    """Parse a single FBA PDF file"""
    filename = os.path.basename(pdf_path)

    # Extract text from PDF
    text = extract_text_from_pdf(pdf_path)
    if not text:
        print(f"Warning: Could not extract text from {pdf_path}", file=sys.stderr)
        return None

    # Extract warehouse code (from text first, fallback to filename)
    warehouse_code = parse_warehouse_code(text)
    if not warehouse_code:
        warehouse_code = extract_warehouse_code_from_filename(filename)

    # Extract other data
    total_boxes = parse_total_boxes(text)
    address = parse_address(text)
    pages = parse_pages(text)
    sku_breakdown = aggregate_sku_breakdown(pages)

    if not warehouse_code:
        print(f"Warning: Could not extract warehouse code from {pdf_path}", file=sys.stderr)
        return None

    if not sku_breakdown:
        print(f"Warning: Could not extract SKU data from {pdf_path}", file=sys.stderr)
        return None

    return {
        'file': filename,
        'warehouse_code': warehouse_code,
        'address': address,
        'total_boxes': total_boxes,
        'sku_breakdown': sku_breakdown
    }


def build_matrix(shipments: List[Dict]) -> Dict[str, Dict[str, int]]:
    """Build cross-reference matrix: {sku: {warehouse: total_qty}}"""
    matrix = {}

    for shipment in shipments:
        warehouse = shipment['warehouse_code']

        for item in shipment['sku_breakdown']:
            sku = item['sku']
            total_qty = item['total_qty']

            if sku not in matrix:
                matrix[sku] = {}

            if warehouse not in matrix[sku]:
                matrix[sku][warehouse] = 0

            matrix[sku][warehouse] += total_qty

    return matrix


def main():
    parser = argparse.ArgumentParser(
        description='Parse Amazon FBA shipment PDFs and output structured JSON'
    )
    parser.add_argument(
        'input_path',
        help='Path to folder containing PDFs or a single PDF file'
    )
    parser.add_argument(
        '--output',
        default='/tmp/fba_shipments.json',
        help='Output JSON file path (default: /tmp/fba_shipments.json)'
    )

    args = parser.parse_args()

    # Collect PDF files
    pdf_files = []
    input_path = Path(args.input_path)

    if input_path.is_file():
        if input_path.suffix.lower() == '.pdf':
            pdf_files.append(str(input_path))
        else:
            print(f"Error: {args.input_path} is not a PDF file", file=sys.stderr)
            sys.exit(1)
    elif input_path.is_dir():
        pdf_files = sorted([str(p) for p in input_path.glob('*.pdf')])
        if not pdf_files:
            print(f"Warning: No PDF files found in {args.input_path}", file=sys.stderr)
    else:
        print(f"Error: {args.input_path} does not exist", file=sys.stderr)
        sys.exit(1)

    # Parse each PDF
    shipments = []
    for pdf_file in pdf_files:
        print(f"Processing {os.path.basename(pdf_file)}...", file=sys.stderr)
        result = parse_pdf(pdf_file)
        if result:
            shipments.append(result)

    if not shipments:
        print("Error: No shipments could be parsed", file=sys.stderr)
        sys.exit(1)

    # Build matrix
    matrix = build_matrix(shipments)

    # Build output
    output = {
        'shipments': shipments,
        'matrix': matrix
    }

    # Write output
    try:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(output, f, indent=2, ensure_ascii=False)
        print(f"Output written to {args.output}", file=sys.stderr)
    except Exception as e:
        print(f"Error writing output: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
