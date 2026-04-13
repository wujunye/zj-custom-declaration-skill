#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shared helper functions for customs declaration generation.
"""

import sys
from typing import Dict, Tuple


def get_lwh(item: dict) -> Tuple[float, float, float]:
    """Extract length, width, height from package_size_cm array."""
    sz = item.get('package_size_cm', [0, 0, 0])
    if len(sz) >= 3:
        return float(sz[0]), float(sz[1]), float(sz[2])
    return 0.0, 0.0, 0.0


def sku_key(item: dict) -> str:
    """Get the SKU identifier from a contract item."""
    return item.get('fba_sku', item.get('sku', ''))


def build_sku_mapping(contract_items: list, fba_matrix: dict) -> Dict[str, str]:
    """
    Build a mapping from contract SKU → FBA matrix SKU.
    First tries exact match, then falls back to quantity-based matching.
    Returns {contract_sku: matrix_sku}
    """
    contract_skus = {sku_key(item): item['quantity'] for item in contract_items}
    matrix_skus = {sku: sum(wh.values()) for sku, wh in fba_matrix.items()}

    mapping = {}
    unmatched_contract = {}
    unmatched_matrix = dict(matrix_skus)

    # Phase 1: exact match
    for c_sku, c_qty in contract_skus.items():
        if c_sku in matrix_skus:
            mapping[c_sku] = c_sku
            unmatched_matrix.pop(c_sku, None)
        else:
            unmatched_contract[c_sku] = c_qty

    # Phase 2: match remaining by quantity
    for c_sku, c_qty in unmatched_contract.items():
        for m_sku, m_qty in list(unmatched_matrix.items()):
            if c_qty == m_qty:
                mapping[c_sku] = m_sku
                unmatched_matrix.pop(m_sku)
                break

    # Phase 3: if still unmatched, warn and skip
    for c_sku in unmatched_contract:
        if c_sku not in mapping:
            print(f"WARNING: Contract SKU '{c_sku}' has no match in FBA matrix", file=sys.stderr)
            mapping[c_sku] = c_sku  # use as-is, will get 0 qty from matrix

    return mapping
