#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate all customs declaration documents:
1. Export Contract (出口合同)
2. Invoice & Packing List (IV&PL)
3. Declaration Draft (报关单草稿)

Input: parsed purchase_contract.json + fba_shipments.json + user parameters
Output: 3 Excel files per selected ticket group

This is the CLI entry point. All generation logic lives in:
  - helpers.py            (shared utilities)
  - generator_base.py     (core class & computation)
  - gen_export_contract.py (出口合同)
  - gen_iv_pl.py          (IV&PL)
  - gen_declaration.py    (报关单草稿)
"""

import argparse
import json
import os
import sys

# Ensure the scripts directory is on the import path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generator_base import CustomsDeclarationGenerator


def main():
    parser = argparse.ArgumentParser(description='Generate customs declaration documents')
    parser.add_argument('--contract', required=True, help='Path to purchase_contract.json')
    parser.add_argument('--shipments', required=True, help='Path to fba_shipments.json')
    parser.add_argument('--knowledge-base', help='Path to knowledge base Excel')
    parser.add_argument('--groups', required=True, help='JSON string defining ticket groups')
    parser.add_argument('--selected-groups', required=True, help='Comma-separated group indices')
    parser.add_argument('--exchange-rate', type=float, required=True)
    parser.add_argument('--shipping-rate', type=float, required=True, help='RMB per kg')
    parser.add_argument('--output-dir', required=True)
    parser.add_argument('--template-dir', help='(reserved for future)')
    parser.add_argument('--price-term', default='CNF', choices=['CNF', 'FOB'])

    args = parser.parse_args()

    groups = json.loads(args.groups)
    selected = [int(x.strip()) for x in args.selected_groups.split(',')]

    gen = CustomsDeclarationGenerator(
        contract_json=args.contract,
        shipments_json=args.shipments,
        groups=groups,
        selected_group_indices=selected,
        exchange_rate=args.exchange_rate,
        shipping_rate=args.shipping_rate,
        output_dir=args.output_dir,
        price_term=args.price_term,
        knowledge_base=args.knowledge_base,
        template_dir=args.template_dir,
    )

    print(gen.generate())


if __name__ == '__main__':
    main()
