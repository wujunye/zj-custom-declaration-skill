#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script for gen_export_contract.py — generates a sample 出口合同 (export contract)
and verifies the output Excel structure and styling.
"""

import os
import sys
import tempfile

# Ensure the scripts directory is on the import path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook
from gen_export_contract import gen_export_contract
from helpers import sku_key, get_lwh

# ─── Sample data (modeled after PO2603230466) ────────────────────────────────

SAMPLE_ITEMS = [
    {
        "name_cn": "人造草坪拼接地板",
        "name_en": "Artificial Grass Interlocking Floor Tiles",
        "spec": "30*30cm 9pcs",
        "fba_sku": "AG-TILE-3030-9P",
        "unit": "件",
        "quantity": 3600,
        "packing_rate": 12,
        "unit_price_with_tax": 35.50,
        "package_size_cm": [43, 45, 54],
        "net_weight_kg": 8.5,
        "gross_weight_kg": 9.2,
        "total_amount": 127800.00,
    },
    {
        "name_cn": "塑料草坪围栏",
        "name_en": "Plastic Grass Fence Panel",
        "spec": "60*40cm",
        "fba_sku": "PG-FENCE-6040",
        "unit": "件",
        "quantity": 2400,
        "packing_rate": 8,
        "unit_price_with_tax": 28.00,
        "package_size_cm": [62, 42, 35],
        "net_weight_kg": 7.0,
        "gross_weight_kg": 7.8,
        "total_amount": 67200.00,
    },
    {
        "name_cn": "仿真植物墙装饰",
        "name_en": "Artificial Plant Wall Decor",
        "spec": "40*60cm",
        "fba_sku": "AP-WALL-4060",
        "unit": "件",
        "quantity": 1800,
        "packing_rate": 6,
        "unit_price_with_tax": 42.00,
        "package_size_cm": [65, 45, 30],
        "net_weight_kg": 6.5,
        "gross_weight_kg": 7.2,
        "total_amount": 75600.00,
    },
]

SAMPLE_CONTRACT = {
    "contract_no": "PO2603230466",
    "date": "2026-03-23",
    "supplier": {"name": "义乌市绿美工艺品有限公司", "city": "义乌"},
    "buyer": {"name": "ZEATALINE INTERNATIONAL TRADING"},
    "grand_total": 270600.00,
}

SAMPLE_KB = {
    "AG-TILE-3030-9P": {
        "tariff_code": "3918909000",
        "english_name": "Artificial Grass Interlocking Floor Tiles",
        "declaration_elements": "0|0|塑料|人造草坪拼接地板|无品牌|无型号",
        "material": "plastic",
    },
    "PG-FENCE-6040": {
        "tariff_code": "3926909090",
        "english_name": "Plastic Grass Fence Panel",
        "declaration_elements": "0|0|塑料|塑料草坪围栏|无品牌|无型号",
        "material": "plastic",
    },
    "AP-WALL-4060": {
        "tariff_code": "6702100000",
        "english_name": "Artificial Plant Wall Decor",
        "declaration_elements": "0|0|塑料|仿真植物墙|无品牌|无型号",
        "material": "plastic",
    },
}

EXCHANGE_RATE = 7.25
SHIPPING_RATE = 8.5

TQ = {sku_key(item): item["quantity"] for item in SAMPLE_ITEMS}
TA = {sku_key(item): item["total_amount"] for item in SAMPLE_ITEMS}


def _compute_shipping(items, ship_rate):
    total_gross = 0.0
    total_vol = 0.0
    for item in items:
        pr = item.get("packing_rate", 1) or 1
        boxes = item["quantity"] / pr
        l, w, h = get_lwh(item)
        total_gross += item["gross_weight_kg"] * boxes
        total_vol += (l * w * h / 6000) * boxes
    chargeable = max(total_gross, total_vol)
    total_ship = chargeable * ship_rate
    use_vol = total_vol > total_gross
    alloc = {}
    for item in items:
        sku = sku_key(item)
        pr = item.get("packing_rate", 1) or 1
        boxes = item["quantity"] / pr
        l, w, h = get_lwh(item)
        sku_vol = (l * w * h / 6000) * boxes
        sku_gross = item["gross_weight_kg"] * boxes
        if use_vol:
            prop = sku_vol / total_vol if total_vol > 0 else 0
        else:
            prop = sku_gross / total_gross if total_gross > 0 else 0
        alloc[sku] = total_ship * prop
    return alloc, total_gross, total_vol, chargeable, total_ship


SHIP_ALLOC, TOTAL_GROSS, TOTAL_VOL, CHARGEABLE, TOTAL_SHIP = _compute_shipping(SAMPLE_ITEMS, SHIPPING_RATE)


# ─── Test ─────────────────────────────────────────────────────────────────────

def test_gen_export_contract():
    with tempfile.TemporaryDirectory() as tmpdir:
        fp = gen_export_contract(
            items=SAMPLE_ITEMS,
            contract=SAMPLE_CONTRACT,
            kb=SAMPLE_KB,
            cno="PO2603230466",
            suffix="",
            tq=TQ,
            ta=TA,
            ship_alloc=SHIP_ALLOC,
            total_gross=TOTAL_GROSS,
            total_vol=TOTAL_VOL,
            chargeable=CHARGEABLE,
            total_ship=TOTAL_SHIP,
            rate=EXCHANGE_RATE,
            ship_rate=SHIPPING_RATE,
            out_dir=tmpdir,
        )

        # 1) File exists
        assert os.path.exists(fp), f"Output file not created: {fp}"
        print(f"[PASS] File created: {os.path.basename(fp)}")

        # 2) Load workbook
        wb = load_workbook(fp)
        ws = wb.active
        assert ws.title == "出口合同", f"Unexpected sheet title: {ws.title}"
        print(f"[PASS] Sheet title: {ws.title}")

        # 3) Title: '采购合同'
        assert ws["A1"].value == "采购合同"
        assert ws["A1"].font.size == 24
        assert ws["A1"].font.name == "宋体"
        assert ws["A1"].alignment.horizontal == "center"
        print("[PASS] Title: '采购合同', font 宋体 size 24, centered")

        # 4) Contract number
        assert ws["I2"].value == "PO2603230466"
        print(f"[PASS] Contract number: {ws['I2'].value}")

        # 5) Date
        assert ws["D3"].value == "2026-03-23"
        print(f"[PASS] Date: {ws['D3'].value}")

        # 6) Supplier and buyer
        assert ws["D4"].value == "义乌市绿美工艺品有限公司"
        assert ws["I4"].value == "ZEATALINE INTERNATIONAL TRADING"
        print(f"[PASS] Supplier: {ws['D4'].value}")
        print(f"[PASS] Buyer: {ws['I4'].value}")

        # 7) Column headers at row 11
        expected_headers = ['产品名称', '产品图片', '规格型号', 'FBA SKU', '单位', '数量',
                           '箱率', '含税单价/元', '包装尺寸/CM', '外箱净重/KG', '外箱毛重/KG', '总额/元']
        for c, h in enumerate(expected_headers, 1):
            actual = ws.cell(row=11, column=c).value
            assert actual == h, f"Header col {c} mismatch: {actual} != {h}"
        print("[PASS] All 12 column headers correct at row 11")

        # 8) Item data rows starting at row 12
        row = 12
        sum_qty = 0
        sum_amt = 0.0
        for idx, item in enumerate(SAMPLE_ITEMS):
            sku = sku_key(item)
            qty = TQ[sku]
            amt = TA[sku]

            # Column A: name_cn + name_en
            name_val = ws.cell(row=row, column=1).value
            assert item["name_cn"] in str(name_val), f"Item {idx+1}: name_cn not found"
            assert item["name_en"] in str(name_val), f"Item {idx+1}: name_en not found"

            # Column C: spec
            assert ws.cell(row=row, column=3).value == item["spec"]

            # Column D: SKU
            assert ws.cell(row=row, column=4).value == sku

            # Column E: unit
            assert ws.cell(row=row, column=5).value == item["unit"]

            # Column F: quantity
            assert ws.cell(row=row, column=6).value == qty

            # Column G: packing rate
            assert ws.cell(row=row, column=7).value == item["packing_rate"]

            # Column H: unit price (含税单价 = total_amount / qty)
            unit_price = ws.cell(row=row, column=8).value
            expected_up = round(amt / qty, 2)
            assert abs(unit_price - expected_up) < 0.01, f"Item {idx+1}: unit price {unit_price} != {expected_up}"

            # Column I: package size
            l, w, h = item["package_size_cm"]
            expected_size = f"{int(l)}*{int(w)}*{int(h)}"
            assert ws.cell(row=row, column=9).value == expected_size

            # Column J: net weight
            assert ws.cell(row=row, column=10).value == item["net_weight_kg"]

            # Column K: gross weight
            assert ws.cell(row=row, column=11).value == item["gross_weight_kg"]

            # Column L: total amount
            total_amt = ws.cell(row=row, column=12).value
            assert abs(total_amt - round(amt, 2)) < 0.01

            # Right-side calculation columns
            # Column T (19): volume weight > 0
            vol_w = ws.cell(row=row, column=19).value
            assert vol_w > 0, f"Item {idx+1}: volume weight should be > 0"

            # Column U (20): shipping alloc > 0
            ship_val = ws.cell(row=row, column=20).value
            assert ship_val > 0, f"Item {idx+1}: shipping alloc should be > 0"

            # Column V (21): C&F total > 0
            cnf_total = ws.cell(row=row, column=21).value
            assert cnf_total > 0, f"Item {idx+1}: C&F total should be > 0"

            # Column W (22): C&F unit > 0
            cnf_unit = ws.cell(row=row, column=22).value
            assert cnf_unit > 0, f"Item {idx+1}: C&F unit price should be > 0"

            print(f"[PASS] Item {idx+1} ({sku}): all fields correct, C&F unit={cnf_unit:.2f}")

            sum_qty += qty
            sum_amt += amt
            row += 1

        # 9) Totals row — "小写合计"
        assert ws.cell(row=row, column=1).value == "小写合计"
        assert ws.cell(row=row, column=6).value == sum_qty
        print(f"[PASS] Subtotal row: qty={sum_qty}")

        # 10) "大写合计" row
        row += 1
        assert ws.cell(row=row, column=1).value == "大写合计"
        print("[PASS] Grand total label row present")

        # 11) "共计" row with total amount
        row += 2
        assert ws.cell(row=row, column=1).value == "共计"
        grand_total = ws.cell(row=row, column=12).value
        assert abs(grand_total - round(sum_amt, 2)) < 0.01
        print(f"[PASS] Grand total: {grand_total}")

        # 12) Exchange rate validation cell
        rate_label = ws.cell(row=row, column=14).value
        assert f"{EXCHANGE_RATE}" in str(rate_label)
        print(f"[PASS] Exchange rate label: {rate_label}")

        # 13) Calculation summary labels (rows 13-16, col N)
        assert ws.cell(row=13, column=14).value == "总毛重"
        assert ws.cell(row=13, column=15).value == round(TOTAL_GROSS, 1)
        assert ws.cell(row=14, column=14).value == "总体积重"
        assert ws.cell(row=14, column=15).value == round(TOTAL_VOL, 2)
        assert ws.cell(row=15, column=14).value == "计费重"
        assert ws.cell(row=15, column=15).value == round(CHARGEABLE, 2)
        assert "总海运费" in str(ws.cell(row=16, column=14).value)
        assert ws.cell(row=16, column=15).value == round(TOTAL_SHIP, 2)
        print("[PASS] Calculation summary: 总毛重/总体积重/计费重/总海运费 all correct")

        # 14) Font checks — all content should be 宋体 size 10
        assert ws.cell(row=12, column=1).font.name == "宋体"
        assert ws.cell(row=12, column=1).font.size == 10
        print("[PASS] Data font: 宋体 size 10")

        # 15) Row heights
        assert ws.row_dimensions[1].height == 42
        print(f"[PASS] Row 1 height: {ws.row_dimensions[1].height}")

        print("\n" + "=" * 60)
        print("ALL TESTS PASSED for gen_export_contract")
        print("=" * 60)


if __name__ == "__main__":
    test_gen_export_contract()
