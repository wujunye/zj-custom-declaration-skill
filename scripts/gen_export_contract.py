#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate Export Contract (出口合同) Excel document.
"""

import os
from typing import Dict, List, Any, Tuple
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from helpers import get_lwh, sku_key


def gen_export_contract(
    items: List[dict],
    contract: dict,
    kb: dict,
    cno: str,
    suffix: str,
    tq: Dict[str, int],
    ta: Dict[str, float],
    ship_alloc: Dict[str, float],
    total_gross: float,
    total_vol: float,
    chargeable: float,
    total_ship: float,
    rate: float,
    ship_rate: float,
    out_dir: str,
) -> str:
    """Generate the export contract Excel file. Returns the output file path."""

    def _info(sku: str, item: dict) -> dict:
        if sku in kb:
            return kb[sku]
        return {
            'tariff_code': '',
            'english_name': item.get('name_en', ''),
            'declaration_elements': '0|0|塑料|塑料草坪|无品牌|无型号',
            'material': 'plastic',
        }

    wb = Workbook()
    ws = wb.active
    ws.title = "出口合同"

    # Set column widths (converted from xlrd units ~256 ≈ openpyxl char width)
    col_widths = {
        0: 5441/256, 1: 4053/256, 2: 3626/256, 3: 4864/256, 4: 2304/256, 5: 2474/256,
        6: 2218/256, 7: 4266/256, 8: 4736/256, 9: 3584/256, 10: 3754/256, 11: 4053/256,
        13: 6869/256, 14: 2730/256, 18: 2773/256, 19: 3242/256, 20: 3669/256, 21: 3114/256
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx + 1)].width = width

    # Set row heights
    row_heights = {
        0: 42, 1: 20, 2: 20, 3: 20, 4: 27, 5: 20, 6: 20, 7: 20, 8: 20, 9: 20, 10: 20,
        11: 24, 12: 24, 13: 24, 14: 24, 15: 24, 16: 13.5, 17: 13.5, 18: 25, 19: 25
    }
    for row_idx, height in row_heights.items():
        ws.row_dimensions[row_idx + 1].height = height

    # Header section (mimic purchase contract layout)
    ws['A1'] = '采购合同'
    ws['A1'].font = Font(name='宋体', size=24, bold=False)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells('A1:L1')

    ws['H2'] = '合同编号：'
    ws['H2'].font = Font(name='宋体', size=10)
    ws['I2'] = cno
    ws['I2'].font = Font(name='宋体', size=10)
    ws['A3'] = '日期：'
    ws['A3'].font = Font(name='宋体', size=10)
    ws['D3'] = contract.get('date', '')
    ws['D3'].font = Font(name='宋体', size=10)

    supplier = contract.get('supplier', {})
    buyer = contract.get('buyer', {})
    ws['A4'] = '供方：'
    ws['A4'].font = Font(name='宋体', size=10)
    ws['D4'] = supplier.get('name', '')
    ws['D4'].font = Font(name='宋体', size=10)
    ws['H4'] = '需方：'
    ws['H4'].font = Font(name='宋体', size=10)
    ws['I4'] = buyer.get('name', '')
    ws['I4'].font = Font(name='宋体', size=10)

    # Column headers (row 11, 1-indexed in openpyxl)
    headers_main = ['产品名称', '产品图片', '规格型号', 'FBA SKU', '单位', '数量',
                    '箱率', '含税单价/元', '包装尺寸/CM', '外箱净重/KG', '外箱毛重/KG', '总额/元']
    headers_calc = ['', '计算逻辑', '', '箱数-计算总海运费使用', '', '', '体积重', '运费平摊', 'C&F总价', 'C&F单价']

    for c, h in enumerate(headers_main, 1):
        cell = ws.cell(row=11, column=c, value=h)
        cell.font = Font(name='宋体', size=10, bold=False)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for c, h in enumerate(headers_calc, 13):
        cell = ws.cell(row=11, column=c, value=h)
        cell.font = Font(name='宋体', size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Item rows
    row = 12
    sum_qty = 0
    sum_amt = 0.0
    calc_labels_written = False

    for item in items:
        sku = sku_key(item)
        qty = tq.get(sku, 0)
        if qty == 0:
            continue

        amt = ta.get(sku, 0)
        unit_price = amt / qty if qty > 0 else 0
        l, w, h_ = get_lwh(item)
        size_str = f'{int(l)}*{int(w)}*{int(h_)}'

        # Full contract boxes (for shipping calc)
        pr = item.get('packing_rate', 1) or 1
        full_boxes = item['quantity'] / pr
        vol_weight = (l * w * h_ / 6000) * full_boxes
        shipping = ship_alloc.get(sku, 0)
        tax_excl = amt / 1.13
        cnf_total = tax_excl + shipping
        cnf_unit = cnf_total / qty if qty > 0 else 0

        name_cn = item.get('name_cn', '')
        name_en = item.get('name_en', '')
        c1 = ws.cell(row=row, column=1, value=f'{name_cn}\n{name_en}')
        c1.font = Font(name='宋体', size=10)
        c1.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

        ws.cell(row=row, column=3, value=item.get('spec', '')).font = Font(name='宋体', size=10)
        ws.cell(row=row, column=4, value=sku).font = Font(name='宋体', size=10)
        ws.cell(row=row, column=5, value=item.get('unit', '件')).font = Font(name='宋体', size=10)
        c6 = ws.cell(row=row, column=6, value=qty)
        c6.font = Font(name='宋体', size=10)
        c6.number_format = '#,##0'
        c6.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=7, value=int(pr)).font = Font(name='宋体', size=10)
        c8 = ws.cell(row=row, column=8, value=round(unit_price, 2))
        c8.font = Font(name='宋体', size=10)
        c8.number_format = '0.00_ '
        c8.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=9, value=size_str).font = Font(name='宋体', size=10)
        c10 = ws.cell(row=row, column=10, value=item.get('net_weight_kg', 0))
        c10.font = Font(name='宋体', size=10)
        c10.number_format = '0.00_ '
        c11 = ws.cell(row=row, column=11, value=item.get('gross_weight_kg', 0))
        c11.font = Font(name='宋体', size=10)
        c11.number_format = '0.00_ '
        c12 = ws.cell(row=row, column=12, value=round(amt, 2))
        c12.font = Font(name='宋体', size=10)
        c12.number_format = '0.0000_ '
        c12.alignment = Alignment(horizontal='center', vertical='center')

        # Right-side calculation columns
        ws.cell(row=row, column=16, value=round(full_boxes)).font = Font(name='宋体', size=10)
        ws.cell(row=row, column=19, value=round(vol_weight, 2)).font = Font(name='宋体', size=10)
        ws.cell(row=row, column=20, value=round(shipping, 2)).font = Font(name='宋体', size=10)
        ws.cell(row=row, column=21, value=round(cnf_total, 2)).font = Font(name='宋体', size=10)
        ws.cell(row=row, column=22, value=round(cnf_unit, 2)).font = Font(name='宋体', size=10)

        # Write calculation labels on first data rows
        if not calc_labels_written:
            ws.cell(row=12, column=14, value='').font = Font(name='宋体', size=10)
            ws.cell(row=13, column=14, value='总毛重').font = Font(name='宋体', size=10)
            ws.cell(row=13, column=15, value=round(total_gross, 1)).font = Font(name='宋体', size=10)
            ws.cell(row=14, column=14, value='总体积重').font = Font(name='宋体', size=10)
            ws.cell(row=14, column=15, value=round(total_vol, 2)).font = Font(name='宋体', size=10)
            ws.cell(row=15, column=14, value='计费重').font = Font(name='宋体', size=10)
            ws.cell(row=15, column=15, value=round(chargeable, 2)).font = Font(name='宋体', size=10)
            ws.cell(row=16, column=14, value=f'总海运费（单价{ship_rate}元/kg）').font = Font(name='宋体', size=10)
            ws.cell(row=16, column=15, value=round(total_ship, 2)).font = Font(name='宋体', size=10)
            calc_labels_written = True

        sum_qty += qty
        sum_amt += amt
        row += 1

    # Totals
    ws.cell(row=row, column=1, value='小写合计').font = Font(name='宋体', size=10)
    ws.cell(row=row, column=6, value=sum_qty).font = Font(name='宋体', size=10)
    row += 1
    ws.cell(row=row, column=1, value='大写合计').font = Font(name='宋体', size=10)
    row += 2
    ws.cell(row=row, column=1, value='共计').font = Font(name='宋体', size=10)
    ws.cell(row=row, column=12, value=round(sum_amt, 2)).font = Font(name='宋体', size=10)

    # Exchange rate validation
    ws.cell(row=row, column=14, value=f'报关汇率{rate}').font = Font(name='宋体', size=10)
    tax_excl_total = sum_amt / 1.13
    inv_usd = (tax_excl_total + total_ship) / rate
    ws.cell(row=row, column=15, value=round(tax_excl_total + total_ship, 2)).font = Font(name='宋体', size=10)

    fn = f'【{cno}】{suffix}出口合同.xlsx'
    fp = os.path.join(out_dir, fn)
    wb.save(fp)
    return fp
