#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate Invoice & Packing List (IV&PL) Excel document.
"""

import os
from typing import Dict, List
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from datetime import datetime

from helpers import get_lwh, sku_key


def gen_iv_pl(
    items: List[dict],
    kb: dict,
    cno: str,
    suffix: str,
    tq: Dict[str, int],
    ta: Dict[str, float],
    ship_alloc: Dict[str, float],
    rate: float,
    out_dir: str,
) -> str:
    """Generate the IV&PL Excel file. Returns the output file path."""

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
    wb.remove(wb.active)

    iv = wb.create_sheet('IV')
    _fill_iv(iv, items, kb, cno, tq, ta, ship_alloc, rate, _info)

    pl = wb.create_sheet('PL')
    _fill_pl(pl, items, kb, cno, tq, _info)

    fn = f'【{cno}】{suffix}IV&PL.xlsx'
    fp = os.path.join(out_dir, fn)
    wb.save(fp)
    return fp


def _fill_iv(ws, items, kb, cno, tq, ta, ship_alloc, rate, info_fn):
    """Fill the Invoice worksheet."""
    # Set column widths
    col_widths = {0: 4793/256, 1: 6400/256, 2: 6609/256, 3: 2669/256, 4: 3257/256,
                  5: 3840/256, 6: 3072/256, 7: 3769/256, 8: 5678/256, 9: 4430/256}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx + 1)].width = width

    # Set row heights
    row_heights = {0: 30, 1: 30, 2: 70, 3: 30, 4: 30, 5: 67, 6: 30, 7: 30, 8: 51, 9: 56}
    for row_idx, height in row_heights.items():
        ws.row_dimensions[row_idx + 1].height = height
    for r in range(10, 13):
        ws.row_dimensions[r + 1].height = 59
    ws.row_dimensions[14].height = 30
    ws.row_dimensions[15].height = 30

    # Title
    ws['A1'] = 'INVOICE'
    ws['A1'].font = Font(name='Arial', size=16, bold=False)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='top')
    ws.merge_cells('A1:I1')

    # Header - with merges
    ws['A2'] = 'Shipper:公司名称'
    ws['A2'].font = Font(name='Arial', size=12)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells('B2:C2')
    ws['B2'] = 'Shenzhen Adhoc Trading Co., Ltd.'
    ws['B2'].font = Font(name='Arial', size=10)
    ws['B2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D2:E2')
    ws['D2'] = 'INVOICE NO:'
    ws['D2'].font = Font(name='Arial', size=12)
    ws['D2'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws.merge_cells('F2:G2')
    ws['F2'] = cno
    ws['F2'].font = Font(name='Arial', size=12)
    ws['F2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws['A3'] = 'ADD:地址'
    ws['A3'].font = Font(name='Arial', size=12)
    ws['A3'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells('B3:C3')
    ws['B3'] = 'Flat 1006, Zhenye International Business Centre,No.3101-90,Qianhai Road, Nanshan District, Shenzhen,China.'
    ws['B3'].font = Font(name='Arial', size=10)
    ws['B3'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D3:E3')
    ws['D3'] = 'DATE:'
    ws['D3'].font = Font(name='Arial', size=12)
    ws['D3'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws.merge_cells('F3:G3')
    ws['F3'] = datetime.now().strftime('%Y/%m/%d')
    ws['F3'].font = Font(name='Arial', size=12)
    ws['F3'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.merge_cells('D4:E4')
    ws['D4'] = 'ORIGIN COUNTRY :'
    ws['D4'].font = Font(name='Arial', size=12)
    ws['D4'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws.merge_cells('F4:G4')
    ws['F4'] = 'CHINA'
    ws['F4'].font = Font(name='Arial', size=12)
    ws['F4'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws['A5'] = 'CNEE:公司名称'
    ws['A5'].font = Font(name='Arial', size=12)
    ws['A5'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.merge_cells('B5:C5')
    ws['B5'] = 'ZEATALINE INTERNATIONAL TRADING'
    ws['B5'].font = Font(name='Arial', size=10)
    ws['B5'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D5:E5')
    ws['D5'] = 'PRICE TERMS:'
    ws['D5'].font = Font(name='Arial', size=12)
    ws['D5'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws['F5'] = 'C&F'
    ws['F5'].font = Font(name='Arial', size=11)
    ws['F5'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws['G5'] = 'USA'
    ws['G5'].font = Font(name='宋体', size=10)
    ws['G5'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws['A6'] = 'ADD:地址'
    ws['A6'].font = Font(name='Arial', size=12)
    ws['A6'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells('B6:C6')
    ws['B6'] = 'zouzhichang\n890 S Azusa Ave\nCity of Industry, CA 91748\nUS'
    ws['B6'].font = Font(name='Arial', size=10)
    ws['B6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D6:E6')
    ws['D6'] = 'Reference No.'
    ws['D6'].font = Font(name='Arial', size=12)
    ws['D6'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)

    # Column headers - English
    hdrs = ['No.', 'Tariff Code', 'Descriptions', 'Qty', 'Unit',
            'Unit Price', 'USD', 'Total Amount', 'material quality', 'picture']
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=7, column=c, value=h)
        cell.font = Font(name='Arial', size=12, bold=False)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Column headers - Chinese
    cn_hdrs = ['', '海关编码（清关用的）', '英文品名', '数量', 'PC(S)',
               '单价', 'USD', '总价', '材质', '产品图片']
    for c, h in enumerate(cn_hdrs, 1):
        cell = ws.cell(row=8, column=c, value=h)
        cell.font = Font(name='宋体', size=12, bold=False)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    row = 9
    total_usd = 0.0
    total_qty = 0
    idx = 1

    for item in items:
        sku = sku_key(item)
        qty = tq.get(sku, 0)
        if qty == 0:
            continue

        amt = ta.get(sku, 0)
        shipping = ship_alloc.get(sku, 0)
        tax_excl = amt / 1.13
        cnf_total_rmb = tax_excl + shipping
        unit_usd = cnf_total_rmb / qty / rate if qty > 0 else 0
        total_item_usd = cnf_total_rmb / rate

        info = info_fn(sku, item)

        c1 = ws.cell(row=row, column=1, value=idx)
        c1.font = Font(name='Arial', size=10)
        c1.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=2, value=info['tariff_code'] or 3918909000).font = Font(name='Arial', size=12)
        ws.cell(row=row, column=3, value=info['english_name'] or item.get('name_en', '')).font = Font(name='Arial', size=12)

        c4 = ws.cell(row=row, column=4, value=qty)
        c4.font = Font(name='Arial', size=12)
        c4.number_format = '0_);[Red]\\(0\\)'
        c4.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=5, value='PC(S)').font = Font(name='Arial', size=12)

        c6 = ws.cell(row=row, column=6, value=unit_usd)
        c6.font = Font(name='Arial', size=12)
        c6.number_format = '0.00_);[Red]\\(0.00\\)'
        c6.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=7, value='USD').font = Font(name='Arial', size=12)

        c8 = ws.cell(row=row, column=8, value=total_item_usd)
        c8.font = Font(name='Arial', size=12)
        c8.number_format = '0.00_);[Red]\\(0.00\\)'
        c8.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=9, value=info['material'] or 'plastic').font = Font(name='Arial', size=12)

        total_usd += total_item_usd
        total_qty += qty
        idx += 1
        row += 1

    # Empty placeholder row
    ws.cell(row=row, column=1, value=idx).font = Font(name='Arial', size=12)
    row += 1

    # Totals
    c4_tot = ws.cell(row=row, column=4, value=total_qty)
    c4_tot.font = Font(name='Arial', size=12)
    c4_tot.number_format = '#0'
    c4_tot.alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=row, column=5, value='/').font = Font(name='Arial', size=12)
    ws.cell(row=row, column=6, value='/').font = Font(name='Arial', size=12)
    ws.cell(row=row, column=7, value='/').font = Font(name='Arial', size=12)

    c8_tot = ws.cell(row=row, column=8, value=total_usd)
    c8_tot.font = Font(name='Arial', size=12)
    c8_tot.number_format = '0.00_);[Red]\\(0.00\\)'
    c8_tot.alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=row, column=9, value='/').font = Font(name='Arial', size=12)
    ws.cell(row=row, column=10, value='/').font = Font(name='Arial', size=12)


def _fill_pl(ws, items, kb, cno, tq, info_fn):
    """Fill the Packing List worksheet."""
    # Set column widths
    col_widths = {0: 4025/256, 1: 7097/256, 2: 6609/256, 4: 3257/256, 5: 3840/256,
                  6: 3072/256, 7: 3118/256, 8: 3140/256, 9: 5032/256}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx + 1)].width = width

    # Set row heights
    row_heights = {0: 30, 1: 30, 2: 61, 3: 30, 4: 30, 5: 67, 6: 30, 7: 30, 8: 77}
    for row_idx, height in row_heights.items():
        ws.row_dimensions[row_idx + 1].height = height
    for r in range(9, 13):
        ws.row_dimensions[r + 1].height = 68 + (r - 9) * 2
    ws.row_dimensions[14].height = 30
    ws.row_dimensions[15].height = 30
    ws.row_dimensions[16].height = 30

    # Title
    ws['A1'] = 'PACKING LIST'
    ws['A1'].font = Font(name='Arial', size=16, bold=False)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='top')
    ws.merge_cells('A1:I1')

    # Header - with merges
    ws['A2'] = 'Shipper:'
    ws['A2'].font = Font(name='Arial', size=12)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells('B2:C2')
    ws['B2'] = 'Shenzhen Adhoc Trading Co., Ltd.'
    ws['B2'].font = Font(name='Arial', size=10)
    ws['B2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D2:E2')
    ws['D2'] = 'INVOICE NO:'
    ws['D2'].font = Font(name='Arial', size=12)
    ws['D2'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws.merge_cells('F2:G2')
    ws['F2'] = cno
    ws['F2'].font = Font(name='Arial', size=12)
    ws['F2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws['A3'] = 'ADD:'
    ws['A3'].font = Font(name='Arial', size=12)
    ws['A3'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells('B3:C3')
    ws['B3'] = 'Flat 1006, Zhenye International Business Centre,No.3101-90,Qianhai Road, Nanshan District, Shenzhen,China.'
    ws['B3'].font = Font(name='Arial', size=10)
    ws['B3'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D3:E3')
    ws['D3'] = 'DATE:'
    ws['D3'].font = Font(name='Arial', size=12)
    ws['D3'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws.merge_cells('F3:G3')
    ws['F3'] = datetime.now().strftime('%Y/%m/%d')
    ws['F3'].font = Font(name='宋体', size=12)
    ws['F3'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.merge_cells('D4:E4')
    ws['D4'] = 'ORIGIN COUNTRY :'
    ws['D4'].font = Font(name='Arial', size=12)
    ws['D4'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws.merge_cells('F4:G4')
    ws['F4'] = 'CHINA'
    ws['F4'].font = Font(name='Arial', size=12)
    ws['F4'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws['A5'] = 'TO:'
    ws['A5'].font = Font(name='Arial', size=12)
    ws['A5'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.merge_cells('B5:C5')
    ws['B5'] = 'ZEATALINE INTERNATIONAL TRADING'
    ws['B5'].font = Font(name='Arial', size=10)
    ws['B5'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D5:E5')
    ws['D5'] = 'PRICE TERMS:'
    ws['D5'].font = Font(name='Arial', size=12)
    ws['D5'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws['F5'] = 'C&F'
    ws['F5'].font = Font(name='Arial', size=11)
    ws['F5'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws['G5'] = 'USA'
    ws['G5'].font = Font(name='宋体', size=10)
    ws['G5'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws['A6'] = 'ADD:'
    ws['A6'].font = Font(name='Arial', size=12)
    ws['A6'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.merge_cells('B6:C6')
    ws['B6'] = 'zouzhichang\n890 S Azusa Ave\nCity of Industry, CA 91748\nUS'
    ws['B6'].font = Font(name='Arial', size=10)
    ws['B6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('D6:E6')
    ws['D6'] = 'Reference No.'
    ws['D6'].font = Font(name='Arial', size=12)
    ws['D6'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)

    # Column headers - English
    hdrs = ['NO', 'Tariff Code', 'Descriptions', 'Qty', 'Unit',
            'Box Qty\n(CTNS)', 'N.W.\n(KG)', 'G.W.\n(KG)', 'VOLUME (CBM)', '']
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=7, column=c, value=h)
        cell.font = Font(name='Arial', size=12, bold=False)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Column headers - Chinese
    cn_hdrs = ['', '海关编码（清关用的）', '英文品名', '数量', 'PC(S)',
               '箱数', '净重', '毛重', '方数', '产品图片']
    for c, h in enumerate(cn_hdrs, 1):
        cell = ws.cell(row=8, column=c, value=h)
        cell.font = Font(name='宋体', size=12, bold=False)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    row = 9
    tot_qty = 0
    tot_boxes = 0
    tot_nw = 0.0
    tot_gw = 0.0
    tot_vol = 0.0
    idx = 1

    for item in items:
        sku = sku_key(item)
        qty = tq.get(sku, 0)
        if qty == 0:
            continue

        pr = item.get('packing_rate', 1) or 1
        boxes = qty / pr
        l, w, h_ = get_lwh(item)
        nw = item.get('net_weight_kg', 0) * boxes
        gw = item.get('gross_weight_kg', 0) * boxes
        volume = (l * w * h_ / 1_000_000) * boxes

        info = info_fn(sku, item)

        c1 = ws.cell(row=row, column=1, value=idx)
        c1.font = Font(name='Arial', size=10)
        c1.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=2, value=info['tariff_code'] or 3918909000).font = Font(name='Arial', size=12)
        ws.cell(row=row, column=3, value=info['english_name'] or item.get('name_en', '')).font = Font(name='Arial', size=12)

        c4 = ws.cell(row=row, column=4, value=qty)
        c4.font = Font(name='宋体', size=14)
        c4.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=row, column=5, value='PC(S)').font = Font(name='Arial', size=12)

        c6 = ws.cell(row=row, column=6, value=boxes)
        c6.font = Font(name='宋体', size=14)
        c6.alignment = Alignment(horizontal='center', vertical='center')

        c7 = ws.cell(row=row, column=7, value=round(nw, 1))
        c7.font = Font(name='Arial', size=12)
        c7.number_format = '0.00'
        c7.alignment = Alignment(horizontal='center', vertical='center')

        c8 = ws.cell(row=row, column=8, value=round(gw, 1))
        c8.font = Font(name='Arial', size=12)
        c8.number_format = '0.00'
        c8.alignment = Alignment(horizontal='center', vertical='center')

        c9 = ws.cell(row=row, column=9, value=round(volume, 5))
        c9.font = Font(name='Arial', size=12)
        c9.number_format = '0.00'
        c9.alignment = Alignment(horizontal='center', vertical='center')

        tot_qty += qty
        tot_boxes += boxes
        tot_nw += nw
        tot_gw += gw
        tot_vol += volume
        idx += 1
        row += 1

    # Empty placeholder
    ws.cell(row=row, column=1, value=idx).font = Font(name='Arial', size=12)
    ws.cell(row=row, column=5, value='PC(S)').font = Font(name='Arial', size=12)
    row += 1

    # Totals
    c4_tot = ws.cell(row=row, column=4, value=tot_qty)
    c4_tot.font = Font(name='Arial', size=12)
    c4_tot.number_format = '#0'
    c4_tot.alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=row, column=5, value='/').font = Font(name='Arial', size=12)

    c6_tot = ws.cell(row=row, column=6, value=round(tot_boxes))
    c6_tot.font = Font(name='Arial', size=12)
    c6_tot.number_format = '#0'
    c6_tot.alignment = Alignment(horizontal='center', vertical='center')

    c7_tot = ws.cell(row=row, column=7, value=round(tot_nw, 1))
    c7_tot.font = Font(name='Arial', size=12)
    c7_tot.number_format = '#0.00'
    c7_tot.alignment = Alignment(horizontal='center', vertical='center')

    c8_tot = ws.cell(row=row, column=8, value=round(tot_gw, 1))
    c8_tot.font = Font(name='Arial', size=12)
    c8_tot.number_format = '#0.00'
    c8_tot.alignment = Alignment(horizontal='center', vertical='center')

    c9_tot = ws.cell(row=row, column=9, value=round(tot_vol, 5))
    c9_tot.font = Font(name='Arial', size=12)
    c9_tot.number_format = '#0.00'
    c9_tot.alignment = Alignment(horizontal='center', vertical='center')
