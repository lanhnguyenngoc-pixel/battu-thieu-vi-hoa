# -*- coding: utf-8 -*-
"""Thống kê đối chiếu thực tế trên cả kho lá số (bản 1.4.x, đề xuất 3.8 – đã duyệt 14/11/2026).

Không thêm quy tắc: với mỗi lá đã ghi "thực tế năm đó" (THEO_DOI cột N), tính lại lá số rồi CỘNG các ô kết quả
có sẵn của workbook – Q (khớp xu thế PP1), R (khớp xu thế PP2), P (đánh giá luận giải) – trên toàn kho."""
import collections
from battu_core.dich_vu import tinh
from export.lam_sach import lam_sach

DANH_GIA = ['Đúng', 'Một phần', 'Sai', 'Chưa rõ']


def thong_ke(kho, tien_do=None):
    """→ dict: tổng hợp + danh sách từng lá. tien_do(i, n) được gọi sau mỗi lá (để giao diện báo tiến độ)."""
    rows = [r for r in kho.tim() if '"theo_doi": {}' not in r['nhap'] and '"theo_doi"' in r['nhap']]
    tong = collections.Counter(); dg = collections.Counter(); tung_la = []
    for i, r in enumerate(rows):
        s = kho.la_so(r['id'])
        if not s.theo_doi:
            continue
        E = tinh(s).E
        nam = q = rr = 0; d = collections.Counter()
        for row in range(12, 112):                       # như THEO_DOI E4:E6 (cột N, Q, R) và E8 (cột P)
            p = E.gia_tri('THEO_DOI', 'P%d' % row)
            if p in DANH_GIA: d[p] += 1
            if E.gia_tri('THEO_DOI', 'N%d' % row) in (None, ''): continue
            nam += 1
            q += E.gia_tri('THEO_DOI', 'Q%d' % row) == 1
            rr += E.gia_tri('THEO_DOI', 'R%d' % row) == 1
        if nam == 0 and not d:
            continue
        tung_la.append({'id': r['id'], 'ten': r['ten'], 'tu_tru': r['tu_tru'], 'so_nam': nam, 'khop_pp1': q, 'khop_pp2': rr,
                        'danh_gia': dict(d), 'goi_y': lam_sach(E.gia_tri('THEO_DOI', 'E7') or ''), 'dang_dung': lam_sach(E.gia_tri('DOI_CHIEU', 'D33') or '')})
        tong['la'] += 1; tong['nam'] += nam; tong['pp1'] += q; tong['pp2'] += rr; dg.update(d)
        if tien_do: tien_do(i + 1, len(rows))
    n_dg = sum(dg.values())
    ty_le_dung = (dg['Đúng'] + 0.5 * dg['Một phần']) / (n_dg - dg['Chưa rõ']) if n_dg - dg['Chưa rõ'] > 0 else None
    return {'so_la': tong['la'], 'so_nam': tong['nam'],
            'ty_le_pp1': tong['pp1'] / tong['nam'] if tong['nam'] else None,
            'ty_le_pp2': tong['pp2'] / tong['nam'] if tong['nam'] else None,
            'danh_gia': {k: dg[k] for k in DANH_GIA}, 'ty_le_dung': ty_le_dung, 'tung_la': tung_la}


def xuat_xlsx(kq, duong_dan):
    """Ghi kết quả thống kê ra tệp Excel (bảng tổng + từng lá)."""
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Tổng hợp'
    pt = lambda x: '' if x is None else '%.0f%%' % (100 * x)
    for row in [['THỐNG KÊ ĐỐI CHIẾU THỰC TẾ TRÊN KHO LÁ SỐ'], [],
                ['Số lá có ghi thực tế', kq['so_la']], ['Số năm đã ghi', kq['so_nam']],
                ['Tỉ lệ khớp xu thế PP1', pt(kq['ty_le_pp1'])], ['Tỉ lệ khớp xu thế PP2', pt(kq['ty_le_pp2'])]] + \
               [['Đánh giá luận giải – ' + k, v] for k, v in kq['danh_gia'].items()] + \
               [['Tỉ lệ đúng (một phần tính ½)', pt(kq['ty_le_dung'])]]:
        ws.append(row)
    ws['A1'].font = Font(bold=True, size=13); ws.column_dimensions['A'].width = 40
    w2 = wb.create_sheet('Từng lá')
    w2.append(['Tên', 'Tứ trụ', 'Số năm', 'Khớp PP1', 'Khớp PP2'] + DANH_GIA + ['Đang dùng', 'Gợi ý phương pháp'])
    for x in kq['tung_la']:
        w2.append([x['ten'], x['tu_tru'], x['so_nam'], x['khop_pp1'], x['khop_pp2']] + [x['danh_gia'].get(k, 0) for k in DANH_GIA] +
                  [x['dang_dung'], x['goi_y']])
    for c, w in zip('ABCDEFGHIJK', (24, 34, 9, 10, 10, 8, 10, 8, 9, 22, 70)): w2.column_dimensions[c].width = w
    for c in w2[1]: c.font = Font(bold=True)
    wb.save(duong_dan)
