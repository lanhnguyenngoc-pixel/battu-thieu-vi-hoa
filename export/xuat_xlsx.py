# -*- coding: utf-8 -*-
"""Xuất XLSX (YÊU CẦU 5.4): bố cục bảng như workbook (không cần công thức).
Giữ đúng địa chỉ ô, ô gộp, độ rộng cột, cột/dòng ẩn, định dạng số của ô; định dạng có điều kiện được tô sẵn theo kết quả
của lá số. Bỏ 3 sheet chỉ tính ẩn (MAU_THUAN, KIEM_THU, CSDL); vùng phụ trợ KET_LUAN 200–251 để ẩn như workbook.
Sheet đầu "THONG_TIN": dữ liệu nhập, phiên bản quy tắc, ngày xuất."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from battu_core import xl as X
from battu_core.du_lieu import model, ref_of
from battu_core.dinh_dang_dk import kieu_sheet
from export.noi_dung import SHEET_AN, DONG_PHU_TRO_KL, TEN_PHAN_MEM, thong_tin_la, chu_ky, ngay_in, la_tieu_de

NEN_TIEU_DE = PatternFill('solid', fgColor='F3F0FF')


def _gia_tri(v):
    if isinstance(v, X.ErrVal): return v.code
    return v


def xuat_xlsx(kq, duong_dan):
    E = kq.E; M = model(); wb = Workbook(); ws0 = wb.active; ws0.title = 'THONG_TIN'
    ws0['A1'] = TEN_PHAN_MEM + ' – kết quả lá số'; ws0['A1'].font = Font(bold=True, size=14)
    dong = 3
    for k, v in thong_tin_la(kq) + [('Phiên bản quy tắc', chu_ky()), ('Ngày xuất', ngay_in()),
                                     ('Ghi chú', 'Bản xuất không chứa công thức; địa chỉ ô trùng workbook 6.3. Nội dung sức khỏe chỉ tham khảo, không phải chẩn đoán.')]:
        ws0.cell(dong, 1, k).font = Font(bold=True); ws0.cell(dong, 2, v); dong += 1
    ws0.column_dimensions['A'].width = 34; ws0.column_dimensions['B'].width = 110
    for sh in M.sheets:
        if sh in SHEET_AN: continue
        si = M.si[sh]; ws = wb.create_sheet(sh); fc = M.formula_cells[si]; dd = M.dinh_dang.get(sh, {})
        dk = kieu_sheet(E, sh); bc = M.bo_cuc.get(sh, {})
        o = set(E.consts[si]) | set(fc) | {(r, c) for (s_, r, c) in E.dt if s_ == si}
        for (r, c) in o:
            if (r, c) in fc or (si, r, c) in E.dt:
                v = _gia_tri(E.raw(si, r, c))
            else:
                v = E.consts[si].get((r, c))
            if v is None or v == '': continue
            x = ws.cell(r, c, v)
            ma = dd.get(ref_of(r, c))
            if ma and ma != 'General' and isinstance(v, (int, float)) and not isinstance(v, bool): x.number_format = ma
            if isinstance(v, str) and len(v) > 30: x.alignment = Alignment(wrap_text=True, vertical='top')
            if (r, c) not in fc and isinstance(v, str) and la_tieu_de(v):
                x.font = Font(bold=True, color='4338CA'); x.fill = NEN_TIEU_DE
            k = dk.get((r, c))
            if k:
                x.font = Font(bold=bool(k.get('dam')) or x.font.bold, color=(k.get('chu') or '#000000')[1:].upper())
                if k.get('nen'): x.fill = PatternFill('solid', fgColor=k['nen'][1:].upper())
        for g in bc.get('gop', []): ws.merge_cells(g)
        for c, w in bc.get('rong', {}).items(): ws.column_dimensions[get_column_letter(int(c))].width = float(w)
        for c in bc.get('an_cot', []): ws.column_dimensions[get_column_letter(int(c))].hidden = True
        an_d = set(bc.get('an_dong', [])) | (set(DONG_PHU_TRO_KL) if sh == 'KET_LUAN' else set())
        for r in an_d: ws.row_dimensions[r].hidden = True
    wb.save(duong_dan)
    return duong_dan
