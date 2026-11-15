"""Phân loại nguyên nhân lệch: do khâu trước (vượng suy / dụng thần khác sách), do vận/biến cố, hay do quy trình lĩnh vực."""
import sys, collections, openpyxl
sys.path.insert(0, '/home/claude/battu/tools')
from tt_co_ban import *
from gd_b_mohinh import co_ban
HANH = ['Mộc','Hỏa','Thổ','Kim','Thủy']; OFF = {'Tỉ kiếp':0,'Thực thương':1,'Tài':2,'Quan sát':3,'Ấn':4}
ws = openpyxl.load_workbook('/home/claude/HO_SO_PM/HO_SO_PM/03_NGUON_CHUAN/BAT_TU_THIEU_VI_HOA_6.4.xlsx', read_only=True)['KIEM_THU']
SACH = {}
for row in ws.iter_rows(min_row=14, max_row=77, values_only=True):
    if row[1]: SACH[row[1]] = (str(row[10] or ''), str(row[11] or ''))
def hanh_dung(r):
    nc = phan_tich(r['tu_tru'])['nc']; g = r['role']['dung']
    return HANH[(HANH_CAN[CAN.index(nc)] + OFF[g]) % 5] if g in OFF else '?'
def phan_loai(r, du_doan):
    ten = r['la'].split(' =')[0]; ly = []
    if ten in SACH:
        st, dg = SACH[ten]; bv = 'Nhược' if st.startswith('Nhược') else 'Vượng' if st.startswith(('Vượng', 'Cách')) else ''
        cv = 'Nhược' if 'NHƯỢC' in r['than'].upper() else 'Vượng'
        if bv and bv != cv: ly.append('vượng suy khác sách (sách %s, công cụ %s)' % (bv, cv))
        hd = hanh_dung(r)
        if dg and dg[:2] not in ('—', 'Kh') and hd not in dg: ly.append('dụng thần khác sách (sách %s, công cụ %s)' % (dg.split(' ')[0], hd))
    if r['nguon'] == 'Ch.12–21' and any(k in r['cau'] for k in ('vận', 'năm ', 'lưu niên')): ly.append('lời sách nói về vận/năm cụ thể')
    return ly or ['quy trình lĩnh vực chưa đúng']
