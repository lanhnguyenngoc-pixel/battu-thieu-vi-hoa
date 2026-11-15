"""Giai đoạn B – trích dấu hiệu (quy tắc đủ điều kiện) theo lĩnh vực cho từng lá trong bộ đáp án."""
import sys, json, re, pickle, os
sys.path.insert(0, '/home/claude/battu'); sys.path.insert(0, '/home/claude/battu/tools')
from do_luan_giai import engine_tu_tru, lay
SHEET = {'HN': ('HON_NHAN', 70, 260), 'SN': ('SU_NGHIEP', 1, 400), 'TL': ('SU_NGHIEP', 1, 400)}
def quy_tac(E, sheet, a, b):
    kn = E.gia_tri('DOI_CHIEU', 'C33'); col = 'F' if kn == 2 else 'E'
    out = {}
    for r in range(a, b):
        ma = E.gia_tri(sheet, 'B%d' % r)
        if isinstance(ma, str) and re.fullmatch(r'[A-Z]{1,4}\d+[a-z]?', ma.strip()):
            out[ma.strip()] = (str(E.gia_tri(sheet, '%s%d' % (col, r))).strip().upper().startswith('ĐỦ'),
                               str(E.gia_tri(sheet, 'D%d' % r))[:120], str(E.gia_tri(sheet, 'G%d' % r))[:160], r)
    return out
def vai_tro(E):
    col = 'D' if E.gia_tri('DOI_CHIEU', 'C33') == 2 else 'C'
    nhom = lambda v: str(v).split('(')[-1].rstrip(')') if v else ''
    return dict(dung=nhom(E.gia_tri('DUNG_HY_KY', col + '106')), hy=nhom(E.gia_tri('DUNG_HY_KY', col + '107')),
                ky=nhom(E.gia_tri('DUNG_HY_KY', col + '108')), trang_thai=str(E.gia_tri('DUNG_HY_KY', col + '6')),
                luong1={str(E.gia_tri('DUNG_HY_KY', 'B%d' % r)): E.gia_tri('DUNG_HY_KY', 'C%d' % r) for r in range(12, 17)},
                luong2={str(E.gia_tri('DUNG_HY_KY', 'B%d' % r)): E.gia_tri('DUNG_HY_KY', 'D%d' % r) for r in range(12, 17)})
def chay(rows_json='/home/claude/battu/tai_lieu/giai_doan_A/do_hien_tai.json', lv_list=('HN',)):
    lv_list = tuple(lv_list)
    R = json.load(open(rows_json))
    cache = '/home/claude/battu/tai_lieu/giai_doan_A/dac_trung_cache.pkl'
    C = pickle.load(open(cache, 'rb')) if os.path.exists(cache) else {}
    G = json.load(open('/home/claude/battu/tai_lieu/giai_doan_A/dap_an_tong.json'))
    import openpyxl
    ws = openpyxl.load_workbook('/home/claude/HO_SO_PM/HO_SO_PM/03_NGUON_CHUAN/BAT_TU_THIEU_VI_HOA_6.4.xlsx', read_only=True)['KIEM_THU']
    TT = {}
    for row in ws.iter_rows(min_row=14, max_row=77, values_only=True):
        if row[1]: TT[row[1]] = ','.join('%s %s' % (row[2+i], row[6+i]) for i in range(4))
    out = []
    for r in R:
        if r['lv'] not in lv_list: continue
        tt = r.get('tu_tru') or (G.get(r['la'], {}).get('tu_tru')) or TT.get(r['la'].split(' =')[0])
        gioi = G[r['la']]['gioi'] if r['la'] in G else None
        if gioi is None:
            import dap_an_chuong; gioi = next(v[1] for v in dap_an_chuong.VD if v[0] == r['la'])
        key = (tt, gioi)
        if key not in C or '_role' not in C[key]:
            E = engine_tu_tru(tt, gioi)
            C[key] = {lv: quy_tac(E, *SHEET[lv]) for lv in SHEET}
            C[key]['_o'] = lay(E)
            C[key]['_than'] = str(E.gia_tri('DOI_CHIEU', 'C34'))
            C[key]['_role'] = vai_tro(E)
        out.append(dict(r, tu_tru=tt, gioi=gioi, qt=C[key].get(r['lv'], {}), o=C[key]['_o'], than=C[key]['_than'], role=C[key]['_role']))
    pickle.dump(C, open(cache, 'wb'))
    return out
if __name__ == '__main__':
    sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
    X = chay(lv_list=tuple(sys.argv[1:]) or ('HN',))
    print(len(X), 'dòng'); print(list(X[0]['qt'].items())[:3])
