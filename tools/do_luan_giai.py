"""Bước 4 – Đo tự động: so kết luận phần mềm (KET_LUAN) với lời sách TVH theo từng lĩnh vực."""
import sys, json, re, collections
sys.path.insert(0, '/home/claude/battu'); sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from battu_core.nhap_lieu import LaSo
from battu_core.dich_vu import tinh
from battu_core.engine import Engine
CAN = ['Giáp','Ất','Bính','Đinh','Mậu','Kỷ','Canh','Tân','Nhâm','Quý']
CHI = ['Tí','Sửu','Dần','Mão','Thìn','Tị','Ngọ','Mùi','Thân','Dậu','Tuất','Hợi']
def nam_tu_tru(can, chi):
    for y in range(1900, 2000):
        if CAN[(y-4) % 10] == can and CHI[(y-4) % 12] == chi: return y
def muc3(s):
    s = s.upper()
    if s.startswith(('BẤT LỢI NHƯNG','THUẬN NHƯNG','TRUNG BÌNH')): return 'V'
    if s.startswith(('THUẬN','NGHIÊNG THUẬN','TỐT','KHÁ','GIÀU','RẤT TỐT')): return 'T'
    if s.startswith(('BẤT LỢI','NGHIÊNG BẤT LỢI','KÉM','XẤU','NGHÈO','VẤT VẢ','RẤT KÉM')): return 'K'
    return 'V'
def lay(E):
    g = lambda r: str(E.gia_tri('KET_LUAN', 'C%d' % r) or '')
    c6, c7, c8, c9, c45 = g(6), g(7), g(8), g(9), g(45)
    o = {}
    m = re.search(r'MỨC PHÚ QUÝ TỔNG HỢP: ([^(]+)\(', c6); tl = m.group(1).strip() if m else ''
    o['TL_raw'] = tl
    o['TL'] = 'T' if tl.startswith('GIÀU') else ('K' if tl.startswith(('VẤT VẢ', 'NGHÈO', 'KHÓ')) else 'V')
    def tt(key, txt):
        m = re.search(key + r':\s*([A-ZÀ-Ỹ ]+?)(?:[;–.(]| -|$)', txt); return m.group(1).strip() if m else ''
    o['SN_raw'] = tt('QUAN VẬN', c7); o['SN'] = muc3(o['SN_raw'])
    o['HOC_raw'] = tt('HỌC TẬP – THI CỬ', c7); o['DANH_raw'] = tt('DANH TIẾNG', c7)
    m = re.search(r'Ông bà – tổ nghiệp: ([^;]+); cha: ([^;]+); mẹ: ([^;]+); anh chị em: ([^;]+); con cái: ([^.]+)\.', c8)
    if m:
        for k, v in zip(['TN', 'CHA', 'ME', 'AE', 'CC'], m.groups()): o[k + '_raw'] = v.strip(); o[k] = muc3(v.strip())
    m = re.search(r'Hôn nhân – mức ([A-ZÀ-Ỹ ]+)\.', c8); o['HN_raw'] = m.group(1).strip() if m else ''; o['HN'] = muc3(o['HN_raw'])
    o['BENH_raw'] = tt('BỆNH TẬT', c9); o['TAI_raw'] = tt('THƯƠNG TẬT – TAI HỌA', c9); o['THO_raw'] = tt('Thọ – yểu \\(P1 tr.173\\)', c9)
    o['CATHUNG_raw'] = tt('Cát – hung', c45)
    return o
def chon(o, lv, cau):
    """Chọn ô của công cụ tương ứng lĩnh vực của câu sách."""
    if lv == 'TL': return o['TL'], o['TL_raw']
    if lv == 'SN': return o['SN'], o['SN_raw']
    if lv == 'HN': return o['HN'], o['HN_raw']
    if lv == 'CC': return o.get('CC'), o.get('CC_raw')
    if lv == 'CD':
        r = o['HOC_raw'] if re.search(r'thi|học', cau) else o['DANH_raw']; return muc3(r), r
    if lv == 'LT':
        if cau.startswith('[Tổ nghiệp]'): return o.get('TN'), o.get('TN_raw')
        if cau.startswith('[Anh em]'): return o.get('AE'), o.get('AE_raw')
        c, m = 'cha' in cau.lower(), 'mẹ' in cau.lower() or 'ấn' in cau.lower()
        if c and not m: return o.get('CHA'), o.get('CHA_raw')
        if m and not c: return o.get('ME'), o.get('ME_raw')
        xs = [o.get('CHA'), o.get('ME')]; bad = 'K' if 'K' in xs else ('V' if 'V' in xs else 'T')
        return bad, 'cha %s / mẹ %s' % (o.get('CHA_raw'), o.get('ME_raw'))
    if lv == 'SK':
        r = o['TAI_raw'] if re.search(r'thương|tai|tàn|đầu|nạn', cau) else o['BENH_raw']; return muc3(r), r
    if lv == 'TH': return muc3(o['TAI_raw']), o['TAI_raw']
    if lv == 'TC': return muc3(o['CATHUNG_raw']), o['CATHUNG_raw']
    return None, None
def engine_tu_tru(tt, gioi):
    p = [x.strip().split() for x in tt.split(',')]
    can = [a for a, b in p]; chi = [b for a, b in p]
    s = LaSo(cach_nhap='canchi', can=can, chi=chi, gioi_tinh=gioi, nam_sinh_tru_nam=nam_tu_tru(can[0], chi[0]), nam_xem=2026)
    return tinh(s).E
def tat_ca():
    from gd_a_tong_hop_data import G   # lá có lời luận nhiều lĩnh vực
    from dap_an_chuong import VD
    rows = []
    cache = {}
    def E_of(key, ten=None, tt=None, gioi='Nam'):
        if key in cache: return cache[key]
        try:
            E = Engine({'NHAP_LIEU': {'C13': ten, 'C12': 'Không'}}) if ten else engine_tu_tru(tt, gioi)
            cache[key] = lay(E)
        except Exception as e:
            cache[key] = None
        return cache[key]
    import openpyxl
    ws = openpyxl.load_workbook('/home/claude/HO_SO_PM/HO_SO_PM/03_NGUON_CHUAN/BAT_TU_THIEU_VI_HOA_6.4.xlsx', read_only=True)['KIEM_THU']
    TT = {}
    for row in ws.iter_rows(min_row=14, max_row=77, values_only=True):
        if row[1]: TT[row[1]] = ','.join('%s %s' % (row[2+i], row[6+i]) for i in range(4))
    TT['P2-416 VD24 (chưa có trong KIEM_THU)'] = None
    for k, v in G.items():
        tt = v.get('tu_tru') or TT.get(k.split(' =')[0])
        if not tt: continue
        o = E_of(k, tt=tt, gioi=v.get('gioi', 'Nam'))
        for lv in ['TC', 'SN', 'TL', 'HN', 'CC', 'LT', 'SK']:
            if lv in v and o:
                cv, raw = chon(o, lv, v[lv][1])
                rows.append(dict(la=k, nguon='VD dụng thần/Ch.23', lv=lv, sach=v[lv][0], cong_cu=cv, cong_cu_raw=raw, cau=v[lv][1][:120], trang=v[lv][2]))
    for (ma, gioi, tr, tt, lv, nh, cau, sv) in VD:
        if '?' in tt or nh == '-' or lv == 'TU': continue
        o = E_of(tt + gioi, tt=tt, gioi=gioi)
        if not o: continue
        cv, raw = chon(o, lv, cau)
        rows.append(dict(la=ma, nguon='Ch.12–21', lv=lv, sach=nh, cong_cu=cv, cong_cu_raw=raw, cau=cau[:120], trang=str(tr), tu_tru=tt))
    return rows
if __name__ == '__main__':
    rows = tat_ca()
    json.dump(rows, open('/home/claude/battu/tai_lieu/giai_doan_A/do_hien_tai.json', 'w'), ensure_ascii=False, indent=1)
    TEN = {'TC':'Tổng thể','SN':'Quan vận','TL':'Tài lộc','HN':'Hôn nhân','CC':'Con cái','LT':'Cha mẹ – anh em – tổ nghiệp','SK':'Sức khỏe – tai nạn','CD':'Học hành – danh tiếng','TH':'Tù tội'}
    tong = collections.defaultdict(lambda: [0, 0, 0])
    for r in rows:
        t = tong[r['lv']]; t[0] += 1; t[1] += r['sach'] == r['cong_cu']
        t[2] += r['cong_cu'] is not None and abs('KVT'.index(r['sach']) - 'KVT'.index(r['cong_cu'])) <= 1
    print('%-30s %5s %8s %10s' % ('Lĩnh vực', 'Số', 'Khớp', 'Lệch≤1'))
    for lv, (n, a, b) in sorted(tong.items(), key=lambda x: -x[1][0]):
        print('%-30s %5d %5d (%2d%%) %4d' % (TEN[lv], n, a, 100*a//n, b))
    print('cache lỗi:', sum(1 for r in rows if r['cong_cu'] is None))
    print('Từ vựng HN:', collections.Counter(r['cong_cu_raw'] for r in rows if r['lv']=='HN'))
    print('Từ vựng LT:', collections.Counter(r['cong_cu_raw'] for r in rows if r['lv']=='LT').most_common(8))
