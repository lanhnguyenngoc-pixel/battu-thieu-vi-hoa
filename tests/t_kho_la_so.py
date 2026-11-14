# -*- coding: utf-8 -*-
"""Kiểm thử kho lá số SQLite (YÊU CẦU 5.3): lưu → mở lại → tính lại cho khối TÓM TẮT trùng lúc lưu; tìm; nhóm; sửa; xóa.
Dùng: python -m tests.t_kho_la_so [số_lá]  (lá lấy từ tests/t8_chuan.json)"""
import sys, os, json, tempfile, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.kho_la_so import Kho, anh_chup
from battu_core.nhap_lieu import LaSo
from battu_core.dich_vu import tinh


def chay(n=12):
    ds = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 't8_chuan.json'), encoding='utf-8'))[:n]
    kho = Kho(os.path.join(tempfile.mkdtemp(), 'kt.sqlite')); t0 = time.time()
    ids = []; dung = 0; loi = []
    for i, x in enumerate(ds):
        s = LaSo(ho_ten=x['id'], ngay=x['ngay'], thang=x['thang'], nam=x['nam'], gio=x['gio'], phut=x['phut'],
                 gioi_tinh=x['gioi_tinh'], nam_xem=2026, khong_ro_gio=(i % 5 == 4))
        kq = tinh(s); ids.append(kho.luu(kq, x['id'], 'Nhóm %d' % (i % 3), 'ghi chú %d' % i))
    for i, id_ in enumerate(ids):
        r = kho.lay(id_); kq = tinh(kho.la_so(id_))
        if json.loads(r['anh_chup']) == anh_chup(kq) and kho.cung_phien_ban(id_): dung += 1
        else: loi.append(r['ten'])
    kiem = {
        'tìm theo tên': len(kho.tim(ds[0]['id'])) >= 1,
        'lọc nhóm': len(kho.tim('', 'Nhóm 1')) == len([i for i in range(len(ds)) if i % 3 == 1]),
        'danh sách nhóm': kho.nhom() == ['Nhóm 0', 'Nhóm 1', 'Nhóm 2'],
    }
    kho.sua_thong_tin(ids[0], 'Tên mới', 'Nhóm X', 'ghi chú mới'); kiem['sửa thông tin'] = kho.lay(ids[0])['ten'] == 'Tên mới' and len(kho.tim('ghi chú mới')) == 1
    kho.xoa(ids[1]); kiem['xóa'] = kho.lay(ids[1]) is None and len(kho.tim()) == len(ds) - 1
    # sao lưu / khôi phục (1.4.0): sao lưu → kho mới trống khôi phục đủ; khôi phục lần hai bỏ qua hết vì trùng
    tm = tempfile.mkdtemp(); p_sl = os.path.join(tm, 'sl.sqlite')
    n_sl = kho.sao_luu(p_sl); kho2 = Kho(os.path.join(tm, 'moi.sqlite'))
    a = kho2.khoi_phuc(p_sl); b = kho2.khoi_phuc(p_sl)
    giong = sorted((r['ten'], r['nhap'], r['anh_chup']) for r in kho2.tim()) == sorted((r['ten'], r['nhap'], r['anh_chup']) for r in kho.tim())
    kiem['sao lưu'] = n_sl == len(ds) - 1
    kiem['khôi phục'] = a == (len(ds) - 1, 0) and b == (0, len(ds) - 1) and giong
    # thống kê đối chiếu (1.4.x): tổng = cộng các ô THEO_DOI E4 (số năm), E5×E4, E6×E4 (số năm khớp PP1, PP2) của từng lá
    from battu_core.thong_ke import thong_ke
    kho3 = Kho(os.path.join(tm, 'tk.sqlite')); mong = [0, 0.0, 0.0]
    for i, x in enumerate(ds[:3]):
        s = LaSo(ho_ten=x['id'], ngay=x['ngay'], thang=x['thang'], nam=x['nam'], gio=x['gio'], phut=x['phut'], gioi_tinh=x['gioi_tinh'], nam_xem=2026,
                 theo_doi={x['nam'] + 20 + j: [['Tốt', 'Xấu', 'Bình thường'][(i + j) % 3], None, ['Đúng', 'Sai', 'Một phần', 'Chưa rõ'][j % 4]] for j in range(6 + i)})
        kq = tinh(s); kho3.luu(kq, x['id']); E = kq.E
        e4 = E.gia_tri('THEO_DOI', 'E4'); mong[0] += e4; mong[1] += E.gia_tri('THEO_DOI', 'E5') * e4; mong[2] += E.gia_tri('THEO_DOI', 'E6') * e4
    kho3.luu(tinh(LaSo(ngay=1, thang=1, nam=1990, gio=1, phut=0)), 'không ghi')
    tk = thong_ke(kho3)
    kiem['thống kê đối chiếu'] = (tk['so_la'], tk['so_nam'], round(tk['ty_le_pp1'] * tk['so_nam']), round(tk['ty_le_pp2'] * tk['so_nam'])) == \
        (3, mong[0], round(mong[1]), round(mong[2])) and sum(tk['danh_gia'].values()) == 6 + 7 + 8
    try: kho2.khoi_phuc(os.path.join(os.path.dirname(os.path.abspath(__file__)), 't8_chuan.json')); kiem['tệp sai bị từ chối'] = False
    except Exception: kiem['tệp sai bị từ chối'] = len(kho2.tim()) == len(ds) - 1
    return len(ids), dung, loi, kiem, time.time() - t0


if __name__ == '__main__':
    n, dung, loi, kiem, dt = chay(int(sys.argv[1]) if len(sys.argv) > 1 else 12)
    print('Kho lá số: mở lại %d/%d lá cho TÓM TẮT trùng lúc lưu (%.1f giây)' % (dung, n, dt), loi)
    for k, v in kiem.items(): print('  %-16s %s' % (k, 'ĐẠT' if v else 'LỖI'))
