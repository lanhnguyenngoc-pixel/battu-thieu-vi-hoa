# -*- coding: utf-8 -*-
"""T15 – "Nâng cao bằng AI" (đề xuất hoàn thiện 15/11/2026, giai đoạn T5). KHÔNG gọi mạng thật: urlopen được giả lập.
  (a) Riêng tư: lời nhắc gửi đi không có họ tên, ngày – tháng sinh, giờ sinh.
  (b) Yêu cầu HTTP: đúng địa chỉ, header x-api-key / anthropic-version, có system, đúng mô hình.
  (c) Phản hồi đúng → qua bộ kiểm tra, lưu đệm; gọi lần hai dùng đệm, không gọi mạng.
  (d) Phản hồi sai (đổi dụng thần, đổi mức năm, bỏ năm, sai năm đại vận) → bị chặn, nêu đủ lỗi.
  (e) Lỗi HTTP 401 / mất mạng / bị cắt (max_tokens) → LoiAI có thông báo tiếng Việt.
  (f) Lưu khóa API (dự phòng tệp 0600 trên hệ không có kho khóa), biến môi trường ưu tiên, xóa khóa.
  (g) Giao diện: trang Luận giải tổng hợp nhận bản AI, bật "Bản AI", xuất PDF bản AI.
Dùng: python -m tests.t15_ai"""
import sys, os, json, io, tempfile, time, urllib.error
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))


class _Res(io.BytesIO):
    def __enter__(self): return self
    def __exit__(self, *a): self.close()


def _mo(chu, ghi, stop='end_turn'):
    def mo(req, timeout=None):
        ghi.append(req)
        return _Res(json.dumps({'content': [{'type': 'text', 'text': chu}], 'stop_reason': stop}).encode('utf-8'))
    return mo


def chay():
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    nha = tempfile.mkdtemp(); os.environ['HOME'] = nha; os.environ['USERPROFILE'] = nha; os.environ.pop('ANTHROPIC_API_KEY', None)
    from battu_core.nhap_lieu import LaSo
    from battu_core.dich_vu import tinh
    from battu_core.tong_hop.viet_ban import ban_luan_giai
    from battu_core.tong_hop import ai_claude as A
    loi = []; t0 = time.time()
    s = LaSo(ho_ten='Nguyễn Ngọc Lanh', cach_nhap='duong', ngay=28, thang=11, nam=1978, gio=8, phut=30, gioi_tinh='Nam', nam_xem=2026)
    kq = tinh(s); B, hs, pd = ban_luan_giai(kq)
    ln = A.loi_nhac(hs, pd, B)
    for x in ('Nguyễn', 'Ngọc Lanh', '28/11', '08:30', 'ngày 28'):
        if x in ln: loi.append('(a) lời nhắc chứa "%s"' % x)
    dem = tempfile.mkdtemp(); md = A._md(B); ghi = []
    BA, l, tu_dem = A.nang_cao('sk-ant-thu', hs, pd, B, 'T', 'claude-sonnet-5-5', thu_muc=dem, _mo=_mo(md, ghi))
    if not ghi: loi.append('(b) không gửi yêu cầu')
    else:
        r = ghi[0]; body = json.loads(r.data.decode('utf-8')); hd = {k.lower(): v for k, v in r.header_items()}
        if r.full_url != A.API_URL or hd.get('x-api-key') != 'sk-ant-thu' or hd.get('anthropic-version') != A.API_VERSION:
            loi.append('(b) sai địa chỉ / header')
        if body.get('model') != 'claude-sonnet-5-5' or not body.get('system') or body['messages'][0]['role'] != 'user':
            loi.append('(b) sai thân yêu cầu')
    if BA is None or l: loi.append('(c) phản hồi đúng bị chặn: %s' % l)
    if tu_dem: loi.append('(c) lần đầu không phải từ đệm')
    ghi2 = []
    BA2, l2, tu_dem2 = A.nang_cao('sk-ant-thu', hs, pd, B, 'T', 'claude-sonnet-5-5', thu_muc=dem, _mo=_mo('x', ghi2))
    if ghi2 or not tu_dem2 or BA2 is None: loi.append('(c) lần hai phải dùng đệm, không gọi mạng')
    sai = (md.replace('| 2032 Nhâm Tí (55) | **Rất xấu**', '| 2032 Nhâm Tí (55) | **Tốt**')
             .replace('Dụng thần Hỏa', 'Dụng thần Thủy', 1).replace('Mậu Thìn (2021–2030)', 'Mậu Thìn (2020–2029)'))
    sai = '\n'.join(d for d in sai.splitlines() if not d.startswith('| 2036 Bính Thìn'))
    BA3, l3, _ = A.nang_cao('k', hs, pd, B, 'T', 'claude-opus-5-5', thu_muc=tempfile.mkdtemp(), _mo=_mo(sai, []))
    can = ['Dụng thần', 'Mức năm 2032', 'thiếu năm 2036', 'Năm của đại vận Mậu Thìn']
    if BA3 is not None or not all(any(c in x for x in l3) for c in can): loi.append('(d) chưa chặn đủ: %s' % l3)

    def mo_loi(e):
        def mo(req, timeout=None): raise e
        return mo
    for e, can in ((urllib.error.HTTPError(A.API_URL, 401, 'x', {}, io.BytesIO(b'{}')), 'Khóa API không hợp lệ'),
                   (urllib.error.URLError('no route'), 'Không kết nối được')):
        try:
            A.goi_api('k', hs, pd, B, _mo=mo_loi(e)); loi.append('(e) không báo lỗi %s' % can)
        except A.LoiAI as x:
            if can not in str(x): loi.append('(e) thông báo sai: %s' % x)
    try:
        A.goi_api('k', hs, pd, B, _mo=_mo(md, [], stop='max_tokens')); loi.append('(e) không báo bị cắt')
    except A.LoiAI as x:
        if 'cắt' not in str(x): loi.append('(e) thông báo cắt sai: %s' % x)
    try: A.goi_api('', hs, pd, B); loi.append('(e) thiếu khóa không báo')
    except A.LoiAI: pass

    from ui import khoa_api as K
    if sys.platform not in ('win32', 'darwin'):
        K.ghi('sk-ant-abcdefghijkl1234')
        p = K._tep(); quyen = oct(os.stat(p).st_mode & 0o777)
        if K.doc() != 'sk-ant-abcdefghijkl1234' or quyen != '0o600': loi.append('(f) tệp khóa: %s %s' % (K.doc(), quyen))
    os.environ['ANTHROPIC_API_KEY'] = 'sk-ant-moitruong'
    if K.doc() != 'sk-ant-moitruong': loi.append('(f) biến môi trường không được ưu tiên')
    os.environ.pop('ANTHROPIC_API_KEY')
    if sys.platform not in ('win32', 'darwin'):
        K.xoa()
        if K.doc(): loi.append('(f) xóa khóa không sạch')
    if K.che('sk-ant-abcdefghijkl1234') != 'sk-ant-…1234': loi.append('(f) che khóa sai')

    from ui.giao_dien import tao_app
    app = tao_app([])
    from ui.trang_tong_hop import TrangTongHop
    w = TrangTongHop(); w.hien(kq)
    w._ai_xong(('ok', BA, [], False))
    if not w.b1.isEnabled() or w._dang_xem() is not BA: loi.append('(g) giao diện không chuyển sang bản AI')
    w._ai_xong(('ok', None, ['Mức năm 2032 sai'], False))
    if 'KHÔNG dùng' not in w.tt.text(): loi.append('(g) không báo bản AI bị chặn')
    from export.xuat_luan_giai import xuat_pdf
    so = xuat_pdf(BA, os.path.join(dem, 'ai.pdf'))
    if not (1 <= so <= 7): loi.append('(g) PDF bản AI %d trang' % so)
    return loi, time.time() - t0


if __name__ == '__main__':
    loi, dt = chay()
    print('T15 nâng cao bằng AI (giả lập, không gọi mạng) – %.1f giây: %s' % (dt, 'ĐẠT' if not loi else 'LỖI %d' % len(loi)))
    for x in loi: print('   ', x)
    sys.exit(1 if loi else 0)
