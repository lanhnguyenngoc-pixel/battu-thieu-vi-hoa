# -*- coding: utf-8 -*-
"""Bản luận giải tổng hợp ~5 trang A4 – KHÔNG CẦN MẠNG (đề xuất hoàn thiện 15/11/2026, giai đoạn T4).

Soạn từ hồ sơ dữ kiện (du_kien) + kết luận đã gỡ mâu thuẫn (phan_doan), 10 mục, khoảng 5 trang A4 (chủ dự án duyệt
15/11/2026): lá số → bản mệnh → tính cách → học vấn, sự nghiệp → tài lộc → hôn nhân, con cái → cha mẹ, anh em, quan hệ
→ sức khỏe → vận trình → hóa giải, kết luận. Mỗi mục mở bằng câu kết luận; mỗi việc chỉ một kết luận (lấy từ phan_doan).

Đầu ra là cấu trúc thuần (danh sách khối), để export/xuat_luan_giai.py dựng HTML/PDF/Word:
  ('h1', chữ) · ('h2', chữ) · ('p', chữ) · ('ul', [chữ]) · ('bang', [tiêu đề], [[ô]]) · ('nho', chữ)
Chữ có thể chứa **đậm**. Chỉ thư viện chuẩn."""
import re
from battu_core.tong_hop import quan_he as Q
from battu_core.tong_hop.du_kien import ho_so
from battu_core.tong_hop.phan_doan import phan_doan

NHAT_CAN = {
    'Giáp': ('cây lớn', 'thẳng thắn, có chí vươn lên, nhân hậu, trọng tình nghĩa', 'đôi khi cứng nhắc, khó uốn mình'),
    'Ất': ('cây cỏ mềm', 'khéo léo, mềm dẻo, giỏi thích nghi, giàu cảm xúc', 'dễ phụ thuộc, thiếu quyết đoán'),
    'Bính': ('mặt trời', 'nhiệt tình, cởi mở, hào phóng, thích được ghi nhận', 'nóng vội, thiếu kiên nhẫn'),
    'Đinh': ('ngọn lửa đèn', 'tinh tế, chu đáo, sâu sắc, giỏi soi xét', 'hay lo nghĩ, dễ tự ái'),
    'Mậu': ('núi đất', 'điềm tĩnh, chắc chắn, giữ chữ tín, dám chịu trách nhiệm', 'bảo thủ, chậm thay đổi'),
    'Kỷ': ('đất ruộng', 'khiêm nhường, tỉ mỉ, biết nuôi dưỡng người khác', 'đa nghi, hay tính toán chi li'),
    'Canh': ('kim loại thô', 'cương trực, quyết đoán, trọng nghĩa khí', 'cứng rắn, dễ va chạm'),
    'Tân': ('châu ngọc', 'tinh tế, trọng danh dự, nhạy cảm', 'kiêu, dễ tổn thương'),
    'Nhâm': ('nước sông lớn', 'thông minh, linh hoạt, tầm nhìn rộng, ham tự do', 'khó chịu gò bó, dễ buông thả'),
    'Quý': ('nước mưa sương', 'kín đáo, giàu trực giác, kiên nhẫn', 'hay suy tư, dễ bi quan')}
NHOM = {
    'Thực thương': ('sắc sảo, nhiều ý tưởng, giỏi diễn đạt', 'nói thẳng, không thích bị gò bó, dễ va chạm cấp trên',
                    'thành công nhờ tài năng, tay nghề và sản phẩm của chính mình; hợp nghề tự do, chuyên gia, tự kinh doanh'),
    'Tài': ('thực tế, nhạy với cơ hội và tiền bạc, giữ chữ tín', 'dễ ôm đồm, nặng lo toan vật chất',
            'thành công nhờ kinh doanh, quản lý tài sản, làm ra của cải cụ thể'),
    'Quan sát': ('có kỷ luật, trách nhiệm, được tín nhiệm', 'dễ chịu áp lực, cứng nhắc theo khuôn phép',
                 'hợp tổ chức, quản lý, công vụ – tiến thân bằng chức vụ'),
    'Ấn': ('hiếu học, chuyên môn sâu, được người trên nâng đỡ', 'dễ ỷ lại, học nhiều làm ít, chần chừ',
           'hợp nghiên cứu, giáo dục, nghề cần bằng cấp và uy tín học thuật'),
    'Tỉ kiếp': ('tự lực, có bạn bè, cộng sự', 'dễ cạnh tranh, tranh giành, chi tiêu rộng',
                'hợp tác, làm nhóm, tự lập nghiệp với cộng sự tin cậy')}
HANH_INFO = {
    'Mộc': ('xanh lá', 'Đông', 'giáo dục, xuất bản, gỗ – nông lâm, dệt may, dược thảo', 'gan – mật, gân, mắt'),
    'Hỏa': ('đỏ, cam, hồng, tím', 'Nam', 'công nghệ thông tin, điện – điện tử, truyền thông, năng lượng, ẩm thực, văn hóa, giảng dạy',
            'tim – mạch, huyết áp, mắt, giấc ngủ'),
    'Thổ': ('vàng, nâu đất', 'trung tâm, Tây Nam, Đông Bắc', 'bất động sản, xây dựng, kiến trúc, nông sản, tư vấn, môi giới, quản lý',
            'tỳ vị, dạ dày'),
    'Kim': ('trắng, bạc, ánh kim', 'Tây, Tây Bắc', 'tài chính – ngân hàng, cơ khí, kim loại, pháp luật, kỹ thuật', 'phổi – đại tràng, da'),
    'Thủy': ('đen, xanh dương', 'Bắc', 'vận tải, thương mại, du lịch, thủy sản, dịch vụ', 'thận – bàng quang, lưng, xương khớp, tai')}
THAP_THAN_NAM = {
    'Tỉ kiên': 'bạn bè, cạnh tranh; giữ tiền', 'Kiếp tài': 'dễ hao tài, thị phi; tránh hùn vốn, bảo lãnh',
    'Thực thần': 'ý tưởng, sản phẩm, thu nhập từ tài năng', 'Thương quan': 'tài năng nổi, dễ khẩu thiệt, va chạm cấp trên',
    'Thiên tài': 'cơ hội tiền bạc, đầu tư', 'Chính tài': 'thu nhập ổn định, tài sản, gia đạo',
    'Thất sát': 'áp lực, trách nhiệm lớn, cơ hội vị trí', 'Chính quan': 'chức vụ, danh tiếng, kỷ luật',
    'Thiên ấn': 'giấy tờ, học tập; dễ chần chừ', 'Chính ấn': 'được nâng đỡ, nhà cửa, giấy tờ'}
_TOT = ('Tốt', 'Rất tốt'); _XAU = ('Xấu', 'Rất xấu')


def _cau(t, n=1):
    """n câu đầu của chữ workbook (bỏ phần căn cứ, ngoặc vuông)."""
    t = re.sub(r'\s*\[[^\]]*\]', '', re.sub(r'\s*\((?:P1|P2|HG|UH)[^)]*\)', '', t or '')).strip()
    p = re.split(r'(?<=[.!?])\s+', t)
    return ' '.join(p[:n]).strip()


def _nhom(dung_than):
    m = re.search(r'\(([^)]+)\)', dung_than or '')
    g = (m.group(1) if m else '').strip()
    for k in NHOM:
        if k.lower() in g.lower(): return k
    return {'Thực thần': 'Thực thương', 'Thương quan': 'Thực thương', 'Chính tài': 'Tài', 'Thiên tài': 'Tài',
            'Chính quan': 'Quan sát', 'Thất sát': 'Quan sát', 'Chính ấn': 'Ấn', 'Thiên ấn': 'Ấn',
            'Tỉ kiên': 'Tỉ kiếp', 'Kiếp tài': 'Tỉ kiếp'}.get(g, '')


def _so_chu(nhom_txt, ten):
    m = re.search(re.escape(ten) + r'\s+(\d+)\s*chữ', nhom_txt or '')
    return int(m.group(1)) if m else None


def _ds(xs): return ', '.join(str(x) for x in xs) if xs else 'không có'


def _gop_ten(qh, loai):
    """Tên quan hệ đã gộp trùng ('lục xung Tí–Ngọ ×2')."""
    dem = {}
    for q in qh:
        if q['loai'] in loai: dem[q['ten']] = dem.get(q['ten'], 0) + 1
    return ['%s%s' % (k, ' ×%d' % v if v > 1 else '') for k, v in dem.items()]


def _quan_he_gon(qh, bo=('can_khac',)):
    """Gộp quan hệ trùng (Mậu–Quý / Quý–Mậu) thành một dòng kèm vị trí: 'ngũ hợp Mậu–Quý (năm–tháng, tháng–giờ)'."""
    nhom = {}
    for q in qh:
        if q['loai'] in bo: continue
        k = (q['loai'], frozenset(q['chu']))
        vt = '–'.join(x.split(' ', 1)[-1] for x in q['vi_tri'])
        nhom.setdefault(k, [q['ten'], []])[1].append(vt)
    return ['%s (%s)' % (ten, ', '.join(dict.fromkeys(vts))) for ten, vts in nhom.values()]


def _nam_dv(t, hs):
    """'Bính Dần (23–32 tuổi)' trong chữ workbook → kèm năm dương lịch của vận."""
    for d in hs['dai_van']:
        if d['ten'] in (t or '') and isinstance(d.get('tuoi_bat_dau'), int):
            t = re.sub(r'(%s \(%d\s*[–-]\s*\d+(?: tuổi)?)(?=[):])' % (re.escape(d['ten']), d['tuoi_bat_dau']),
                       r'\g<1>, %s–%s' % (d.get('nam_bat_dau'), d.get('nam_ket_thuc')), t)
    return t


def _viec_nam(y, hs, pd_y):
    """Việc chính của một năm: nghĩa thập thần can năm + quan hệ then chốt + hạn nặng."""
    ph = []
    tt = y.get('thap_than_can') or ''
    if tt: ph.append('%s: %s' % (tt, THAP_THAN_NAM.get(tt, '')))
    qh = y.get('quan_he', [])
    xd = [q for q in qh if q['loai'] == 'xung' and q.get('cham_dung')]
    if xd:
        ph.append('%s – chạm dụng thần%s' % (_gop_ten(xd, ('xung',))[0], ', cung vợ chồng' if any(q.get('cham_cung_phu_the') for q in xd) else ''))
    ph += [x for x in _gop_ten(qh, ('thien_khac_dia_xung',))]
    hop = [q for q in qh if q['loai'] in ('ban_hop', 'tam_hop', 'tam_hoi') and q.get('hanh') in (hs['ban_menh']['hanh_dung'], hs['ban_menh']['hanh_ky'])]
    for q in hop[:1]:
        ph.append('%s (%s)' % (q['ten'], 'thêm dụng thần' if q['hanh'] == hs['ban_menh']['hanh_dung'] else 'thêm kỵ thần'))
    for q in qh:
        if q['loai'] == 'luc_hop' and q.get('cham_cung_phu_the'): ph.append('%s – hợp cung vợ chồng' % q['ten']); break
    for q in qh:
        if q['loai'] == 'can_hop' and q['vi_tri'][1] == 'can ngày': ph.append('%s nhật chủ' % q['ten']); break
    bm = hs['ban_menh']
    for la_can, chu in ((True, y.get('can')), (False, y.get('chi'))):
        h = Q.HANH_CAN[Q.CAN.index(chu)] if la_can and chu in Q.CAN else (Q.HANH_CHI[Q.CHI.index(chu)] if chu in Q.CHI else None)
        if h is not None and Q.HANH[h] == bm.get('hanh_ky'): ph.append('%s %s (%s) là kỵ thần' % ('can' if la_can else 'chi', chu, Q.HANH[h]))
    for q in y.get('quan_he_voi_dai_van', []):
        if q['loai'] == 'xung': ph.append('tuế vận xung (%s)' % q['ten'].replace('lục xung ', ''))
    if pd_y.get('han_nang'): ph.append('%d dấu hiệu hạn nặng – giữ sức khỏe, đi lại cẩn thận' % pd_y['han_nang'])
    return '; '.join(ph)


THAN_SAT = {
    'NGUYỆT ĐỨC': 'phúc tinh, hóa hung thành cát', 'THIÊN ĐỨC': 'phúc tinh, gặp nạn có cứu', 'THIÊN ẤT QUÝ NHÂN': 'được quý nhân giúp',
    'THÁI CỰC QUÝ NHÂN': 'thông minh, chuyên chú, có đầu có đuôi', 'KIM DƯ': 'phú quý, tính ôn hòa', 'VĂN XƯƠNG': 'học hành, văn chương',
    'KIẾP SÁT': 'đề phòng tai nạn, mất mát', 'QUẢ TÚ': 'hôn nhân dễ muộn, ít người chia sẻ', 'CÔ THẦN': 'hôn nhân dễ muộn, tính cô độc',
    'KHÔNG VONG THEO TRỤ': 'việc thuộc trụ đó dễ hụt, đến muộn', 'DỊCH MÃ': 'đi lại, thay đổi chỗ ở – công việc',
    'ĐÀO HOA': 'duyên dáng, giao thiệp rộng', 'HOA CÁI': 'nghệ thuật, tâm linh, thích một mình', 'DƯƠNG NHẬN': 'cương mạnh, dễ va chạm',
    'KÌNH DƯƠNG': 'cương mạnh, dễ va chạm', 'HỒNG LOAN': 'duyên hôn nhân', 'THIÊN HỶ': 'tin vui', 'TƯỚNG TINH': 'có uy, dễ làm người đứng đầu',
    'CÔ LOAN': 'hôn nhân dễ trắc trở', 'TỨ PHẾ': 'sức yếu, việc dễ dang dở'}


def _diem(t, n=3):
    """Các ý '• …' của chữ workbook (bỏ căn cứ, mã câu) → tối đa n ý."""
    xs = [re.sub(r'\s*\[[^\]]*\]|\s*\((?:P1|P2|HG|UH|TTT)[^)]*\)', '', x).strip(' ;.') for x in re.split(r'•', t or '')[1:]]
    return [x for x in xs if x][:n]


def _y_nghia(t):
    """'NGHIÊNG THUẬN (… ) ► Ý nghĩa cụ thể: khá – …' → (mức, ý nghĩa)."""
    muc = re.split(r'\s*\(', t or '', 1)[0].strip()
    m = re.search(r'Ý nghĩa cụ thể:\s*([^►]*?)(?:\s*\[|$)', t or '', re.S)
    return muc.capitalize(), _cau(m.group(1), 1) if m else ''


def _kl(hs, muc, nhan):
    """Chữ của nhãn `nhan` trong mục KET_LUAN có tên bắt đầu bằng `muc` ('F. Sự nghiệp' → 'F. Sự nghiệp – công danh')."""
    for ten, d in hs['ket_luan_muc'].items():
        if ten.startswith(muc):
            return next((v for k, v in d.items() if k.startswith(nhan) and isinstance(v, str)), '')
    return ''


def _cau_con_cai(hs, pd, an_m, gon=False):
    """Đoạn kết luận Con cái của MỘT lá số (dùng cho lá chính và lá hôn phối)."""
    lt = pd['luc_than']
    d = lt.get('Con cái', {}); cc = hs.get('con_cai') or {}
    th = (cc.get('tong_hop') or '').strip()
    if th:      # mục 2.8: câu tổng hợp của workbook (đã gồm con đầu lòng)
        s2 = ('' if gon else '**Con cái:** ') + th.replace('CON CÁI: ', '', 1)
        nb = (cc.get('nam_bat_loi') or '').strip()
        if gon: nb = ''
        if nb and not nb.startswith('('): s2 += ' Năm có dấu hiệu bất lợi cho con (thương, thực gặp quan sát): %s.' % nb.rstrip('.;')
        if not gon: s2 += ' (Phần sức khỏe của con chỉ để tham khảo, không phải chẩn đoán.)'
    else:
        s2 = '**Con cái: %s.** %s' % (d.get('muc') or '—', _cau(d.get('chinh', '').split(':', 1)[-1], 1))
        if 'Con đầu lòng' not in an_m:
            cdl = _kl(hs, 'G. Gia đạo', 'Con đầu lòng').split('Vì sao')[0].strip()
            if cdl: s2 += ' Con đầu lòng (theo sách): %s.' % cdl.rstrip('.').lower()
    unh = _kl(hs, 'G. Gia đạo', 'Con cái (UH)')
    if 'không có con' in unh:
        s2 += ' Uyên Hải Tử Bình (trọng số thấp hơn) nêu sao con nhập mộ, vô khí – đọc là con muộn hoặc ít, không phải không có con.'
    vn = _diem(_kl(hs, 'G. Gia đạo', 'Vãn niên – tuổi già (HG)'), 1)
    if vn: s2 += ' Vãn niên: %s.' % vn[0].replace('[Hóa giải toàn thư] ', '').rstrip('.')
    return s2


def doc_tu_tru(txt):
    """\"Nhâm Ngọ, Canh Tuất, Ất Mão, Mậu Dần\" → (can[4], chi[4]); sai định dạng hoặc tên can chi lạ → None."""
    from battu_core.nhap_lieu import CAN, CHI
    cap = [x.strip().split() for x in re.split(r'[,;/]', txt or '') if x.strip()]
    if len(cap) != 4 or any(len(c) != 2 for c in cap): return None
    can, chi = [c[0].capitalize() for c in cap], [c[1].capitalize() for c in cap]
    return (can, chi) if all(x in CAN for x in can) and all(x in CHI for x in chi) else None


def _nam_theo_can_chi(can, chi, gan):
    """Năm dương lịch có can chi này, gần năm `gan` nhất (can chi lặp 60 năm); chỉ để điền ô năm sinh trừ năm khi nhập can chi."""
    from battu_core.nhap_lieu import CAN, CHI
    ds = [y for y in range(1900, 2100) if CAN[(y - 4) % 10] == can and CHI[(y - 4) % 12] == chi]
    return min(ds, key=lambda y: abs(y - gan)) if ds else None


def _dong_hon_phoi(la_so):
    """Mục 2.9: MỘT dòng tham khảo con cái theo lá số hôn phối. Có hai cách nhập: đủ ngày, tháng, năm, giờ sinh dương lịch, hoặc tứ trụ theo can chi.
    Không đổi kết luận nào của lá chính; lá hôn phối tính bằng đúng bộ tính hiện có, giới tính ngược với lá chính."""
    if la_so is None: return ''
    pn = list(getattr(la_so, 'phoi_ngau', None) or []) + [None] * 5
    d, m, y, h, mi = pn[:5]
    from battu_core.nhap_lieu import LaSo
    from battu_core.dich_vu import tinh
    nam = (la_so.gioi_tinh or 'Nam') == 'Nam'
    chung = dict(mui_gio=la_so.mui_gio, gioi_tinh='Nữ' if nam else 'Nam', nam_xem=la_so.nam_xem, vuong_suy=la_so.vuong_suy, che_do_nht=la_so.che_do_nht)
    if None not in (d, m, y, h):
        s2 = LaSo(cach_nhap='duong', ngay=int(d), thang=int(m), nam=int(y), gio=int(h), phut=int(mi or 0), **chung)
    else:
        tt = doc_tu_tru(getattr(la_so, 'phoi_tu_tru', ''))
        if not tt: return ''
        can, chi = tt
        gan = (la_so.nam or 1980) if la_so.cach_nhap != 'canchi' else (la_so.nam_sinh_tru_nam or 1980)
        ns = _nam_theo_can_chi(can[0], chi[0], gan)
        if ns is None: return ''
        s2 = LaSo(cach_nhap='canchi', can=can, chi=chi, nam_sinh_tru_nam=ns, **chung)
    try:
        kq2 = tinh(s2); hs2 = ho_so(kq2.E, s2); pd2 = phan_doan(hs2, s2)
    except Exception:
        return ''
    cau = _cau_con_cai(hs2, pd2, {x['muc'] for x in pd2['an_muc']}, gon=True).replace('**', '')
    tt = ', '.join('%s %s' % (t['can'], t['chi']) for t in hs2['tu_tru'])
    return 'Tham khảo thông tin về Con cái theo lá số của %s (%s) là: %s' % ('vợ' if nam else 'chồng', tt, cau)


dong_hon_phoi = _dong_hon_phoi


def soan(hs, pd, la_so=None, ho_ten=''):
    """→ danh sách khối của bản luận giải (10 mục, khoảng 5 trang A4)."""
    bm = hs['ban_menh']; tru = hs['tu_tru']; ls = hs['la_so']; B = []
    nc = tru[2]['can']; nam_xem = ls.get('nam_xem')
    dung, hy, ky = bm['hanh_dung'], bm['hanh_hy'], bm['hanh_ky']
    nhom_dung = _nhom(bm.get('dung_than')); nd = NHOM.get(nhom_dung)
    vs = pd['vuong_suy']; vuong = (bm.get('vuong_suy') or '').strip().upper().startswith('VƯỢNG')
    ty = pd['tho_yeu']; tl = pd['tai_loc']; lt = pd['luc_than']; an_m = {x['muc'] for x in pd['an_muc']}
    ln = {y['nam']: y for y in hs['luu_nien']}; pln = {y['nam']: y for y in pd['luu_nien']}
    nhom_txt = hs.get('thap_than_nhom')
    so = {k: _so_chu(nhom_txt, k) for k in ('Tỉ kiếp', 'Thực thương', 'Tài', 'Quan sát', 'Ấn')}
    do = bm.get('do_theo_hanh') or {}; tong = sum(do.values()) or 1

    B.append(('h1', 'Luận giải Bát tự%s' % (' – ' + ho_ten if ho_ten else '')))
    sinh = ''
    if la_so is not None and getattr(la_so, 'cach_nhap', '') == 'duong' and la_so.nam:
        sinh = '%s, sinh %02d:%02d ngày %02d/%02d/%d (dương lịch). ' % (ls.get('gioi_tinh') or '', la_so.gio or 0, la_so.phut or 0,
                                                                       la_so.ngay, la_so.thang, la_so.nam)
    B.append(('nho', '%sNăm xem %s · tuổi %s (tuổi mụ %s). Bản tổng hợp không cần mạng của phần mềm Bát Tự Thiệu Vĩ Hoa.' % (
        sinh, nam_xem, ls.get('tuoi_thuc'), ls.get('tuoi_mu'))))

    i_tom = len(B)                                              # tóm tắt nhanh – chèn sau khi đã có đủ kết luận

    # 1. Lá số
    B.append(('h2', '1. Lá số'))
    B.append(('p', 'Tứ trụ **%s**. Nhật chủ **%s %s**, sinh tháng %s.' % (
        ls['tu_tru'], nc, Q.HANH[Q.HANH_CAN[Q.CAN.index(nc)]] if nc in Q.CAN else '', tru[1]['chi'])))
    hang = lambda f: [t.get(f, '') for t in tru]
    B.append(('bang', ['', 'Năm', 'Tháng', 'Ngày', 'Giờ'],
              [['Thiên can'] + [('%s – %s' % (t['can'], t['thap_than_can'].split('–')[-1].strip())) if i != 2 else '%s – Nhật chủ' % t['can'] for i, t in enumerate(tru)],
               ['Địa chi'] + hang('chi_hanh'),
               ['Tàng can'] + [', '.join(x for x in (t['tang_ban_khi'], t['tang_trung_khi'], t['tang_du_khi']) if x and x != '—') for t in tru],
               ['Nhật can tại chi'] + [t['truong_sinh'].split(':')[-1].strip() for t in tru],
               ['Cung'] + ['tổ nghiệp, cha mẹ', 'cha mẹ, anh em', 'bản thân, vợ/chồng', 'con cái, hậu vận']]))
    if do:
        xs = sorted(do.items(), key=lambda x: -x[1]); thieu = [h for h, v in do.items() if v == 0]
        B.append(('p', '**Lực ngũ hành:** %s.%s Nhóm thập thần: %s.' % (
            ' · '.join('%s %d%%' % (h, round(100 * v / tong)) for h, v in xs), ' Không có hành %s.' % ', '.join(thieu) if thieu else '',
            ', '.join('%s %d chữ' % (k, v) for k, v in so.items() if v is not None))))
    qtt = _quan_he_gon(hs['quan_he_tu_tru'])
    if qtt: B.append(('p', '**Quan hệ trong trụ:** %s.' % '; '.join(qtt)))
    ts = re.findall(r'▸\s*([A-ZÀ-Ỹ ]+?)\s*\(([^)]*)\)', bm.get('than_sat_hieu_luc') or '')
    if ts:
        B.append(('ul', ['**%s** (%s)%s' % (a.strip().capitalize(), b_, (': ' + THAN_SAT[a.strip()]) if a.strip() in THAN_SAT else '') for a, b_ in ts]))
        B.insert(len(B) - 1, ('p', '**Thần sát còn hiệu lực** (yếu tố phụ, chỉ làm tăng giảm mức độ):'))

    # 2. Bản mệnh
    B.append(('h2', '2. Bản mệnh: vượng suy, cách cục, dụng thần'))
    chinh_cach = (bm.get('cach_cuc') or 'Chính cách').startswith('Chính')
    thong_nhat = 'Cả hai' in (bm.get('vuong_suy') or '')
    dau = ('Thân %s%s' % ('vượng' if vuong else 'nhược', ' nhẹ (gần trung hòa)' if vs['sat_nguong'] else '')) if chinh_cach else bm.get('cach_cuc').split('\n')[0]
    B.append(('p', '**%s. Dụng thần %s, hỷ thần %s, kỵ thần %s.**%s' % (
        dau, bm.get('dung_than'), bm.get('hy_than'), bm.get('ky_than'), ' Cách cục: %s.' % bm.get('cach_cuc') if chinh_cach else
        ' Cách đặc biệt chọn dụng thần theo thế của cách, không theo phù ức thân vượng – nhược.')))
    if chinh_cach and not thong_nhat:
        B.append(('p', 'Hai phương pháp đánh giá vượng suy cho kết quả khác nhau; kết luận trên theo phương pháp được khuyến nghị '
                       '(%s). Nên đối chiếu thêm với sự việc đã qua.' % re.sub(r'\s+', ' ', (bm.get('vuong_suy') or '').split('\n')[0])))
    elif chinh_cach and isinstance(vs.get('phe_minh_pt'), (int, float)):
        B.append(('p', vs['kiem_chung'] if vs['sat_nguong'] else 'Phe sinh trợ nhật chủ chiếm %.1f%% tổng lực (ngưỡng 40%%).' % vs['phe_minh_pt']))
    vsd = _cau(bm.get('vi_sao_dung'), 3)
    if vsd: B.append(('p', '**Vì sao chọn dụng thần.** %s' % vsd.replace('PP1: ', '')))
    if bm.get('dieu_hau'): B.append(('p', '**Điều hầu.** %s' % _cau(bm['dieu_hau'], 2)))
    from battu_core.tong_hop import dung_hy_ky_hien as H
    hn = Q.HANH[Q.HANH_CAN[Q.CAN.index(nc)]] if nc in Q.CAN else ''
    for t in (H.the_ngu_hanh(hn, vuong), H.theo_mua(hn, tru[1]['chi'])):
        if t: B.append(('p', t))
    cu = H.hy_cuu_ung(hn, bm.get('hanh_dung'), vuong, bool(bm.get('dung_khuyet')) and chinh_cach)
    if cu and chinh_cach: B.append(('p', cu))
    vk = H.hy_ky_viec(bm.get('dung_than') or '')
    if vk: B.append(('p', 'Tham khảo thêm – dụng thần gặp thuận (P1 tr.96–98): %s. Khi bị kỵ thần phá: %s. %s' % (vk[0], vk[1], H.LUU_Y_TR98)))
    if bm.get('ung_vien_dung'):
        B.append(('p', '**Ứng viên dụng thần thay thế (tham khảo).** %s %s' % (
            '; '.join(x.strip('• ').split(':')[0] + ' – ' + _cau(x.split(':', 1)[1], 1) for x in bm['ung_vien_dung'].split('\n') if ':' in x),
            (bm.get('muc_chac_dung') or '').split(';')[0] + '.')))
    if la_so is not None and any(r and r[0] for r in (getattr(la_so, 'su_kien', None) or [])):
        from battu_core.tong_hop.phan_doan import dung_than_su_kien
        kd = dung_than_su_kien(hs, la_so)
        B.append(('p', '**Kiểm chứng bằng sự việc đã qua (P2 tr.445).** %s%s' % (
            (bm.get('goi_y_su_kien') or '').replace('⇒', '–'), (' ' + kd['ket_luan']) if kd else '')))
    m = re.search(r'Cứu ứng:(.*?(?:KHÔNG CÓ CỨU|KHÔNG CỨU|CÓ CỨU)[^.;\n]*)', bm.get('ket_luan') or '', re.S)
    if m: B.append(('p', '**Bệnh và thuốc.** %s.' % re.sub(r'\s+', ' ', m.group(1)).replace('→', '–').strip().rstrip('.')))
    rows = []
    for vai, ten_h, nhom_t in (('Dụng thần', dung, bm.get('dung_than')), ('Hỷ thần', hy, bm.get('hy_than')), ('Kỵ thần', ky, bm.get('ky_than'))):
        g = _nhom(nhom_t)
        rows.append([vai, ten_h, g, NHOM[g][0] if g in NHOM and vai != 'Kỵ thần' else (NHOM[g][1] if g in NHOM else '')])
    B.append(('bang', ['Vai trò', 'Hành', 'Thập thần', 'Ý nghĩa'], rows))
    cat_hung = _kl(hs, 'I. Vận trình', 'Cát – hung').split('(')[0].strip().lower()
    sang = _kl(hs, 'F. Sự nghiệp', 'Sang – hèn').split('(')[0].strip().lower()
    B.append(('p', '**Mệnh cục.** Tài lộc: %s. Quyền chức: %s. Cát – hung: %s. Thọ – yểu: %s (lý do xem mục 8).' % (
        (tl['muc'] or '—').lower(), sang.replace('nghiêng về hèn', 'ở mức vừa (bảng P1 tr.173 nghiêng về "hèn" vì thiếu Quan sát)') or '—',
        cat_hung or '—', ty['muc'].lower())))

    # 3. Tính cách
    B.append(('h2', '3. Tính cách – năng khiếu'))
    hinh, tot, xau = NHAT_CAN.get(nc, ('', '', ''))
    B.append(('p', 'Nhật chủ %s là tượng %s: **%s**; mặt cần rèn là %s.' % (nc, hinh, tot, xau)))
    ul = []
    if nd: ul.append('Dụng thần là %s: %s; mặt trái: %s.' % (nhom_dung, nd[0], nd[1]))
    if so['Quan sát'] == 0: ul.append('Không có Quan sát: ít chịu khuôn khổ, nên tự đặt kỷ luật và hạn chót cho mình.')
    if so['Tài'] and so['Tài'] >= 4: ul.append('Tài tinh nhiều (%d chữ): thực tế, hào phóng, nhạy cơ hội; dễ ôm đồm.' % so['Tài'])
    if so['Ấn'] and so['Ấn'] >= 3 and 'Ấn' in (bm.get('ky_than') or ''): ul.append('Ấn là kỵ thần: hiếu học nhưng dễ "học nhiều làm ít", chần chừ chờ đủ điều kiện.')
    if so['Tỉ kiếp'] == 0: ul.append('Không có Tỉ kiếp: tự lo liệu, ít dựa vào anh em bạn bè.')
    if do:
        mx = max(do, key=do.get); ul.append('Hành mạnh nhất là %s (%d%%), quyết định nét nổi trội.' % (mx, round(100 * do[mx] / tong)))
    B.append(('ul', ul))
    dt = _diem(_kl(hs, 'D. Tính cách', 'Tính cách (chương 22)').split('Điểm cần lưu ý')[0], 4)
    nk = _diem(_kl(hs, 'D. Tính cách', 'Trí tuệ – năng khiếu').split('Điểm yếu')[0], 3)
    if nk: B.append(('p', '**Năng khiếu theo sách:** %s.' % '; '.join(nk)))
    if dt: B.append(('p', '**Đức tính theo sách:** %s.' % '; '.join(dt)))

    # 4. Học vấn – sự nghiệp – công danh
    B.append(('h2', '4. Học vấn – sự nghiệp – công danh'))
    B.append(('p', '**%s**' % ((nd[2].capitalize() + '.') if nd else 'Sự nghiệp theo dụng thần %s.' % dung)))
    rows = []
    for nhan, ten in (('Học tập – thi cử', 'Học tập – thi cử'), ('Quan vận – chức vụ', 'Chức vụ'), ('Danh tiếng', 'Danh tiếng')):
        muc, yn = _y_nghia(_kl(hs, 'F. Sự nghiệp', nhan))
        if muc and not (ten == 'Học tập – thi cử' and 'Mốc học tập' in an_m and False): rows.append([ten, muc, yn])
    if rows: B.append(('bang', ['Việc', 'Mức', 'Ý nghĩa'], rows))
    ul = _diem(_kl(hs, 'F. Sự nghiệp', 'Chức nghiệp (HG)'), 3)
    if so['Quan sát'] == 0 and not any('tự do' in x for x in ul): ul.append('Không có Quan sát: hợp nghề tự do, chuyên gia, tự kinh doanh hơn bộ máy cấp bậc cứng nhắc')
    if dung in HANH_INFO: ul.append('**Nghề hợp dụng thần %s:** %s' % (dung, HANH_INFO[dung][2]))
    if hy in HANH_INFO: ul.append('**Nghề hợp hỷ thần %s:** %s' % (hy, HANH_INFO[hy][2]))
    td = _diem(_kl(hs, 'F. Sự nghiệp', 'Thiên di'), 1)
    if td: ul.append('Đi lại – thay đổi: %s' % td[0])
    B.append(('ul', [x.rstrip('.') + '.' for x in ul]))
    cd = [n for n in pln if ln[n].get('thap_than_can') in ('Chính quan', 'Thất sát', 'Chính ấn') and pln[n]['muc'] in _TOT]
    va = [n for n in pln if ln[n].get('thap_than_can') == 'Thương quan']
    B.append(('p', '**Năm thuận công danh, vị trí:** %s. **Năm dễ va chạm cấp trên, nên giữ lời:** %s.' % (_ds(sorted(cd)), _ds(sorted(va)))))

    # 5. Tài lộc
    B.append(('h2', '5. Tài lộc'))
    B.append(('p', '**Tài lộc: %s.** %s' % (tl['muc'] or '—', tl['dien_giai'])))
    ul = ['Thành phần: ' + '; '.join('%s %s' % (k, str(v).lower()) for k, v in tl['thanh_phan'].items() if v)]
    ul += ['Nguồn tiền: ' + x for x in _diem(_kl(hs, 'E. Tài lộc', 'Nguồn tiền'), 3)]
    gc = _kl(hs, 'E. Tài lộc', 'Giữ của – hao tài')
    ul += ['Giữ của: ' + x for x in _diem(gc.split('Bất lợi')[0], 1)] + ['Hao tài: ' + x for x in _diem(gc.split('Bất lợi')[-1], 1) if 'Bất lợi' in gc]
    B.append(('ul', [x.strip().rstrip('.').strip() + '.' for x in ul]))
    giu = [y['nam'] for y in pd['luu_nien'] if (ln[y['nam']].get('thap_than_can') in ('Tỉ kiên', 'Kiếp tài') and vuong) or y['muc'] in _XAU]
    thuan = [y['nam'] for y in pd['luu_nien'] if y['muc'] in _TOT and y['nam'] not in giu
             and ln[y['nam']].get('thap_than_can') in ('Thiên tài', 'Chính tài', 'Thực thần', 'Thương quan')]
    B.append(('p', '**Năm thu nhập thuận:** %s. **Năm cần giữ tiền, tránh cho vay – hùn vốn – bảo lãnh:** %s.' % (_ds(thuan), _ds(giu))))
    if 'Tài' in (bm.get('hy_than') or '') + (bm.get('dung_than') or ''):
        B.append(('p', 'Tài là %s: nên chủ động tạo thêm nguồn thu và tích lũy tài sản thực (nhà, đất, tài sản sinh lời) trong các vận, năm thuận; vẫn giữ quỹ dự phòng.' % (
            'dụng thần' if 'Tài' in (bm.get('dung_than') or '') else 'hỷ thần')))

    # 6. Hôn nhân – con cái
    B.append(('h2', '6. Hôn nhân – con cái'))
    vc = lt.get('Vợ / chồng', {})
    B.append(('p', '**Hôn nhân: %s.** %s' % (vc.get('muc') or '—', _nam_dv(_cau(vc.get('chinh', '').split(':', 1)[-1], 2), hs))))
    ul = []
    so_tai = so['Tài'] if (ls.get('gioi_tinh') or 'Nam') == 'Nam' else so['Quan sát']
    if so_tai is not None:
        sao = 'Tài' if (ls.get('gioi_tinh') or 'Nam') == 'Nam' else 'Quan sát'
        o_ngay = any(sao.split()[0].lower() in (tru[2].get(k) or '').lower() for k in ('tang_ban_khi', 'tang_trung_khi', 'tang_du_khi'))
        ul.append('Sao phối ngẫu (%s) %d chữ%s' % (sao, so_tai, ', có ngay ở cung vợ chồng – phối ngẫu đảm đang, giúp được việc' if o_ngay else ''))
    pt = [q['ten'] for q in hs['quan_he_tu_tru'] if q.get('cham_cung_phu_the') and q['loai'] not in ('can_khac', 'kho')]
    if pt: ul.append('Cung vợ chồng (chi ngày %s): %s' % (tru[2]['chi'], '; '.join(dict.fromkeys(pt))))
    tsh = [a.strip().capitalize() for a, _ in ts if a.strip() in ('QUẢ TÚ', 'CÔ THẦN', 'CÔ LOAN', 'ĐÀO HOA', 'HỒNG LOAN')]
    if tsh: ul.append('Thần sát liên quan: %s' % ', '.join(tsh))
    c_ngay = tru[2]['chi']
    hop = sorted({x for (ba, _) in Q.TAM_HOP if c_ngay in ba for x in ba if x != c_ngay} |
                 {x for k in Q.LUC_HOP for x in k if c_ngay in k and x != c_ngay})
    ul.append('Người hợp tác, kết đôi thường bổ trợ tốt: nhật can hành %s; chi năm sinh %s (hợp với chi ngày %s). Cần xét đủ lá số hai người' % (
        ' hoặc '.join(x for x in (dung, hy) if x), ', '.join(hop), c_ngay))
    B.append(('ul', [x.rstrip('.') + '.' for x in ul]))
    gd_tot = [y['nam'] for y in hs['luu_nien'] if y.get('loi_ket_hon') or any(q['loai'] in ('luc_hop', 'can_hop') and (q.get('cham_cung_phu_the') or q['vi_tri'][1] == 'can ngày') for q in y['quan_he'])]
    gd_xau = [y['nam'] for y in hs['luu_nien'] if y.get('bat_loi_hon_nhan') or any(q['loai'] == 'xung' and q.get('cham_cung_phu_the') for q in y['quan_he'])]
    lan = [n for n in gd_tot if n in gd_xau]                    # một năm chỉ ở MỘT danh sách
    gd_tot = [n for n in gd_tot if n not in lan]; gd_xau = [n for n in gd_xau if n not in lan]
    t1, t2 = ('Năm thuận gia đạo', 'Năm cần nhường nhịn') if 'Thời điểm kết hôn' in an_m else ('Năm thuận kết hôn, tin vui', 'Năm cần thận trọng')
    B.append(('p', '**%s:** %s. **%s:** %s.%s' % (t1, _ds(gd_tot), t2, _ds(gd_xau),
                                                  ' Năm tốt xấu đan xen (có cả dấu hiệu hợp và dấu hiệu khắc): %s.' % _ds(lan) if lan else '')))
    s2 = _cau_con_cai(hs, pd, an_m)
    B.append(('p', s2))
    dp = _dong_hon_phoi(la_so)
    if dp: B.append(('p', dp))

    # 7. Cha mẹ – anh em – quan hệ
    B.append(('h2', '7. Cha mẹ – anh em – quan hệ xã hội'))
    ul = []
    for ten in ('Ông bà', 'Cha', 'Mẹ', 'Anh chị em'):
        d = lt.get(ten)
        if not d: continue
        s3 = '**%s: %s.** %s' % (ten, d['muc'] or '—', _nam_dv(_cau(d['chinh'].split(':', 1)[-1], 1), hs))
        if d.get('nguon_khac'): s3 += ' (%s)' % d['dien_giai']
        ul.append(s3)
    cha = [y['nam'] for y in hs['luu_nien'] if y.get('bat_loi_cha')]; me = [y['nam'] for y in hs['luu_nien'] if y.get('bat_loi_me')]
    ul.append('Năm nên quan tâm sức khỏe cha: %s; mẹ: %s.' % (_ds(cha), _ds(me)))
    qn = bm.get('quy_nhan') or ''
    ul.append('Quý nhân: %s.' % ('có trong tứ trụ' if qn and not qn.startswith('Không') else 'không có Thiên ất quý nhân trong tứ trụ – thành công dựa nhiều vào sức mình'))
    if so['Tỉ kiếp'] is not None:
        ul.append('Bạn bè, cộng sự (Tỉ kiếp %d chữ, %s): %s.' % (
            so['Tỉ kiếp'], 'phe gây lệch' if vuong and chinh_cach else 'phe có lợi' if chinh_cach else 'theo cách cục',
            'hợp tác được nhưng tách bạch tiền bạc' if vuong else 'nên dựa vào cộng sự tin cậy'))
    B.append(('ul', ul))

    # 8. Sức khỏe
    B.append(('h2', '8. Sức khỏe'))
    B.append(('p', '**Thọ – yểu: %s.** %s' % (ty['muc'].lower(), ty['dien_giai'] or _cau(ty['goc_workbook']))))
    ul = []
    for nhan in ('Sức khỏe chung', 'Bệnh tật – tạng phủ', 'Thương tật – tai họa theo vận'):
        muc, yn = _y_nghia(_kl(hs, 'H. Sức khỏe', nhan))
        if muc: ul.append('%s: %s%s' % (nhan, muc.lower(), (' – ' + yn) if yn else ''))
    if ky in HANH_INFO: ul.append('Hành kỵ %s thái quá: %s' % (ky, HANH_INFO[ky][3]))
    if dung in HANH_INFO: ul.append('Hành dụng %s bị khắc: %s' % (dung, HANH_INFO[dung][3]))
    dh = _cau(_kl(hs, 'J. Hóa giải', '[Kinh nghiệm] Điều hầu').split('.', 1)[-1], 2)
    if dh: ul.append('Sinh hoạt: %s' % dh.split('(')[0].strip())
    B.append(('ul', [x.rstrip('.') + '.' for x in ul]))
    rows = []
    for y in hs['luu_nien']:
        tp = y.get('tang_phu') or ''
        m2 = re.search(r'•\s*(NẶNG|ĐỀ PHÒNG|CHÚ Ý)\s*–\s*([^:;]+)', tp)
        if pln[y['nam']]['muc'] in _XAU or (m2 and m2.group(1) != 'CHÚ Ý'):
            rows.append(['%d %s' % (y['nam'], y['luu_nien']), pln[y['nam']]['muc'], ('%s – %s' % (m2.group(1).capitalize(), m2.group(2).strip())) if m2 else ''])
    if rows: B.append(('bang', ['Năm cần giữ sức', 'Mức năm', 'Tạng phủ cần chú ý'], rows))
    from battu_core.tong_hop import tai_uong as TU
    _tu = TU.phan_tich(hs); _cy = TU.nam_ngan(hs)
    _ct = TU.cau_tong_hop(hs) if _tu['co_it_nhat_mot'] else ''
    if _cy: _ct = (_ct + ' ' if _ct else '') + '**Năm cần chú ý (10 năm từ năm xem):** %s.' % _cy
    B.append(('nho', (_ct + ' ' if _ct else '') + 'Đây là xu hướng mệnh lý để tham khảo, không thay cho khám và tư vấn của bác sĩ; xem thêm trang "Sức khỏe 10 năm".'))

    # 9. Vận trình
    B.append(('h2', '9. Đại vận và lưu niên'))
    dv = pd['dai_van']; dang = next((d for d in dv if d.get('dang_hanh')), None)
    tot_dv = [d for d in dv if d['muc'] in _TOT]
    B.append(('p', '%s%s' % ('**Đang hành vận %s (%s–%s): %s.** ' % (dang['ten'], dang['nam_bat_dau'], dang['nam_ket_thuc'], dang['muc']) if dang else '',
                             'Các vận thuận: %s.' % ', '.join('%s (%s–%s)' % (d['ten'], d['nam_bat_dau'], d['nam_ket_thuc']) for d in tot_dv) if tot_dv else '')))
    gd = _kl(hs, 'I. Vận trình', 'BỐN GIAI ĐOẠN ĐỜI')
    gds = re.findall(r'■\s*TRỤ\s+(\S+)\s*\(([^)]*)\)\s*–\s*([^:]*?)(?:\s*\([^)]*\))?:\s*([A-ZÀ-Ỹ ]+?)\s{2,}', gd + '  ')
    if gds: B.append(('p', '**Bốn giai đoạn đời:** %s.' % '; '.join('%s (%s): %s' % (c.lower().strip(), a, d_.strip().lower()) for _, a, c, d_ in gds)))
    dvh = {d['ten']: d for d in hs['dai_van']}
    rows = []
    for d in dv[:9]:
        h = dvh.get(d['ten'], {})
        g = _gop_ten(h.get('quan_he', []), ('xung', 'thien_khac_dia_xung', 'tam_hop', 'tam_hoi'))
        rows.append(['%s (%s–%s)' % (d['ten'], d['nam_bat_dau'], d['nam_ket_thuc']), h.get('thap_than_can', ''), '**%s**' % d['muc'],
                     'cả 10 năm %s (can làm chủ, chi bổ trợ)' % d['muc6'].lower() + (('; ' + '; '.join(g)) if g else '')])
    B.append(('bang', ['Đại vận', 'Can vận', 'Mức', 'Ghi chú'], rows))
    i_dang = dv.index(dang) if dang in dv else None
    for d in ([dang] + ([dv[i_dang + 1]] if i_dang is not None and i_dang + 1 < len(dv) else [])) if dang else []:
        B.append(('p', '**Vận %s (%s–%s).** %s' % (d['ten'], d['nam_bat_dau'], d['nam_ket_thuc'], d['giai_thich'] or d['muc6'])))
    rows = [['%d %s (%s)' % (y['nam'], y['can_chi'], y['tuoi_mu']), '**%s**' % y['muc'], _viec_nam(ln[y['nam']], hs, y)] for y in pd['luu_nien']]
    B.append(('bang', ['Năm (tuổi mụ)', 'Mức', 'Việc chính'], rows))
    for n in (nam_xem, (nam_xem or 0) + 1, (nam_xem or 0) + 2):
        y = ln.get(n)
        if not y: continue
        vu = (y.get('viec_ung') or '').strip()
        tp = re.search(r'•\s*(NẶNG|ĐỀ PHÒNG|CHÚ Ý)\s*–\s*([^:;]+)', y.get('tang_phu') or '')
        dh = [x for x in _diem(y.get('dau_hieu'), 3)]
        B.append(('p', '**Năm %d %s – %s.** %s%s%s' % (
            n, y.get('luu_nien'), pln[n]['muc'].lower(), ('Việc ứng theo can năm (sách): %s. ' % vu.rstrip('.')) if vu else '',
            ('Dấu hiệu của năm: %s. ' % '; '.join(re.sub(r'^\[[^\]]*\]\s*', '', x) for x in dh)) if dh else '',
            ('Sức khỏe: %s – %s.' % (tp.group(1).lower(), tp.group(2).strip())) if tp else '')))
        nu = (y.get('nguoi_ung') or '').strip()
        if nu:
            cites = re.findall(r'\((P[12] tr\.[^)]*)\)', nu)
            chinh = re.sub(r'\s*\(P[12] tr\.[^)]*\)', '', nu)
            B.append(('p', '**%s** (📖 Căn cứ sách: %s)' % (chinh, '; '.join(dict.fromkeys(cites)))))

    # 10. Hóa giải – kết luận
    B.append(('h2', '10. Hóa giải, lời khuyên và kết luận'))
    rows = []
    if dung in HANH_INFO and hy in HANH_INFO:
        rows.append(['Màu sắc', '%s (dụng); %s (hỷ)' % (HANH_INFO[dung][0], HANH_INFO[hy][0]), HANH_INFO[ky][0] if ky in HANH_INFO else ''])
        rows.append(['Phương vị', '%s; %s' % (HANH_INFO[dung][1], HANH_INFO[hy][1]), HANH_INFO[ky][1] if ky in HANH_INFO else ''])
        rows.append(['Nghề – việc', nd[2] if nd else HANH_INFO[dung][2], NHOM.get(_nhom(bm.get('ky_than')), ('', '', ''))[1] if _nhom(bm.get('ky_than')) else ''])
    sh = _kl(hs, 'J. Hóa giải', '[Kinh nghiệm] Bổ dụng thần')
    gio = re.findall(r'giờ\s+([^;]+?\))', sh)
    if gio: rows.append(['Giờ làm việc quan trọng', '; '.join(gio[:2]), ''])
    if rows: B.append(('bang', ['Phương diện', 'Nên', 'Tránh'], rows))
    for nhan in ('[Sách] Phương vị năm xem', '[Sách] Phương vị năm sau'):
        pv = _kl(hs, 'J. Hóa giải', nhan)
        if pv:
            pv = re.sub(r'\s*\n\s*', ' ', pv)
            B.append(('p', '**%s.** %s' % (_cau(pv, 1).rstrip('.'), _cau(pv.split('.', 1)[-1].split('Căn cứ:')[0], 3))))
    lu = _cau(_kl(hs, 'J. Hóa giải', 'Lưu ý khi hóa giải'), 2)
    if lu: B.append(('nho', lu.replace('[Sách] ', '')))
    dep = sorted(sorted([y for y in pd['luu_nien'] if y['muc'] in _TOT], key=lambda y: -y['diem'])[:4], key=lambda y: y['nam'])
    can = [y['nam'] for y in pd['luu_nien'] if y['muc'] in _XAU]
    B.append(('p', '**Kết luận.** Lá số %s%s. %s Giai đoạn thuận: %s. Năm nên dồn sức cho việc lớn: %s; năm nên chậm lại, giữ sức khỏe và gia đạo: %s.' % (
        bm.get('cach_cuc', '').lower(), ', dụng thần %s (%s)' % (dung, nhom_dung) if nhom_dung else '',
        (nd[2].capitalize() + '.') if nd else '', ', '.join('%s (%s–%s)' % (d['ten'], d['nam_bat_dau'], d['nam_ket_thuc']) for d in tot_dv[:4]) or '—',
        _ds([y['nam'] for y in dep]), _ds(can))))
    dang_t = ('đang hành vận %s (%s–%s) – %s' % (dang['ten'], dang['nam_bat_dau'], dang['nam_ket_thuc'], dang['muc'].lower())) if dang else ''
    tom = ['**Bản mệnh:** %s; dụng thần %s, hỷ thần %s, kỵ thần %s.' % (dau.lower() if chinh_cach else dau, dung, hy, ky),
           '**Sự nghiệp:** %s.' % (nd[2] if nd else 'theo dụng thần %s' % dung),
           '**Tài lộc:** %s. **Hôn nhân:** %s. **Con cái:** %s.' % ((tl['muc'] or '—').lower(), (vc.get('muc') or '—').lower(), (lt.get('Con cái', {}).get('muc') or '—').lower()),
           '**Sức khỏe:** thọ – yểu %s; chú ý %s.' % (ty['muc'].lower(), HANH_INFO[ky][3] if ky in HANH_INFO else '—'),
           '**Vận trình:** %s; năm dồn sức: %s; năm cần chậm lại: %s.' % (dang_t, _ds([y['nam'] for y in dep]), _ds(can))]
    B[i_tom:i_tom] = [('h2', 'Tóm tắt nhanh'), ('ul', tom)]
    B.append(('nho', 'Độ tin cậy: vượng suy %s. Bản tổng hợp dựa trên kết quả phần mềm (khớp workbook) và quy tắc tổng hợp %s; '
                     'mệnh lý là xu hướng để tham khảo, không phải định mệnh.' % (
                         'sát ngưỡng – nên đối chiếu sự việc đã qua' if vs['sat_nguong'] else 'rõ', pd['phien_ban_quy_tac'])))
    return B


def _chu(B):
    """Toàn bộ chữ của các khối (để tìm, kiểm tra)."""
    ra = []
    for k in B:
        if k[0] in ('h1', 'h2', 'p', 'nho'): ra.append(k[1])
        elif k[0] == 'ul': ra += k[1]
        elif k[0] == 'bang': ra += list(k[1]) + [o for r in k[2] for o in r]
    return '\n'.join(ra)


def ban_luan_giai(kq):
    """Từ KetQua (battu_core.dich_vu) → (khối, hồ sơ, phán đoán)."""
    s = kq.la_so
    hs = ho_so(kq.E, s); pd = phan_doan(hs, s)
    return soan(hs, pd, s, getattr(s, 'ho_ten', '') or ''), hs, pd
