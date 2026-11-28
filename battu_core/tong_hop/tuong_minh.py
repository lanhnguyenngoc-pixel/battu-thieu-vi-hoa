# -*- coding: utf-8 -*-
"""Viết lại luận giải TƯỜNG MINH (28/11/2026, chủ dự án duyệt): mỗi mục có bốn phần – Kết luận, Căn cứ trên lá số,
Thời điểm (vận, năm), Hệ quả cụ thể. KHÔNG đổi kết luận gốc của Thiệu Vĩ Hoa: chỉ trình bày lại chữ và gom thêm chi tiết
(vận, năm, hành, tạng phủ) từ dữ liệu đã có. Hai yếu tố trái nhau được tách theo giai đoạn.

Mọi hàm nhận dữ liệu đã tính (hs, pd) và trả về chữ/khối; không tính lại gì của workbook. Chỉ thư viện chuẩn."""
import re

_XAU = ('Xấu', 'Rất xấu'); _TOT = ('Tốt', 'Rất tốt')

# Tạng phủ theo hành và chuyên khoa nên khám (để người đọc lá số trao đổi với đương sự)
KHOA = {'Mộc': 'gan – mật, gân, mắt (khám gan mật, mắt, thần kinh)',
        'Hỏa': 'tim – mạch, huyết áp (khám tim mạch, điện tâm đồ, huyết áp)',
        'Thổ': 'tỳ vị, dạ dày (khám tiêu hóa, nội soi dạ dày)',
        'Kim': 'phổi – đại tràng, da (khám hô hấp, chụp phổi, nội soi đại tràng)',
        'Thủy': 'thận – bàng quang, xương khớp, tai (khám thận – tiết niệu, xương khớp, tai)'}
KHOA_NGAN = {'Mộc': 'gan – mật, mắt', 'Hỏa': 'tim mạch', 'Thổ': 'tiêu hóa', 'Kim': 'hô hấp', 'Thủy': 'thận – xương khớp'}


def bo_nghieng(t):
    """'nghiêng nhẹ về hung' → 'hung'; 'nghiêng về giàu' → 'giàu'. Không đổi chiều kết luận, chỉ bỏ từ giảm nhẹ."""
    t = re.sub(r'(?i)nghiêng\s+(?:nhẹ\s+)?về\s+', '', t or '')
    return re.sub(r'(?i)\bhơi\s+', '', t).strip()


def _ten_van(d):
    t = d.get('tuoi_bat_dau')
    return 'vận %s (%s, %s–%s)' % (d['ten'], ('%s–%s tuổi' % (t, t + 9)) if isinstance(t, int) else 'tuổi ?', d['nam_bat_dau'], d['nam_ket_thuc'])


def dong_tung_van(pd):
    """Mỗi đại vận một dòng riêng: tuổi, năm, mức, lý do ngắn (không gộp giai đoạn)."""
    out = []
    for d in pd['dai_van']:
        ly = re.split(r'\s+[–-]\s+', d.get('giai_thich') or '', 1)
        ly = re.sub(r'\s*\(P[12] tr[^)]*\)', '', ly[1] if len(ly) > 1 else ly[0])
        m = re.search(r'làm chủ cả 10 năm;?\s*(.*)', ly, re.S); ly = _gon(m.group(1) if m else ly, 80)
        out.append('%s: %s%s' % (_ten_van(d)[0].upper() + _ten_van(d)[1:], (d.get('muc') or '—').lower(), (' – ' + ly) if ly else ''))
    return out


def _ds_van(ds): return '; '.join(_ten_van(d) for d in ds) if ds else 'không có vận nào'
def _tom_van(ds): return ', '.join('%s (%s–%s)' % (d['ten'], d['nam_bat_dau'], d['nam_ket_thuc']) for d in ds) if ds else 'không có'


def _ds(xs): return ', '.join(str(x) for x in xs) if xs else 'không có'


def van_tot_xau(pd):
    dv = pd['dai_van']
    return ([d for d in dv if d['muc'] in _TOT], [d for d in dv if d['muc'] in _XAU])


def _bo_nguon(t):
    """Bỏ chú thích kỹ thuật '(PP1 xếp …, PP2 xếp … – kết luận theo PP1 – …)' khỏi chữ gốc, trả (chữ sạch, ghi chú một câu)."""
    m = re.search(r'\s*\(PP1 xếp ([^,]+), PP2 xếp ([^–)]+?)\s*–[^)]*\)', t or '')
    sach = re.sub(r'\s*\(PP1 xếp[^)]*\)', '', t or '').strip()
    gc = ''
    if m and m.group(1).strip().lower() != m.group(2).strip().lower():
        gc = 'Hai cách tính vượng suy cho kết quả khác nhau (cách 1: %s; cách 2: %s); kết luận lấy theo cách 1, cách được khuyến nghị.' % (
            m.group(1).strip().lower(), m.group(2).strip().lower())
    return sach, gc


# ------------------------------------------------------------------ lục thân, hôn nhân
def _tach_ds(t):
    """'Tân Hợi (37–46 tuổi, 2010–2019) Canh Tuất (…)' → mỗi vận cách nhau dấu chấm phẩy."""
    return re.sub(r'\)\s+(?=[A-ZÀ-Ỹ])', '); ', t or '')


def _tach_van(chinh):
    """Từ chữ workbook tách: câu kết luận, danh sách vận khởi sắc, danh sách vận kém."""
    t, _ = _bo_nguon(chinh)
    t = re.sub(r'^[A-ZÀ-Ỹ ]+:\s*', '', t)
    ks = re.search(r'Khởi sắc trong đại vận[^:]*:\s*(.*?)(?:;|\.)\s*(?:kém đi|Kém đi|$)', t, re.S)
    kem = re.search(r'[kK]ém đi trong đại vận[^:]*:\s*(.*?)(?:\.|$)', t, re.S)
    chinh_cau = re.split(r'\s*Khởi sắc', t, 1)[0].strip()
    return chinh_cau, _tach_ds(ks.group(1).strip() if ks else ''), _tach_ds(kem.group(1).strip() if kem else '')


def luc_than(ten, d, nam_dv, nam_xau_cha_me=None):
    """Một dòng lục thân theo bốn phần. d = pd['luc_than'][ten]; nam_dv: hàm đổi 'Tân Hợi (37–46 tuổi)' thêm năm dương lịch."""
    muc = (d.get('muc') or '—')
    cau, ks, kem = _tach_van(d.get('chinh') or '')
    _, gc = _bo_nguon(d.get('chinh') or '')
    out = '**%s: %s – %s**' % (ten, muc.lower(), cau.rstrip('.').rstrip(',')[:1].lower() + cau.rstrip('.')[1:] + '.')
    if ks or kem:
        out += ' Theo giai đoạn: %s%s%s.' % (
            ('tốt lên ở các vận ' + nam_dv(ks)) if ks else '',
            '; ' if ks and kem else '',
            ('kém đi ở các vận ' + nam_dv(kem)) if kem else '')
    if nam_xau_cha_me is not None:
        out += ' Năm nên chú ý sức khỏe %s trong 10 năm tới: %s.' % (ten.lower(), _ds(nam_xau_cha_me))
    for n in d.get('nguon_khac') or []:
        out += ' Nguồn tham khảo khác (trọng số thấp hơn) đọc ngược lại: %s.' % re.sub(r'^[^:]+:\s*', '', n.get('noi', '')).rstrip('.')
    if gc: out += ' ' + gc
    return out


def hon_nhan(vc, nam_dv, gd_tot, gd_xau, lan, van_xung_cung, la_nam):
    """Đoạn Hôn nhân theo bốn phần."""
    muc = (vc.get('muc') or '—')
    cau, ks, kem = _tach_van(vc.get('chinh') or '')
    _, gc = _bo_nguon(vc.get('chinh') or '')
    pt = 'vợ' if la_nam else 'chồng'
    out = '**Hôn nhân: %s.** Kết luận: %s.' % (muc.lower(), cau.rstrip('.'))
    if ks or kem:
        out += ' Theo giai đoạn: %s%s%s.' % (('thuận ở các vận ' + nam_dv(ks)) if ks else '', '; ' if ks and kem else '',
                                              ('trắc trở ở các vận ' + nam_dv(kem)) if kem else '')
    if van_xung_cung:
        out += ' Vận xung, khắc cung %s (chi ngày): %s – giai đoạn dễ bất hòa, xa cách hoặc đổ vỡ nhất.' % (pt, '; '.join(van_xung_cung))
    out += ' Năm tốt cho kết hôn, gia đạo êm ấm: %s. Năm bất hòa, cần nhường nhịn: %s.' % (_ds(gd_tot), _ds(gd_xau))
    if lan: out += ' Năm vừa có dấu hiệu hợp vừa có dấu hiệu khắc (gia đạo dễ biến động hai chiều): %s.' % _ds(lan)
    if gc: out += ' ' + gc
    return out


# ------------------------------------------------------------------ thọ – yểu
def _lay(dien_giai, tien_to):
    m = re.search(re.escape(tien_to) + r'[^:]*:\s*(.*?)(?:\.\s+(?=[A-ZÀ-Ỹ])|\.$|$)', dien_giai or '', re.S)
    return m.group(1).strip().rstrip('.') if m else ''


def tho_yeu(ty, pd, hs, ky, dung, hanh_info, khoi_nam):
    """Hai khối: đoạn thọ – yểu bốn phần. khoi_nam: chữ các năm đáng chú ý nhất (tai_uong.khoi_nam_chu_y)."""
    dg = ty.get('dien_giai') or ''
    dau = (ty.get('muc') or '').split(' – ')[0].strip()
    m_tho = re.search(r'Dấu hiệu mệnh thọ \((\d+)\)', dg); m_yeu = re.search(r'Dấu hiệu mệnh yểu \((\d+)\)', dg)
    n_tho = int(m_tho.group(1)) if m_tho else 0; n_yeu = int(m_yeu.group(1)) if m_yeu else 0
    tho = _lay(dg, 'Dấu hiệu mệnh thọ'); yeu = _lay(dg, 'Dấu hiệu mệnh yểu')
    hg = _lay(dg, 'Yếu tố hóa giải'); nang = _lay(dg, 'Yếu tố làm nặng')
    tot, xau = van_tot_xau(pd)
    out = '**Thọ – yểu: mệnh %s.** Căn cứ: %d dấu hiệu thọ%s; %d dấu hiệu yểu%s.' % (
        dau.lower(), n_tho, (' (%s)' % tho) if tho else '', n_yeu, (' (%s)' % yeu) if yeu else '')
    if hg or nang:
        out += ' Theo giai đoạn: yếu tố làm nặng%s phát huy ở các vận xấu (nguy hiểm nhất cho sức khỏe); ' \
               'yếu tố hóa giải%s giúp qua hạn ở các vận thuận. Từng đại vận xem danh sách ngay dưới.' % (
                   (' (%s)' % nang) if nang else '', (' (%s)' % hg) if hg else '')
    else:
        out += ' Từng đại vận xem danh sách ngay dưới.'
    if khoi_nam: out += ' Năm có dấu hiệu nặng nhất trong 10 năm từ năm xem: %s.' % khoi_nam
    cho = [h for h in (ky, dung) if h in hanh_info]
    if cho:
        out += ' Hệ quả: %s sức khỏe suy ở %s; %s' % (
            'khi kỵ thần %s thái quá và dụng thần %s bị khắc,' % (ky, dung) if ky and dung else '',
            ' và '.join(KHOA_NGAN[h] for h in cho if h in KHOA_NGAN),
            'nguy cơ tuổi thọ ngắn tập trung ở các vận xấu nêu trên.' if dau.lower().startswith('yểu') else
            'thể trạng nhìn chung bền, vẫn cần giữ sức ở các vận xấu nêu trên.')
    return out


# ------------------------------------------------------------------ sức khỏe: bệnh tật, tai họa
def benh_tat(ky, dung, dong_10_nam, hanh_info):
    """Dòng Bệnh tật – tạng phủ: tạng nào, năm nào, khám gì. dong_10_nam = tai_uong.bang_suc_khoe(hs)."""
    nang = [d for d in dong_10_nam if d['tang_phu_muc'] == 'NẶNG']
    dp = [d for d in dong_10_nam if d['tang_phu_muc'] == 'ĐỀ PHÒNG']
    s = 'Bệnh tật – tạng phủ: dễ bệnh nhất ở %s' % ' và '.join(KHOA[h] for h in (ky, dung) if h in KHOA) if (ky in KHOA or dung in KHOA) else 'Bệnh tật – tạng phủ'
    if nang: s += '; năm có dấu hiệu nặng: %s' % '; '.join('%d (%s)' % (d['nam'], d['tang_phu_ten']) for d in nang)
    if dp: s += '; năm cần đề phòng: %s' % '; '.join('%d (%s)' % (d['nam'], d['tang_phu_ten']) for d in dp[:6])
    if not nang and not dp: s += '; 10 năm tới không có năm nào đạt ngưỡng cảnh báo tạng phủ, nên khám định kỳ mỗi năm'
    else: s += '; nên khám đúng chuyên khoa trước các năm này'
    return s


def _gon(t, n=90):
    t = re.sub(r'\s+', ' ', t or '').strip().rstrip('.;')
    if len(t) <= n: return t
    c = t[:n]; k = max(c.rfind('; '), c.rfind(', '), c.rfind(' – '))
    return (c[:k] if k > n // 3 else c.rsplit(' ', 1)[0]).rstrip(' ;,–:(')


def tai_hoa(dong_10_nam, tai_uong_cau, van_xau):
    """Dòng Thương tật – tai họa: năm, loại, vận."""
    nam = [d for d in dong_10_nam if d['so_dau_hieu_nang'] or d['kieu']]
    s = 'Thương tật – tai họa'
    if nam:
        s += ': năm có dấu hiệu hạn nặng hoặc dính kiểu lá số dễ gặp tai ương: ' + '; '.join(
            '%d (%s)' % (d['nam'], '; '.join(([('%d dấu hiệu hạn nặng' % d['so_dau_hieu_nang'])] if d['so_dau_hieu_nang'] else []) +
                                              [x.split(':')[0] for x in d['kieu']] +
                                              ([_gon(re.sub(r'^\[[^\]]*\]\s*', '', d['dau_hieu'][0]), 90)] if d['dau_hieu'] else []))) for d in nam[:6])
    else:
        s += ': 10 năm tới không có năm nào đạt ngưỡng hạn nặng'
    if van_xau: s += '; vận cần đi lại, lao động, lái xe cẩn thận nhất: %s' % '; '.join(_ten_van(d) for d in van_xau)
    return s
