# -*- coding: utf-8 -*-
"""Nâng cao bản luận giải bằng AI – Claude API (đề xuất hoàn thiện 15/11/2026, giai đoạn T5). TÙY CHỌN, mặc định tắt.

Quyết định đã duyệt: KHÔNG gửi họ tên; người dùng tự nhập khóa API và tự trả phí. Chỉ gửi:
  • hồ sơ gọn (tứ trụ, kết luận đã tính, kết luận đã gỡ mâu thuẫn – không họ tên, không ngày giờ sinh chi tiết);
  • bản không mạng (viet_ban) làm khung.
AI chỉ được VIẾT LẠI cho mượt và sâu hơn, KHÔNG được đổi dữ kiện. Bộ kiểm tra (kiem_tra) đối chiếu bản AI với hồ sơ:
sai tứ trụ, dụng – hỷ – kỵ, năm đại vận, can chi hay mức của năm → báo lỗi và giữ bản không mạng.
Chỉ thư viện chuẩn (urllib); không gọi mạng khi không bấm nút."""
import json, re, os, sys, ssl, hashlib, subprocess, urllib.request, urllib.error
from battu_core.tong_hop import quan_he as Q

API_URL = 'https://api.anthropic.com/v1/messages'
API_VERSION = '2023-06-01'
MO_HINH = [('claude-sonnet-5-5', 'Claude Sonnet 5.5 – cân bằng chất lượng và chi phí (khuyên dùng)'),
           ('claude-opus-5-5', 'Claude Opus 5.5 – sâu nhất, chi phí cao hơn'),
           ('claude-haiku-5-5', 'Claude Haiku 5.5 – nhanh, rẻ')]
MO_HINH_MAC_DINH = MO_HINH[0][0]
MAX_TOKENS = 12000
PHIEN_BAN_LOI_NHAC = 'T5-1'


class LoiAI(Exception):
    pass


# ------------------------------------------------------------------ hồ sơ gọn (dữ liệu gửi đi)
def ho_so_gon(hs, pd):
    """Dữ kiện tối thiểu cho AI – KHÔNG họ tên, không ngày giờ sinh (chỉ năm sinh, giới tính)."""
    bm = hs['ban_menh']
    dv_hs = {d['ten']: d for d in hs['dai_van']}
    ln_hs = {y['nam']: y for y in hs['luu_nien']}
    return {
        'gioi_tinh': hs['la_so'].get('gioi_tinh'), 'nam_sinh': hs['la_so'].get('nam_sinh'), 'nam_xem': hs['la_so'].get('nam_xem'),
        'tuoi_thuc': hs['la_so'].get('tuoi_thuc'), 'tuoi_mu': hs['la_so'].get('tuoi_mu'), 'tu_tru': hs['la_so'].get('tu_tru'),
        'tru': [{k: t.get(k) for k in ('tru', 'can', 'chi', 'thap_than_can', 'chi_hanh', 'tang_ban_khi', 'tang_trung_khi', 'tang_du_khi', 'truong_sinh')}
                for t in hs['tu_tru']],
        'thap_than_nhom': hs.get('thap_than_nhom'),
        'ban_menh': {k: bm.get(k) for k in ('vuong_suy', 'phe_minh_pt', 'do_theo_hanh', 'cach_cuc', 'dung_than', 'hy_than', 'ky_than',
                                            'vi_sao_dung', 'dieu_hau', 'than_sat_hieu_luc', 'quy_nhan')},
        'quan_he_tu_tru': [q['ten'] + ' (' + '–'.join(q['vi_tri']) + ')' for q in hs['quan_he_tu_tru'] if q['loai'] != 'can_khac'],
        'ket_luan_da_go': {
            'vuong_suy': pd['vuong_suy'], 'tho_yeu': pd['tho_yeu'], 'tai_loc': pd['tai_loc'], 'an_muc': pd['an_muc'],
            'luc_than': {k: {'muc': v['muc'], 'chinh': v['chinh'][:400], 'ghi_chu': v.get('dien_giai', '')} for k, v in pd['luc_than'].items()},
            'dai_van': [{'ten': d['ten'], 'nam': '%s–%s' % (d['nam_bat_dau'], d['nam_ket_thuc']), 'thap_than_can': dv_hs.get(d['ten'], {}).get('thap_than_can'),
                         'muc': d['muc'], '5_nam_dau': d['muc_5_nam_dau'], '5_nam_sau': d['muc_5_nam_sau'], 'can_cu': d['ly_do']} for d in pd['dai_van']],
            'luu_nien': [{'nam': y['nam'], 'can_chi': y['can_chi'], 'tuoi_mu': y['tuoi_mu'], 'muc': y['muc'], 'can_cu': y['ly_do'],
                          'thap_than_can': ln_hs.get(y['nam'], {}).get('thap_than_can'), 'viec_ung': ln_hs.get(y['nam'], {}).get('viec_ung'),
                          'tang_phu': (ln_hs.get(y['nam'], {}).get('tang_phu') or '')[:200]} for y in pd['luu_nien']]},
        'ket_luan_workbook': {k: (v[:600] if isinstance(v, str) else v) for k, v in hs.get('tom_tat', {}).items()},
    }


def _md(B):
    """Khối viet_ban → markdown (khung gửi cho AI)."""
    ra = []
    for i, k in enumerate(B):
        if k[0] == 'h1': continue
        if k[0] == 'nho' and i <= 1: continue              # dòng đầu có ngày giờ sinh – không gửi (chỉ năm sinh ở hồ sơ gọn)
        if k[0] == 'h2': ra.append('\n## ' + k[1])
        elif k[0] in ('p', 'nho'): ra.append(k[1])
        elif k[0] == 'ul': ra += ['- ' + x for x in k[1]]
        elif k[0] == 'bang':
            ra.append('| ' + ' | '.join(k[1]) + ' |'); ra.append('|' + '---|' * len(k[1]))
            ra += ['| ' + ' | '.join(o.replace('|', '/') for o in r) + ' |' for r in k[2]]
    return '\n'.join(ra).strip()


HE_THONG = """Bạn là chuyên gia luận giải Bát tự (Tứ trụ) nhiều năm kinh nghiệm, theo trường phái Thiệu Vĩ Hoa (Tứ trụ dự đoán học), có đối chiếu Uyên Hải Tử Bình, Cùng Thông Bảo Giám, Trích Thiên Tủy.
Nhiệm vụ: VIẾT LẠI bản luận giải dưới đây thành văn phong chuyên gia, mạch lạc, sâu sắc, dễ hiểu, dài khoảng 5 trang A4 (2.600–3.400 chữ), tiếng Việt.
QUY TẮC BẮT BUỘC:
1. Chỉ dùng dữ kiện trong HỒ SƠ và BẢN KHUNG. Không tự tính lại lá số, không đổi tứ trụ, thập thần, vượng suy, dụng thần, hỷ thần, kỵ thần, cách cục.
2. Giữ nguyên MỨC của mỗi đại vận và mỗi năm (Rất tốt / Tốt / Trung bình / Xấu / Rất xấu) và các danh sách năm. Có thể giải thích vì sao, không được đổi.
3. Giữ đúng 11 đề mục và thứ tự: "## Tóm tắt nhanh", "## 1. Lá số", "## 2. Bản mệnh: vượng suy, cách cục, dụng thần", "## 3. Tính cách – năng khiếu", "## 4. Học vấn – sự nghiệp – công danh", "## 5. Tài lộc", "## 6. Hôn nhân – con cái", "## 7. Cha mẹ – anh em – quan hệ xã hội", "## 8. Sức khỏe", "## 9. Đại vận và lưu niên", "## 10. Hóa giải, lời khuyên và kết luận".
4. Giữ hai bảng ở mục 9: bảng đại vận (cột: Đại vận | Can vận | Mức | Ghi chú; ô Đại vận dạng "Tên (năm đầu–năm cuối)") và bảng lưu niên (cột: Năm (tuổi mụ) | Mức | Việc chính; ô đầu dạng "2026 Bính Ngọ (49)").
5. Mỗi mục mở bằng câu kết luận. Văn xuôi là chính, gạch đầu dòng khi liệt kê. Không lặp ý giữa các mục. Không dùng ngôn từ nặng nề, hù dọa; điều bất lợi nói rõ kèm cách ứng xử.
6. Sức khỏe chỉ là xu hướng mệnh lý, nhắc khám bác sĩ; không đưa lời khuyên y khoa cụ thể. Không hứa hẹn tiền bạc, không khuyên đầu tư cụ thể.
7. Định dạng: chỉ dùng "## " cho đề mục, "- " cho gạch đầu dòng, **đậm**, và bảng markdown "| … |". Không dùng tiêu đề cấp khác, không HTML, không lời mở đầu hay kết thúc ngoài bản luận giải."""


def loi_nhac(hs, pd, B):
    return ('HỒ SƠ (JSON, đã tính và đã gỡ mâu thuẫn – là nguồn sự thật):\n%s\n\nBẢN KHUNG (bản không mạng của phần mềm – viết lại cho hay hơn, '
            'giữ đúng dữ kiện):\n%s' % (json.dumps(ho_so_gon(hs, pd), ensure_ascii=False), _md(B)))


def uoc_luong_token(hs, pd, B):
    """Ước lượng thô (tiếng Việt ≈ 2,5 ký tự / token) → (token vào, token ra tối đa). Giá: xem bảng giá Anthropic Console."""
    n = len(HE_THONG) + len(loi_nhac(hs, pd, B))
    return int(n / 2.5), MAX_TOKENS


# ------------------------------------------------------------------ gọi API
def _ngu_canh_ssl():
    """SSLContext có chứng chỉ gốc. Bộ cài macOS (PyInstaller) có thể không thấy kho chứng chỉ của Python → nạp chứng chỉ gốc
    hệ thống bằng lệnh security (chỉ khi kho rỗng)."""
    ctx = ssl.create_default_context()
    if sys.platform == 'darwin' and not ctx.cert_store_stats().get('x509_ca'):
        try:
            pem = subprocess.run(['/usr/bin/security', 'find-certificate', '-a', '-p',
                                  '/System/Library/Keychains/SystemRootCertificates.keychain'], capture_output=True, text=True, timeout=20).stdout
            if pem: ctx.load_verify_locations(cadata=pem)
        except Exception:
            pass
    return ctx


def goi_api(khoa, hs, pd, B, mo_hinh=MO_HINH_MAC_DINH, thoi_gian=180, _mo=None):
    """→ chữ markdown của AI. _mo: hàm thay urllib.request.urlopen (kiểm thử)."""
    if not khoa: raise LoiAI('Chưa nhập khóa API.')
    body = json.dumps({'model': mo_hinh, 'max_tokens': MAX_TOKENS, 'system': HE_THONG,
                       'messages': [{'role': 'user', 'content': loi_nhac(hs, pd, B)}]}).encode('utf-8')
    req = urllib.request.Request(API_URL, data=body, method='POST', headers={
        'x-api-key': khoa, 'anthropic-version': API_VERSION, 'content-type': 'application/json'})
    try:
        mo = _mo or (lambda rq, timeout: urllib.request.urlopen(rq, timeout=timeout, context=_ngu_canh_ssl()))
        with mo(req, timeout=thoi_gian) as r:
            d = json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        try: msg = json.loads(e.read().decode('utf-8')).get('error', {}).get('message', '')
        except Exception: msg = ''
        raise LoiAI({401: 'Khóa API không hợp lệ.', 403: 'Khóa API không có quyền dùng mô hình này.', 429: 'Vượt giới hạn hoặc hết hạn mức – thử lại sau.',
                     529: 'Máy chủ AI đang quá tải – thử lại sau.'}.get(e.code, 'Lỗi API %s.' % e.code) + (' ' + msg if msg else ''))
    except ssl.SSLError as e:
        raise LoiAI('Lỗi chứng chỉ bảo mật khi kết nối máy chủ AI: %s' % e)
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        raise LoiAI('Không kết nối được tới máy chủ AI (kiểm tra mạng): %s' % e)
    chu = ''.join(x.get('text', '') for x in d.get('content', []) if x.get('type') == 'text').strip()
    if not chu: raise LoiAI('AI không trả về nội dung.')
    if d.get('stop_reason') == 'max_tokens': raise LoiAI('Bản AI bị cắt giữa chừng (quá dài) – giữ bản không mạng.')
    return chu


# ------------------------------------------------------------------ đọc kết quả → khối
def doc_md(chu, tieu_de):
    B = [('h1', tieu_de)]; ul = None; bang = None
    def dong_lai():
        nonlocal ul, bang
        if ul: B.append(('ul', ul)); ul = None
        if bang:
            if len(bang) >= 1: B.append(('bang', bang[0], [r + [''] * (len(bang[0]) - len(r)) for r in bang[1:]]))
            bang = None
    for d in chu.splitlines():
        s = d.strip()
        if not s: dong_lai(); continue
        if s.startswith('|'):
            o = [x.strip() for x in s.strip('|').split('|')]
            if all(re.fullmatch(r':?-{2,}:?', x) for x in o if x): continue
            if ul: B.append(('ul', ul)); ul = None
            bang = (bang or []) + [o]; continue
        if bang: dong_lai()
        if s.startswith('#'):
            dong_lai(); B.append(('h2', s.lstrip('#').strip()))
        elif re.match(r'^[-*•]\s+', s):
            ul = (ul or []) + [re.sub(r'^[-*•]\s+', '', s)]
        else:
            if ul: B.append(('ul', ul)); ul = None
            B.append(('p', s))
    dong_lai()
    return B


# ------------------------------------------------------------------ bộ kiểm tra
def kiem_tra(B, hs, pd):
    """Đối chiếu bản AI với dữ kiện → danh sách lỗi (rỗng = dùng được)."""
    from battu_core.tong_hop.viet_ban import _chu as _gop
    loi = []; t = _gop(B); bm = hs['ban_menh']
    can_muc = ['Tóm tắt nhanh', '1. Lá số', '2. Bản mệnh', '3. Tính cách', '4. Học vấn', '5. Tài lộc', '6. Hôn nhân', '7. Cha mẹ',
               '8. Sức khỏe', '9. Đại vận', '10. Hóa giải']
    h2 = [k[1] for k in B if k[0] == 'h2']
    for m in can_muc:
        if not any(h.startswith(m) for h in h2): loi.append('Thiếu mục "%s".' % m)
    if hs['la_so']['tu_tru'] not in t and hs['la_so']['tu_tru'].replace(' · ', ', ') not in t:
        loi.append('Không thấy đúng tứ trụ %s.' % hs['la_so']['tu_tru'])
    for vai, h in (('dụng thần', bm['hanh_dung']), ('hỷ thần', bm['hanh_hy']), ('kỵ thần', bm['hanh_ky'])):
        sai = {x for x in re.findall(r'(?i)%s\s+(?:là\s+)?(Mộc|Hỏa|Thổ|Kim|Thủy)' % vai, t)} - {h}
        if sai: loi.append('%s ghi %s, đúng là %s.' % (vai.capitalize(), ', '.join(sorted(sai)), h))
    bang = [k for k in B if k[0] == 'bang']
    dvb = next((b for b in bang if b[1] and b[1][0].lower().startswith('đại vận')), None)
    lnb = next((b for b in bang if b[1] and 'tuổi mụ' in b[1][0].lower()), None)          # bảng lưu niên (không lẫn bảng sức khỏe)
    if not dvb: loi.append('Thiếu bảng đại vận.')
    else:
        muc = {d['ten']: d for d in pd['dai_van']}
        for r in dvb[2]:
            m = re.match(r'\s*(\S+ \S+)\s*\((\d{4})\s*[–-]\s*(\d{4})\)', r[0])
            if not m: continue
            d = muc.get(m.group(1))
            if not d: loi.append('Đại vận lạ: %s.' % r[0]); continue
            if (int(m.group(2)), int(m.group(3))) != (d['nam_bat_dau'], d['nam_ket_thuc']): loi.append('Năm của đại vận %s sai.' % d['ten'])
            mu = next((o.strip('* ') for o in r[1:] if o.strip('* ') in ('Rất tốt', 'Tốt', 'Trung bình', 'Xấu', 'Rất xấu')), None)
            if mu and mu != d['muc']: loi.append('Mức đại vận %s ghi "%s", đúng là "%s".' % (d['ten'], mu, d['muc']))
    if not lnb: loi.append('Thiếu bảng lưu niên.')
    else:
        muc = {y['nam']: y for y in pd['luu_nien']}; thay = set()
        for r in lnb[2]:
            m = re.match(r'\s*(\d{4})\s+(\S+ \S+)', r[0])
            if not m: continue
            y = muc.get(int(m.group(1)))
            if not y: continue
            thay.add(y['nam'])
            if m.group(2) != y['can_chi']: loi.append('Năm %s ghi %s, đúng là %s.' % (y['nam'], m.group(2), y['can_chi']))
            mu = next((o.strip('* ') for o in r[1:] if o.strip('* ') in ('Rất tốt', 'Tốt', 'Trung bình', 'Xấu', 'Rất xấu')), None)
            if mu != y['muc']: loi.append('Mức năm %s ghi "%s", đúng là "%s".' % (y['nam'], mu, y['muc']))
        thieu = sorted(set(muc) - thay)
        if thieu: loi.append('Bảng lưu niên thiếu năm %s.' % ', '.join(map(str, thieu)))
    if len(t) < 6000: loi.append('Bản AI quá ngắn (%d ký tự).' % len(t))
    return loi


# ------------------------------------------------------------------ lưu đệm
def khoa_dem(hs, pd, mo_hinh):
    """Khóa lưu đệm: băm hồ sơ gọn + mô hình + phiên bản lời nhắc (đổi lá số / quy tắc → gọi lại)."""
    s = json.dumps([ho_so_gon(hs, pd), mo_hinh, PHIEN_BAN_LOI_NHAC], ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(s.encode('utf-8')).hexdigest()[:24]


def thu_muc_dem():
    p = os.path.join(os.path.expanduser('~'), '.battu_thieu_vi_hoa', 'ai')
    os.makedirs(p, exist_ok=True)
    return p


def doc_dem(khoa, thu_muc=None):
    p = os.path.join(thu_muc or thu_muc_dem(), khoa + '.md')
    return open(p, encoding='utf-8').read() if os.path.exists(p) else None


def ghi_dem(khoa, chu, thu_muc=None):
    p = os.path.join(thu_muc or thu_muc_dem(), khoa + '.md')
    open(p, 'w', encoding='utf-8').write(chu)
    return p


def nang_cao(khoa_api, hs, pd, B, tieu_de, mo_hinh=MO_HINH_MAC_DINH, thu_muc=None, _mo=None):
    """Toàn bộ quy trình → (khối AI hoặc None, [lỗi kiểm tra], từ_đệm). Lỗi mạng/API → LoiAI."""
    k = khoa_dem(hs, pd, mo_hinh); chu = doc_dem(k, thu_muc); tu_dem = chu is not None
    if chu is None: chu = goi_api(khoa_api, hs, pd, B, mo_hinh, _mo=_mo)
    BA = doc_md(chu, tieu_de); loi = kiem_tra(BA, hs, pd)
    if not loi and not tu_dem: ghi_dem(k, chu, thu_muc)
    BA.insert(1, ('nho', 'Bản do AI (%s) viết lại từ dữ kiện phần mềm đã tính; bộ kiểm tra đã đối chiếu tứ trụ, dụng – hỷ – kỵ, '
                         'đại vận và mức từng năm.' % mo_hinh))
    return (BA if not loi else None), loi, tu_dem
