# -*- coding: utf-8 -*-
"""Làm sạch chữ HIỂN THỊ (quyết định 14/11/2026): phần mềm độc lập với workbook, nên mọi tham chiếu nội bộ của workbook
(tên sheet, địa chỉ ô, "dòng N", "cột X", mã giả định G…, mã quy tắc nội bộ) được bỏ khỏi giao diện, PDF, XLSX.
Lõi tính KHÔNG đổi: giá trị gốc vẫn trùng workbook 100% (kiểm thử và JSON dùng giá trị gốc).
Chỉ bỏ phần tham chiếu; nội dung luận đoán và trích dẫn sách (P1 tr…, P2 tr…, HG tr…, UH tr…) giữ nguyên."""
import re
from battu_core.du_lieu import model

# Tên hiển thị khi một sheet được nhắc như một phần nội dung (không kèm địa chỉ ô)
TEN = {'KET_LUAN': 'Luận giải chung', 'LUAN_GIAI': 'Luận giải chi tiết', 'DOC_NHANH': 'Đọc nhanh', 'DOI_CHIEU': 'phần Đối chiếu',
       'DAI_VAN': 'Đại vận', 'LUU_NIEN': 'Lưu niên', 'THEO_DOI': 'Theo dõi', 'SU_NGHIEP': 'Sự nghiệp', 'SUC_KHOE': 'Sức khỏe',
       'HON_NHAN': 'Hôn nhân', 'LUC_THAN': 'Lục thân', 'MENH_CUC': 'Mệnh cục', 'TAI_LOC': 'Tài lộc', 'TO_CHAT': 'Tố chất',
       'THAN_SAT': 'Thần sát', 'THONG_CAN': 'Thông căn', 'PP1_DINH_TINH': 'PP1', 'PP2_DINH_LUONG': 'PP2',
       'DUNG_HY_KY': 'Dụng – hỷ – kỵ', 'KHONG_RO_GIO': 'Không rõ giờ', 'NGU_HANH_THIEU': 'Ngũ hành thiếu',
       'NHAP_LIEU': 'Nhập liệu', 'SU_KIEN': 'Kiểm chứng sự việc', 'THAM_KHAO': 'Tham khảo', 'TIET_KHI': 'bảng tiết khí',
       'TRUNG_GIAN': 'Trung gian', 'BAO_PHU': 'Bảo phù', 'TRANG_CHU': 'Trang chủ', 'HUONG_DAN': 'Hướng dẫn',
       'MAU_THUAN': 'ghi chú giả định', 'KIEM_THU': 'bộ lá kiểm thử', 'CSDL': 'cơ sở dữ liệu quy tắc', 'SAO': 'Bảng sao'}
_SH = sorted([s for s in model().sheets if s != 'SAO'], key=len, reverse=True)
SH = r'(?:' + '|'.join(map(re.escape, _SH)) + r')'
SH_SAO = r'(?:' + SH + r'|SAO)'                                     # "SAO" chỉ coi là sheet khi trong danh sách [..] hoặc kèm ô
O = r'\$?[A-Z]{1,3}\$?\d{1,4}(?::\$?[A-Z]{1,3}\$?\d{1,4})?'          # địa chỉ ô / vùng
CHI_DAN = (r'(?:\s*(?:!' + O + r'|\s' + O + r'(?![\wÀ-ỹ])|\s(?:dòng|cột|mục|bảng)\s+[A-ZĐ0-9][\w]*(?:\s*[–\-,]\s*[A-ZĐ0-9][\w]*)*'
           r'|\s(?:từ\s+)?dòng\s+\d+(?:\s*[–\-]\s*\d+)?))*')
AN = {'MAU_THUAN', 'KIEM_THU', 'CSDL'}                              # chỉ tính ẩn – không có trên giao diện
_MUC_KL = re.compile(r'xem KET_LUAN(?:\s+dòng\s+\d+)?\s*\(mục ([A-K])\)')
_TRO = re.compile(r'\b(xem|chi tiết(?: và cách chấm)?:?|đủ ở|lấy từ|vào|ở|tra|trong)\s+(?:sheet\s+|bảng\s+)?(' + SH + r')' + CHI_DAN +
                  r'((?:\s*(?:,|và|/)\s*(?:' + SH + r')' + CHI_DAN + r')*)')
_SHEET_O = re.compile(r'(?:\bsheet\s+)?(' + SH + r')((?:!|\s)' + O + r'|\s(?:dòng|cột|mục|bảng)\s+[A-ZĐ0-9][\w]*(?:\s*[–\-,]\s*[A-ZĐ0-9][\w]*)*)' + CHI_DAN)
_SHEET = re.compile(r'(?:\bsheet\s+)?\b(' + SH + r')\b')
_DS_SHEET = re.compile(r'\s*[\(\[]\s*' + SH_SAO + r'(?:\s*(?:,|/|và)\s*' + SH_SAO + r')*\s*[\)\]]')
_O_DON = re.compile(r'\s*\((?:ô\s+)?' + O + r'(?:\s*[,;]\s*' + O + r')*\)')
_DONG_DON = re.compile(r'\s*\((?:xem\s+)?(?:dòng|cột)\s+[\w–\-,/ …]+\)')
_MA = re.compile(r'\s*\((?:UH|UT|TC|TK|MK|G)\d+[a-z]?(?:\s*,\s*(?:UH|UT|TC|TK|MK|G)\d+[a-z]?)*\)')
_GIA_DINH = re.compile(r'\s*[–\-]?\s*(?:giả định\s+)\bG\d{1,3}\b')


def _tro(m):
    ds = [m.group(2)] + re.findall(SH, m.group(3) or '')
    ten = [TEN[x] for x in ds if x not in AN]
    if not ten: return ''
    dong = m.group(1)
    if dong == 'lấy từ': return ''
    return '%s %s' % (dong if dong.endswith(':') or dong in ('vào', 'ở', 'trong', 'tra') else dong, ', '.join(dict.fromkeys(ten)))


_MA_QT = re.compile(r'^(?!P\d$|PP\d$|PA\d$)(?:[A-Z]{1,3}\d{1,3}[a-z]?|G\d{1,3})$')            # mã quy tắc / giả định nội bộ: TC66, CM14, UH3b, G17


def _bo_ma(m):
    mo, nd, dong = m.group(1), m.group(2), m.group(3)
    phan = [x.strip() for x in re.split(r'[,;]', nd)]
    if not any(_MA_QT.match(x) for x in phan): return m.group(0)
    con = [x for x in phan if x and not _MA_QT.match(x)]
    return (' ' + mo + '; '.join(con) + dong) if con else ''


def lam_sach(t):
    if not isinstance(t, str) or not t: return t
    s = _MUC_KL.sub(r'xem mục \1 (Theo lĩnh vực)', t)
    for _ in range(2):                                                # ghi chú kỹ thuật trong ngoặc: Excel, sheet, ô, ngày chốt
        s = re.sub(r'\s*\([^()]*(?:\bExcel\b|\bsheet\b|\bô\s+[A-Z]{1,2}\d|\bchốt\s+\d)[^()]*\)', '', s)
    s = re.sub(r'\bcác sheet\b', 'các phần', s); s = re.sub(r'\bsheet này\b', 'phần này', s)
    s = _GIA_DINH.sub('', s)
    s = _DS_SHEET.sub('', s)                                          # [TAI_LOC, MENH_CUC]  (DOI_CHIEU)
    s = re.sub(r'\s*\(\s*(?:sheet\s+)?' + SH + CHI_DAN + r'\s*\)', '', s)      # (DOI_CHIEU C33)  (MENH_CUC dòng 16-22)
    s = _TRO.sub(_tro, s)                                             # xem SU_NGHIEP mục B → xem Sự nghiệp; xem MAU_THUAN M14 → bỏ
    def _so(m, s_=None):                                              # PP1_DINH_TINH!C91, LUU_NIEN cột BN
        trc = m.string[max(0, m.start() - 2):m.start()]
        return TEN[m.group(1)] if (':' in trc and m.group(1) not in AN) else ''
    s = _SHEET_O.sub(_so, s)
    s = _SHEET.sub(lambda m: '' if m.group(1) in AN else TEN[m.group(1)], s)
    s = _O_DON.sub(lambda m: m.group(0) if re.search(r'\b(?:P|PP|PA)\d\b', m.group(0)) else '', s)
    s = _DONG_DON.sub('', s); s = _MA.sub('', s)
    s = re.sub(r'\s*([\[(])([^\[\]()]*)([\])])', _bo_ma, s)                 # [TC66, P2 tr.577] → [P2 tr.577]
    s = re.sub(r'\s*\((?:xem\s+)?[A-H]\d{2,3}\)', '', s)                     # (xem D28)
    s = re.sub(r'\b[A-H]\d{2,3}:\s*', '', s)                                  # D190: …
    s = re.sub(r'(?<=\w)\s+[A-H]\d{2,3}\b', '', s)                            # số ngày tự nhập C28
    s = re.sub(r'\s*[–\-]\s*G\d{1,3}\b', '', s)                                # – G34
    s = re.sub(r'\b(PP1|PP2): \1\b', r'\1', s)
    s = re.sub(r'\s*\((?:ô\s+)?[A-H]\d{1,3}\s*=\s*[^()]*\)', '', s)                 # (C13 = Tự nhập)
    s = re.sub(r'\s*\(cột\s+[A-Z](?:[^()=]*)\)', '', s)                         # (cột I, J… ngoài 8 vận …)
    s = re.sub(r'\s*\((?:xem\s+)?dòng\s+\d+(?:/\d+)+\)', '', s)                  # (dòng 54/69)
    s = re.sub(r'(?<![\w])(?!PP\d|PA\d)(?:[A-Z]{2}\d{1,3}[a-z]?)(?:[.,;]|\b)\s*', '', s)   # mã quy tắc nội bộ: HN31. / TC21,
    s = re.sub(r'\bG\d{2,3}\b[.,;:]?\s*', '', s)                                  # mã giả định G30
    s = re.sub(r'TÓM TẮT\s+dòng\s+\d+(?:\s*[–\-]\s*\d+)?', 'Luận giải chung', s)
    s = re.sub(r'\s+(?:ở|xem)\s+dòng\s+\d+(?:\s*[–\-]\s*\d+)?(?:\s+(?:và|,))?', '', s)
    s = re.sub(r'\s+(?:từ\s+)?dòng\s+\d+(?:\s*[–\-]\s*\d+)?(?:\s*,\s*\d+(?:\s*[–\-]\s*\d+)?)*', '', s)
    s = re.sub(r'\s+cột\s+[A-Z]{1,2}(?:\s*[–\-,]\s*[A-Z]{1,2})*\b', '', s)
    # dọn dấu câu thừa
    s = re.sub(r'\(\s*[,;:–\-]?\s*\)', '', s); s = re.sub(r'\[\s*\]', '', s)
    s = re.sub(r'\b(PP1|PP2):\s*(?=[·;|]|$)', '', s)
    s = re.sub(r'[ \t]{2,}', ' ', s); s = re.sub(r'\s+([.,;:)])', r'\1', s); s = re.sub(r'\(\s+', '(', s)
    s = re.sub(r'[;,]\s*\)', ')', s); s = re.sub(r'[–\-]\s*([.;)])', r'\1', s); s = re.sub(r':\s*\.', '.', s)
    s = re.sub(r'\.\s*\.', '.', s); s = re.sub(r'·\s*·', '·', s); s = re.sub(r'^\s*·\s*|\s*·\s*$', '', s)
    s = re.sub(r'[ \t]+\n', '\n', s)
    if not t[:1].isspace(): s = s.lstrip()
    if not t[-1:].isspace(): s = s.rstrip()
    return s


def hien_thi_sach(E, sheet, ref):
    """Chữ hiển thị trên màn hình: bỏ tham chiếu nội bộ, hiện đầy đủ chữ bị workbook rút gọn (export.khoi_phuc_cat),
    ghi kèm năm dương lịch cho khoảng tuổi đại vận (export.ghi_tuoi)."""
    from battu_core.hien_thi import hien_thi_o
    from battu_core.du_lieu import rc
    from export.khoi_phuc_cat import khoi_phuc
    from export.ghi_tuoi import them_nam
    r, c = rc(ref)
    return lam_sach(them_nam(E, khoi_phuc(E, sheet, r, c, hien_thi_o(E, sheet, ref))))
