# -*- coding: utf-8 -*-
"""T12 – Quan hệ can chi và hồ sơ dữ kiện (đề xuất hoàn thiện 15/11/2026, giai đoạn T2).
  (a) Đối chiếu battu_core.tong_hop.quan_he với phần hợp/xung workbook ĐÃ tính (Engine = workbook, kiểm ở T1/T2/T7),
      trên 501 lá (cases501_a/b) + 20 lá T7, cho từng cặp trụ kề nhau (năm–tháng, tháng–ngày, ngày–giờ):
        • TRUNG_GIAN dòng 55 (bán hợp: hành) = bán hợp của quan_he
        • TRUNG_GIAN dòng 64 (lục hợp: hành) = lục hợp của quan_he
        • PP2_DINH_LUONG dòng 31 = 1 (lục xung thành lập) ⇒ quan_he có lục xung cặp đó
        • TRUNG_GIAN dòng 52 = 1 (chi thuộc cục 3 chi) ⇒ quan_he có tam hợp / tam hội chứa chi đó
  (b) Bảng tra: đủ 5 ngũ hợp, 6 lục hợp, 6 xung, 6 hại, 6 phá; tính đối xứng.
  (c) Lá mẫu Mậu Ngọ · Quý Hợi · Giáp Ngọ · Mậu Thìn: có Mậu–Quý hợp, Ngọ–Ngọ tự hình, ám hợp Hợi–Ngọ, Thìn kho Thủy;
      năm 2032 Nhâm Tí xung hai Ngọ (chạm dụng thần, chạm cung phu thê); 2029 Kỷ hợp Giáp, Dậu hợp Thìn.
  (d) Hồ sơ dữ kiện của 20 lá T7 dựng được, ghi được JSON, đủ 4 trụ, ≥ 8 đại vận, tối đa 11 năm lưu niên liên tục từ năm xem.
Dùng: python -m tests.t12_quan_he"""
import sys, os, json, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine
from battu_core.tong_hop import quan_he as Q
from tests.so_dap_an import oracle_nhap

GOC = os.path.dirname(os.path.abspath(__file__))
CAP = [(0, 1), (1, 2), (2, 3)]


def _hanh(qh, loai, i, j):
    for q in qh:
        if q['loai'] == loai and q['vi_tri'] == ['chi ' + Q.TRU[i], 'chi ' + Q.TRU[j]]: return Q.HANH.index(q['hanh'])
    return -1


def _doi_chieu(cases):
    n = 0; loi = []; dem = {'ban_hop': 0, 'luc_hop': 0, 'xung': 0, 'cuc3': 0}
    for c, nhap in oracle_nhap(cases):
        E = Engine({'NHAP_LIEU': nhap}, None); M = E.M
        tg, pp2, nl = M.si['TRUNG_GIAN'], M.si['PP2_DINH_LUONG'], M.si['NHAP_LIEU']
        tru = [(E.raw(nl, 37, k) or E.raw(nl, 7, k), E.raw(nl, 38, k) or E.raw(nl, 8, k)) for k in range(3, 7)]
        if not all(isinstance(a, str) and a and isinstance(b, str) and b for a, b in tru): continue
        qh = Q.trong_tu_tru(tru); n += 1
        for k, (i, j) in enumerate(CAP):
            col = 3 + k
            for loai, dong in (('ban_hop', 55), ('luc_hop', 64)):
                wb = int(E.raw(tg, dong, col)); mine = _hanh(qh, loai, i, j)
                if wb != -1 or mine != -1: dem[loai] += 1
                if wb != mine: loi.append((c['id'], loai, Q.TRU[i] + '–' + Q.TRU[j], tru, 'workbook %d, quan_he %d' % (wb, mine)))
            if E.raw(pp2, 31, col) == 1.0:
                dem['xung'] += 1
                if not any(q['loai'] == 'xung' and q['vi_tri'] == ['chi ' + Q.TRU[i], 'chi ' + Q.TRU[j]] for q in qh):
                    loi.append((c['id'], 'xung', Q.TRU[i] + '–' + Q.TRU[j], tru, 'workbook có xung, quan_he không'))
        for i in range(4):
            if E.raw(tg, 52, 3 + i) == 1.0:
                dem['cuc3'] += 1
                if not any(q['loai'] in ('tam_hop', 'tam_hoi') and ('chi ' + Q.TRU[i]) in q['vi_tri'] for q in qh):
                    loi.append((c['id'], 'cuc3', Q.TRU[i], tru, 'workbook có cục 3 chi, quan_he không'))
    return n, dem, loi


def _bang():
    loi = []
    for ten, bang, so in (('ngũ hợp', Q.NGU_HOP, 5), ('lục hợp', Q.LUC_HOP, 6), ('xung', Q.LUC_XUNG, 6), ('hại', Q.LUC_HAI, 6), ('phá', Q.LUC_PHA, 6)):
        if len(bang) != so: loi.append('%s: %d cặp (cần %d)' % (ten, len(bang), so))
    for a in Q.CHI:
        for b in Q.CHI:
            if sorted(x[0] for x in Q._cap_chi(a, b)) != sorted(x[0] for x in Q._cap_chi(b, a)): loi.append('chi không đối xứng %s %s' % (a, b))
    for a in Q.CAN:
        for b in Q.CAN:
            if sorted(x[0] for x in Q._cap_can(a, b)) != sorted(x[0] for x in Q._cap_can(b, a)): loi.append('can không đối xứng %s %s' % (a, b))
    return loi


def _la_mau():
    tru = [('Mậu', 'Ngọ'), ('Quý', 'Hợi'), ('Giáp', 'Ngọ'), ('Mậu', 'Thìn')]
    qh = Q.danh_dau(Q.trong_tu_tru(tru), tru, 'Hỏa', 'Thổ', 'Thủy'); ten = [q['ten'] for q in qh]; loi = []
    for can in ('ngũ hợp Mậu–Quý', 'tự hình Ngọ–Ngọ', 'ám hợp Hợi–Ngọ (Đinh–Nhâm)', 'Thìn là kho Thủy'):
        if can not in ten: loi.append('lá mẫu thiếu: ' + can)
    if 'Mậu khắc Quý' in ten: loi.append('lá mẫu: cặp ngũ hợp không được luận khắc')
    y = Q.danh_dau(Q.voi_tu_tru(tru, 'Nhâm', 'Tí', 'năm 2032'), tru, 'Hỏa', 'Thổ', 'Thủy')
    xung = [q for q in y if q['loai'] == 'xung']
    if len(xung) != 2 or not all(q['cham_dung'] for q in xung) or not any(q['cham_cung_phu_the'] for q in xung):
        loi.append('2032: cần 2 lục xung Tí–Ngọ chạm dụng thần, có chạm cung phu thê')
    y = [q['ten'] for q in Q.voi_tu_tru(tru, 'Kỷ', 'Dậu', 'năm 2029')]
    for can in ('ngũ hợp Kỷ–Giáp', 'lục hợp Dậu–Thìn'):
        if can not in y: loi.append('2029 thiếu: ' + can)
    return loi


def _ho_so(cases):
    from battu_core.tong_hop.du_kien import ho_so
    loi = []; n = 0
    for c, nhap in oracle_nhap(cases):
        E = Engine({'NHAP_LIEU': nhap}, None); E.tinh_tat_ca()
        try:
            h = json.loads(json.dumps(ho_so(E), ensure_ascii=False)); n += 1
        except Exception as e:
            loi.append((c['id'], 'hồ sơ lỗi: %r' % e)); continue
        if len(h['tu_tru']) != 4 or not all(t['can'] and t['chi'] for t in h['tu_tru']): loi.append((c['id'], 'tứ trụ thiếu'))
        if len(h['dai_van']) < 8: loi.append((c['id'], 'đại vận < 8'))
        nam = [y['nam'] for y in h['luu_nien']]          # năm xem ngoài 100 năm của lá (lá sinh 1900/2100) → ít hơn 11 năm
        if len(nam) > 11 or nam != list(range(nam[0], nam[0] + len(nam))) if nam else False:
            loi.append((c['id'], 'lưu niên không liên tục: %s' % nam))
        if not h['ban_menh']['hanh_dung']: loi.append((c['id'], 'không đọc được dụng thần'))
    return n, loi


def chay():
    t0 = time.time()
    cases = json.load(open(os.path.join(GOC, 'cases501_a.json'), encoding='utf-8')) + json.load(open(os.path.join(GOC, 'cases501_b.json'), encoding='utf-8'))
    t7 = json.load(open(os.path.join(GOC, 't7_ngau_nhien.json'), encoding='utf-8'))
    n, dem, loi = _doi_chieu(cases + t7)
    lb = _bang(); lm = _la_mau(); nh, lh = _ho_so(t7)
    return n, dem, loi, lb, lm, nh, lh, time.time() - t0


if __name__ == '__main__':
    n, dem, loi, lb, lm, nh, lh, dt = chay()
    print('T12 quan hệ can chi – %.1f giây' % dt)
    print('  (a) đối chiếu workbook: %d lá; cặp có bán hợp %d, lục hợp %d, xung %d, chi thuộc cục 3 chi %d – %s' % (
        n, dem['ban_hop'], dem['luc_hop'], dem['xung'], dem['cuc3'], 'ĐẠT' if not loi else 'LỆCH %d' % len(loi)))
    for x in loi[:15]: print('     ', x)
    print('  (b) bảng tra: %s' % ('ĐẠT' if not lb else lb))
    print('  (c) lá mẫu: %s' % ('ĐẠT' if not lm else lm))
    print('  (d) hồ sơ dữ kiện %d lá: %s' % (nh, 'ĐẠT' if not lh else lh[:10]))
    sys.exit(1 if (loi or lb or lm or lh) else 0)
