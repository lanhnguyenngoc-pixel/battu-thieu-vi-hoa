# -*- coding: utf-8 -*-
"""Bản 6.10 (B1, người dùng yêu cầu làm 23/11/2026): hợp cục của chi đại vận / chi lưu niên với chi tứ trụ (bán hợp, tam hợp, tam hội).
Căn cứ: P2 tr.343 (bán hợp cần hai chi kề nhau; tam hội, tam hợp đủ ba chi không cần kề; hợp hóa chỉ thành khi có thiên can cùng hành hóa thần "dẫn",
nếu không thì "hợp mà không hóa"); P1 tr.62, 66-68 (thứ bậc tam hội > tam hợp > bán hợp). Bằng chứng ngoài bộ sách chính: "Một trăm cách đoán mệnh"
VD1 (Giáp Dần: Dần–Ngọ), VD3 (Nhâm Thìn: Tí–Thìn, can Nhâm dẫn), VD6 (Giáp Tuất: Dần Ngọ Tuất).
SUY LUẬN (cần người dùng duyệt): chi đại vận không có "vị trí kề" trong tứ trụ nên coi chi vận kề mọi chi tứ trụ; bán hợp phải có chi vượng (Tí, Mão, Ngọ, Dậu).
CAN_DAN=True: hợp cục chỉ làm đổi hành chi vận (dòng 130) khi có can dẫn (theo P2 tr.343). CAN_DAN=False: không đòi can dẫn (cách tác giả sách một trăm cách dùng ở VD1, VD6)."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

GIO = 'NHAP_LIEU!$C$9<>"Không rõ giờ"'


def _has(c):
    return '(COUNTIF(TRUNG_GIAN!$C$6:$E$6,%s)+IF(AND(%s,TRUNG_GIAN!$F$6=%s),1,0)>0)' % (c, GIO, c)


def hop_cuc_van(p, why, can_dan=True, doi_muc=True):
    D = 'DAI_VAN'
    cols = 'CDEFGHIJKL'
    lab = {140: 'G. HỢP CỤC CỦA CHI VẬN VỚI CHI TỨ TRỤ (bán hợp, tam hợp, tam hội – P2 tr.343; P1 tr.62, 66–68) – bản 6.10',
           141: 'Tam hội đủ ba chi: chi vận cùng hai chi tứ trụ cùng phương (1/0)',
           142: 'Tam hợp đủ ba chi: chi vận cùng hai chi tứ trụ cùng cục (1/0)',
           143: 'Chi tứ trụ bán hợp với chi vận (−1 = không; cần có chi vượng Tí, Mão, Ngọ, Dậu trong cặp)',
           144: 'Loại hợp cục của chi vận (0 không, 1 bán hợp, 2 tam hợp, 3 tam hội)',
           145: 'Hành của hợp cục (−1 = không)',
           146: 'Có can dẫn: can cùng hành hóa thần lộ ở tứ trụ hoặc là can vận (1/0)',
           147: 'Hóa thần của chi vận do hợp cục (−1 = không hóa; hợp mà không hóa thì chi chỉ giữ hành riêng, P1 tr.62)' + ('' if can_dan else ' – bản không đòi can dẫn') + (' – chỉ tam hợp/tam hội có can dẫn mới đổi mức vận; bán hợp chỉ ghi' if doi_muc == 'tam' else '' if doi_muc else ' – CHỈ GHI, không đổi mức vận (đo 23/11/2026: cho hợp cục đổi mức làm khớp đại vận Thiệu Vĩ Hoa giảm 76 → 72 trên 171 vận)'),
           148: 'Ghi chú hợp cục của chi vận'}
    for r, t in lab.items():
        p.set_text(D, 'B%d' % r, t, style_from='B130', why=why, must_be_empty=True)
    for c in cols:
        X = '%s$34' % c; A = 'MOD(%s+4,12)' % X; B = 'MOD(%s+8,12)' % X
        U = 'INT(MOD(%s+1,12)/3)' % X
        pos = ['MOD(3*%s-1+%d,12)' % (U, j) for j in range(3)]
        # tam hội: cả ba chi của phương, chi vận là một trong ba; hai chi còn lại có trong tứ trụ
        hoi = '+'.join('IF(OR(%s=%s,%s),1,0)' % (x, X, _has(x)) for x in pos)
        p.new_formula(D, '%s141' % c, 'IF((%s)=3,1,0)' % hoi, style_from='C130', why=why)
        p.new_formula(D, '%s142' % c, 'IF(AND(%s,%s),1,0)' % (_has(A), _has(B)), style_from='C130', why=why)
        bh = ('IF(AND(%s,NOT(%s),OR(MOD(%s,3)=0,MOD(%s,3)=0)),%s,IF(AND(%s,NOT(%s),OR(MOD(%s,3)=0,MOD(%s,3)=0)),%s,-1))'
              % (_has(A), _has(B), X, A, A, _has(B), _has(A), X, B, B))
        p.new_formula(D, '%s143' % c, bh, style_from='C130', why=why)
        p.new_formula(D, '%s144' % c, 'IF(%s141=1,3,IF(%s142=1,2,IF(%s143>=0,1,0)))' % (c, c, c), style_from='C130', why=why)
        p.new_formula(D, '%s145' % c, 'IF(%s144=3,CHOOSE(%s+1,4,0,1,3),IF(%s144>=1,CHOOSE(MOD(%s,4)+1,4,3,1,0),-1))' % (c, U, c, X), style_from='C130', why=why)
        lo = '(COUNTIF(TRUNG_GIAN!$C$7:$E$7,%s145)+IF(AND(%s,TRUNG_GIAN!$F$7=%s145),1,0)+IF(%s$42=%s145,1,0)>0)' % (c, GIO, c, c, c)
        p.new_formula(D, '%s146' % c, 'IF(%s145<0,0,IF(%s,1,0))' % (c, lo), style_from='C130', why=why)
        dk = '%s144>0' % c if not can_dan else 'AND(%s144>0,%s146=1)' % (c, c)
        p.new_formula(D, '%s147' % c, 'IF(%s,%s145,-1)' % (dk, c), style_from='C130', why=why)
        chi = lambda x: 'INDEX(CHI_TEN,%s+1)' % x
        ban = 'IF(%s144=1,chi_ban,IF(%s144=2,%s&", "&%s,%s&", "&%s&", "&%s))' % (c, c, chi(A), chi(B), chi(pos[0]), chi(pos[1]), chi(pos[2]))
        ban = ban.replace('chi_ban', chi('%s143' % c))
        hn = 'INDEX(HANH_TEN,%s145+1)' % c
        cho = '", chờ lưu niên có can "&' + hn
        note = ('IF(%s144=0,"","Chi vận "&%s&" "&CHOOSE(%s144,"bán hợp với chi ","tam hợp với hai chi ","hội đủ ba chi ")&%s&" thành "&%s&" cục"&'
                'IF(%s146=1," – có can "&%s&" dẫn: hợp hóa thành, lấy "&%s&" xét cát hung",'
                '" – chưa có can "&%s&" dẫn: hợp mà không hóa"&%s)&" (P2 tr.343; P1 tr.62, 66–68; chi vận coi như kề các chi tứ trụ – suy luận)")'
                % (c, chi(X), c, ban, hn, c, hn, hn, hn, cho))
        p.new_formula(D, '%s148' % c, note, text=True, style_from='C130', why=why)
    # dòng 130: nếu không có lục hợp hóa thì lấy hợp cục
    for c in cols:
        f = p.formula(D, '%s130' % c); assert f.endswith(',-1))))'), (c, f[-30:])
        if doi_muc == 'tam': p.set_formula(D, '%s130' % c, f[:-len(',-1))))')] + ',IF(%s144>=2,%s147,-1)))))' % (c, c), why=why)
        elif doi_muc: p.set_formula(D, '%s130' % c, f[:-len(',-1))))')] + ',%s147))))' % c, why=why)
    # câu giải thích dòng 138: nối câu hợp cục
    for c in cols:
        f = p.formula(D, '%s138' % c)
        p.set_formula(D, '%s138' % c, f + '&IF(%s144>0," Hợp cục: "&%s148&".","")' % (c, c), text=True, why=why)
    # dòng 92 tóm tắt sự kiện của vận: nối câu hợp cục (chỉ ghi)
    for c in cols:
        f = p.formula(D, '%s92' % c)
        p.set_formula(D, '%s92' % c, f + '&IF(LEN(%s148)>0," • "&%s148&". ","")' % (c, c), text=True, why=why)


def ban_hop_luu_nien(p, why):
    """Bán hợp của chi lưu niên với một chi tứ trụ / đại vận – chỉ GHI vào 'sự kiện của năm', không đổi xu thế (suy luận, như bán hợp đại vận)."""
    from step20_bo_sung import LN_ROWS
    from step21_con_cai import _cot
    L = 'LUU_NIEN'
    DU = 'IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)'
    HY = 'IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)'
    KY = 'IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)'
    c1, c2 = _cot(142), _cot(143)
    p.set_text(L, '%s29' % c1, 'Phụ (bán hợp lưu niên) – chi tứ trụ/đại vận bán hợp với chi năm', why=why, must_be_empty=True, style_from='DN29')
    p.set_text(L, '%s29' % c2, 'Phụ (bán hợp lưu niên) – câu ghi', why=why, must_be_empty=True, style_from='DN29')
    for r in LN_ROWS:
        K, G, P = 'K%d' % r, 'G%d' % r, 'P%d' % r

        def has_o(c):
            return '(COUNTIF(TRUNG_GIAN!$C$6:$E$6,%s)+IF(AND(%s,TRUNG_GIAN!$F$6=%s),1,0)+IF(%s=%s,1,0)>0)' % (c, GIO, c, G, c)
        A, B = 'MOD(%s+4,12)' % K, 'MOD(%s+8,12)' % K
        f = ('IF(AND(%s,NOT(%s),OR(MOD(%s,3)=0,MOD(%s,3)=0)),%s,IF(AND(%s,NOT(%s),OR(MOD(%s,3)=0,MOD(%s,3)=0)),%s,-1))'
             % (has_o(A), has_o(B), K, A, A, has_o(B), has_o(A), K, B, B))
        p.new_formula(L, '%s%d' % (c1, r), f, style_from='BN30', why=why)
        h = 'CHOOSE(MOD(%s,4)+1,4,3,1,0)' % K
        lo = '(COUNTIF(TRUNG_GIAN!$C$7:$E$7,%s)+IF(AND(%s,TRUNG_GIAN!$F$7=%s),1,0)+IF(%s=%s,1,0)>0)' % (h, GIO, h, P, h)
        vai = ('IF(%s<0,"",IF(OR(%s=%s,%s=%s)," – hành này là dụng/hỷ thần: tốt",IF(%s=%s," – hành này là kỵ thần: hung"," – hành này là nhàn thần")))'
               % (DU, h, DU, h, HY, h, KY))
        t = ('IF(%s%d<0,"","• Chi năm "&INDEX(CHI_TEN,%s+1)&" bán hợp chi "&INDEX(CHI_TEN,%s%d+1)&" (tứ trụ/đại vận) thành "&INDEX(HANH_TEN,%s+1)&" cục"&'
             'IF(%s," – có can "&INDEX(HANH_TEN,%s+1)&" dẫn: hợp hóa thành"," – chưa có can "&INDEX(HANH_TEN,%s+1)&" dẫn: hợp mà không hóa, chờ năm/vận có can "&INDEX(HANH_TEN,%s+1))&%s&'
             '" (P2 tr.343; P1 tr.62; bán hợp lưu niên là suy luận, chỉ ghi, không đổi xu thế). ")'
             % (c1, r, K, c1, r, h, lo, h, h, h, vai))
        p.new_formula(L, '%s%d' % (c2, r), t, text=True, style_from='BN30', why=why)
        old = p.formula(L, 'DY%d' % r); assert old.startswith('TRIM(') and old.endswith(')'), old[:40]
        p.set_formula(L, 'DY%d' % r, old[:-1] + '&%s%d)' % (c2, r), text=True, why=why)
