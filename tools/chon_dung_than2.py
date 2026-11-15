"""Mẫu thử DT-2: chọn dụng thần như cách tác giả vận dụng – phù ức là chính, xét lần lượt từng ứng viên."""
HT = ['Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy']
G = ['Tỉ kiếp', 'Thực thương', 'Tài', 'Quan sát', 'Ấn']
TK, TT, TA, QS, AN = range(5)
DONG = {'Hợi', 'Tí', 'Sửu'}; HA = {'Tị', 'Ngọ', 'Mùi'}

def ung_vien(nhuoc, a):
    """P1 tr.99 + P2 tr.380–381: danh sách ứng viên theo thứ tự."""
    if nhuoc:
        dom = max((TT, TA, QS), key=lambda i: a[i])
        return {QS: [AN, TK], TA: [TK, AN], TT: [AN, TK]}[dom], 'Thân nhược, %s nhiều nhất' % G[dom], dom
    dom = AN if a[AN] > a[TK] else TK
    return {AN: [TA, QS, TT], TK: [QS, TT, TA]}[dom], 'Thân vượng, %s nhiều' % G[dom], dom

def chon(nh, nhuoc, a, thang, yeu=0.05, dieu_hau=True):
    t = sum(a) or 1; p = [x / t for x in a]; B = []
    L, ly, dom = ung_vien(nhuoc, a)
    B.append(('1. Ứng viên theo sinh phù – áp chế (P1 tr.99; P2 tr.380–381)', '%s → xét lần lượt: %s.' % (ly, ' → '.join(G[g] for g in L)),
              '"…lấy ấn tinh làm dụng thần… Nếu không có ấn tinh thì lấy tỉ kiếp" (P1 tr.99)'))
    phe_minh = p[TK] + p[AN]
    hop = []
    for g in L:
        if p[g] <= yeu:
            B.append(('2. Xét %s' % G[g], '%s %s (%.0f%%) → không đủ lực, xét ứng viên sau.' % (G[g], 'khuyết' if p[g] == 0 else 'rất yếu', 100 * p[g]),
                      '"Nếu không có … thì lấy …" (P1 tr.99); dụng thần phải "kiện toàn, có lực" (P1 tr.97)'))
            continue
        if not nhuoc and g in (TT, TA, QS) and p[g] >= phe_minh:
            B.append(('2. Xét %s' % G[g], '%s đã mạnh (%.0f%% ≥ phe mình %.0f%%) – đã đủ chế/tiết thân → xét ứng viên sau.' % (G[g], 100 * p[g], 100 * phe_minh),
                      '"Có quan chế áp hỏa nên tương đối cân bằng, vì vậy lấy kim" (VD36 P2 tr.427); "thực thương quá mạnh cũng không thể dùng" (VD10 tr.405)'))
            continue
        hop.append(g)
    dh = 1 if thang in DONG else (4 if thang in HA else None); gdh = (dh - nh) % 5 if dh is not None else None
    if not hop:
        hop = [g for g in L if not (not nhuoc and g in (TT, TA, QS) and p[g] >= phe_minh)] or L[:]
        B.append(('3. Không ứng viên nào đủ lực', 'Lấy ứng viên khuyết theo thứ tự: %s (chờ tuế vận bổ cứu).' % G[hop[0]],
                  '"Nguyên cục không có dụng thần thì chỉ còn hy vọng hành đến vận gặp được dụng thần" (P1 tr.124)'))
    d = hop[0]
    if dieu_hau and gdh is not None and gdh in hop and gdh != d and p[gdh] < 0.2:
        B.append(('4. Điều hầu (P1 tr.76–79, 100)', 'Sinh mùa %s, %s (%s) còn ít và cũng là ứng viên hợp lệ → ưu tiên điều hầu.' % ('đông' if dh == 1 else 'hạ', HT[dh], G[gdh]),
                  '"Sinh tháng hạ dùng thủy, tháng đông dùng hỏa" (P1 tr.100)'))
        d = gdh
    else:
        B.append(('4. Điều hầu (P1 tr.76–79, 100)', 'Không cần đổi.' if gdh is None else 'Sinh mùa %s; điều hầu %s không đổi lựa chọn.' % ('đông' if dh == 1 else 'hạ', HT[dh]), 'P1 tr.100–101'))
    con = [g for g in hop if g != d] + [g for g in L if g not in hop and g != d]
    h = con[0] if con else (d + 4) % 5
    B.append(('5. Kết luận', 'Dụng thần %s (%s); hỷ thần %s (%s).' % (G[d], HT[(nh + d) % 5], G[h], HT[(nh + h) % 5]), ''))
    return d, h, B
