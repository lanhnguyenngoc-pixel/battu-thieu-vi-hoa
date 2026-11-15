"""Mẫu thử DT-2: chọn dụng thần nhiều bước như cách tác giả vận dụng (chưa vào workbook).
Ba nguyên tắc P1 tr.99 (sinh phù–áp chế; thông quan; điều hầu) + bảng thế ngũ hành P1 tr.74–75
+ điều kiện dụng thần có lực, không bị thương (P1 tr.97). Mỗi nguyên tắc 'bỏ phiếu' cho ứng viên;
điểm bằng nhau giữa các nguyên tắc (không chỉnh trọng số theo ví dụ)."""
HT = ['Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy']
G = ['Tỉ kiếp', 'Thực thương', 'Tài', 'Quan sát', 'Ấn']
TK, TT, TA, QS, AN = range(5)
DONG = {'Hợi', 'Tí', 'Sửu'}; HA = {'Tị', 'Ngọ', 'Mùi'}


def phu_uc(nhuoc, a):
    if nhuoc:
        dom = max((TT, TA, QS), key=lambda i: a[i])
        return {QS: [AN, TK], TA: [TK, AN], TT: [AN, TK]}[dom], 'thân nhược, %s nhiều nhất' % G[dom]
    dom = AN if a[AN] > a[TK] else TK
    return ({AN: [TA, QS, TT], TK: [QS, TT, TA]}[dom]), 'thân vượng, %s nhiều' % G[dom]


def the_ngu_hanh(nh, a):
    D = max(range(5), key=lambda i: a[i])
    o = {AN: [TA, TT], TK: [QS, TA], QS: [TT, TK], TA: [TK, AN], TT: [AN, QS]}[D]
    if nh == 1 and D == AN: o = [QS, TA]
    if nh == 3 and D == TA: o = [AN, TK]
    return o, '%s quá mạnh' % G[D]


def chon(nh, nhuoc, a, thang, tong=None):
    """nh: hành nhật can (0..4); a: lượng 5 nhóm (PP2); thang: chi tháng. Trả (nhóm dụng, nhóm hỷ, bước lý giải)."""
    diem = [0.0] * 5; B = []
    L1, ly1 = phu_uc(nhuoc, a)
    for k, g in enumerate(L1[:2]): diem[g] += 2 - k
    B.append(('1. Sinh phù – áp chế (P1 tr.99)', '%s → ưu tiên %s.' % (ly1, ', '.join(G[g] for g in L1))))
    L2, ly2 = the_ngu_hanh(nh, a)
    for k, g in enumerate(L2): diem[g] += 2 - k
    B.append(('2. Thế ngũ hành của can ngày (P1 tr.74–75)', '%s → nên %s, thứ nữa %s.' % (ly2, G[L2[0]], G[L2[1]])))
    dh = 1 if thang in DONG else (4 if thang in HA else None)
    if dh is not None:
        g = (dh - nh) % 5; t = sum(a) or 1
        if a[g] / t < 0.2:
            diem[g] += 2
            B.append(('3. Điều hầu (P1 tr.76–79, 100)', 'Sinh mùa %s, %s (%s) còn ít → cần điều hầu.' % ('đông' if dh == 1 else 'hạ', HT[dh], G[g])))
        else:
            B.append(('3. Điều hầu (P1 tr.76–79, 100)', 'Sinh mùa %s nhưng %s đã đủ.' % ('đông' if dh == 1 else 'hạ', HT[dh])))
    else:
        B.append(('3. Điều hầu', 'Sinh mùa xuân/thu – không bắt buộc điều hầu (P1 tr.101).'))
    # 4. lực của ứng viên (P1 tr.97)
    ghi = []
    for g in range(5):
        if diem[g] <= 0: continue
        khac = (g + 3) % 5
        if a[khac] > a[g] and a[g] > 0 and g not in (TK, AN) or (g in (TK, AN) and nhuoc is False and False):
            pass
        if a[khac] > 2 * max(a[g], 1):
            diem[g] -= 1; ghi.append('%s bị %s chế mạnh (−1)' % (G[g], G[khac]))
        if g not in (TK, AN) and a[g] == max(a[x] for x in (TT, TA, QS)) and a[g] > a[TK] + a[AN]:
            diem[g] -= 1; ghi.append('%s đã quá nhiều (−1)' % G[g])
    B.append(('4. Dụng thần có lực, không bị thương (P1 tr.97)', '; '.join(ghi) or 'Các ứng viên không bị chế, không quá nhiều.'))
    thu_tu = sorted(range(5), key=lambda g: (-diem[g], L1.index(g) if g in L1 else 9))
    d, h = thu_tu[0], thu_tu[1]
    B.append(('5. Chọn', 'Điểm: ' + ', '.join('%s %g' % (G[g], diem[g]) for g in thu_tu if diem[g] != 0) +
              ' → dụng %s (%s), hỷ %s (%s).' % (G[d], HT[(nh + d) % 5], G[h], HT[(nh + h) % 5])))
    return d, h, B
