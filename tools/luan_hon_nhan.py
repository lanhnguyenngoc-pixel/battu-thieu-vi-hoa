"""Nguyên mẫu: luận HÔN NHÂN theo quy trình của tác giả, có lý giải từng bước (chưa vào workbook)."""
import sys, collections
sys.path.insert(0, '/home/claude/battu/tools')
from gd_b_mohinh import co_ban
from tt_co_ban import *
MUC = ['Kém (hôn nhân không thuận / dễ kết hôn hai lần)', 'Vừa (có trắc trở)', 'Tốt (hòa thuận)']
XUNG = {frozenset(p) for p in [('Tí','Ngọ'),('Sửu','Mùi'),('Dần','Thân'),('Mão','Dậu'),('Thìn','Tuất'),('Tị','Hợi')]}
HINH = {frozenset(p) for p in [('Dần','Tị'),('Tị','Thân'),('Dần','Thân'),('Sửu','Tuất'),('Tuất','Mùi'),('Sửu','Mùi'),('Tí','Mão')]}
AD_LECH = {'Bính Tí','Đinh Sửu','Mậu Dần','Tân Mão','Nhâm Thìn','Quý Tị','Bính Ngọ','Đinh Mùi','Mậu Thân','Tân Dậu','Nhâm Tuất','Quý Hợi'}
KD = {'Giáp':'Mão','Bính':'Ngọ','Mậu':'Ngọ','Canh':'Dậu','Nhâm':'Tí'}
def luan(r):
    a = co_ban(r); d = a['dem']; lo, bk, tg, ro, tat, chi, can, nc = a['lo'], a['bk'], a['tang'], a['ro'], a['tat'], a['chi'], a['can'], a['nc']
    nu = r['gioi'] == 'Nữ'; sao = 'QS' if nu else 'TA'; ten_sao = 'quan sát (chồng)' if nu else 'tài (vợ)'
    B = []
    B.append(('1. Thân và dụng thần', 'Nhật chủ %s %s; dụng thần %s, kỵ %s. %s.' % (nc, 'NHƯỢC' if a['nhuoc'] else 'VƯỢNG', r['role']['dung'], r['role']['ky'], 'Nữ: sao chồng là quan sát' if nu else 'Nam: sao vợ là tài'),
              '"Nam lấy tài làm vợ, nữ lấy quan làm chồng" (P1 tr.153); cung phối ngẫu là chi ngày.'))
    n_sao = d(sao, tat); n_ro = d(sao, ro)
    t2 = 'Sao %s: %s.' % (ten_sao, 'không có' if n_sao == 0 else '%d chữ rõ, %d chữ ẩn' % (n_ro, n_sao - n_ro))
    vt = a['vai'][sao]; t2 += ' Là %s.' % {'dung':'dụng thần','hy':'hỷ thần','ky':'kỵ thần','nhan':'nhàn thần'}[vt]
    B.append(('2. Sao phối ngẫu', t2, '"Có tài thì có vợ, ngược lại thì khó khăn" (P2 tr.520); nữ không có quan tinh "không là vợ lẽ thì là vợ kế" (Ch.15 VD27, tr.483).'))
    # 3. cung phối ngẫu
    cung = []
    if HANH_CAN[CAN.index(nc)] == HANH_CAN[CAN.index(TANG[chi[2]][0])]: cung.append('can chi ngày cùng hành (ngày tự ngang)')
    if can[2] + ' ' + chi[2] in AD_LECH: cung.append('ngày âm dương sai lệch')
    if nc == 'Giáp' and chi[2] in ('Thìn', 'Tuất'): cung.append('ngày Giáp Thìn/Giáp Tuất')
    if any(frozenset((chi[2], chi[i])) in XUNG for i in (1, 3)): cung.append('chi ngày bị tháng/giờ xung')
    if any(frozenset((chi[2], chi[i])) in HINH for i in (1, 3)) or chi[2] in (chi[1], chi[3]) and chi[2] in ('Thìn','Ngọ','Dậu','Hợi'): cung.append('chi ngày bị hình')
    B.append(('3. Cung phối ngẫu (chi ngày)', ('; '.join(cung) if cung else 'không bị xung hình, không sai lệch') + '.',
              '"Trụ ngày âm dương sai lệch… hôn nhân không thuận" (P2 tr.462; C23-605); "can chi trụ ngày giống nhau là khắc vợ" (Ch.15 VD8, tr.472); "ngày giáp thìn là tiêu chí khắc vợ" (tr.472).'))
    # 4. dấu hiệu khắc rõ
    kh = []
    if lo.count('Thương') >= 1: kh.append(('thương quan lộ', '"Nam gặp thương quan là có hai lần kết hôn, nữ gặp thương quan là…" (P2 tr.470)'))
    if sum(x in ('Tỉ','Kiếp') for x in ro) >= 3: kh.append(('tỉ kiếp trùng trùng (%d chữ rõ)' % sum(x in ('Tỉ','Kiếp') for x in ro), '"Tài yếu, thân vượng, tỉ kiếp nhiều là khắc vợ liên miên" (tr.469); "nam gặp tỉ kiếp tất tranh vợ, nữ tất tranh chồng" (tr.471)'))
    if nc in KD and KD[nc] in chi: kh.append(('có kình dương', '"Nam gặp kình dương là lấy vợ hai lần, nữ gặp kình dương tất sẽ tái giá" (tr.471)'))
    if nu:
        if 'Quan' in ro and 'Sát' in ro: kh.append(('quan sát hỗn tạp', '"quan sát hỗn tạp… " (tr.471); Ch.15 VD24, VD28'))
        if n_sao == 0: kh.append(('không có quan sát', 'Ch.15 VD27, VD33 (tr.483–485)'))
        if sum(x in ('Kiêu','Ấn') for x in ro) >= 3: kh.append(('kiêu ấn trùng lặp', '"Kiêu ấn trùng trùng là người phụ nữ tử biệt sinh ly" (Ch.15 VD22, VD30)'))
        if sum(c in ('Thìn','Tuất','Sửu','Mùi') for c in chi) >= 3: kh.append(('đủ thìn tuất sửu mùi', '"Nữ gặp thìn tuất sửu mùi là người vừa khắc chồng…" (Ch.16 VD11, tr.511)'))
        if all(c in ('Ất','Đinh','Kỷ','Tân','Quý') for c in can): kh.append(('tứ trụ thuần âm', '"thuần âm thì nữ mất chồng" (tr.471); Ch.15 VD30'))
    else:
        if lo.count('CTài') and lo.count('TTài') or ('CTài' in ro and 'TTài' in ro and d('TA', lo)): kh.append(('chính tài và thiên tài cùng thấu', '"chính tài, thiên tài thấu ra tất có hai vợ" (C23-614, C23-623)'))
        if n_sao == 0: kh.append(('không có tài', '"có tài thì có vợ, ngược lại thì khó khăn" (tr.520); Ch.15 VD15'))
        if all(c in ('Giáp','Bính','Mậu','Canh','Nhâm') for c in can): kh.append(('tứ trụ thuần dương', '"trong Tứ trụ thuần dương thì nam mất vợ" (tr.471)'))
        if sum(c in ('Thìn','Tuất','Sửu','Mùi') for c in chi) >= 3: kh.append(('gặp thìn tuất sửu mùi nhiều', '"Nếu gặp thìn, tuất, sửu, mùi nhiều là khắc vợ" (tr.469)'))
    so_cung = len(cung)
    B.append(('4. Dấu hiệu khắc phối ngẫu (ở mức rõ)', ('; '.join(k for k, _ in kh) if kh else 'không có') + '.',
              ' · '.join(c for _, c in kh) or 'Tác giả chỉ dùng dấu hiệu ở mức rõ/vượng: "thương quan hai rõ ba mờ", "tỉ kiếp trùng trùng" (Ch.15 VD22, VD25).'))
    # 5. cứu
    cu = []
    if nu and all(d(g, ro) >= 1 for g in ('TA', 'QS', 'AN')) and not ('Quan' in ro and 'Sát' in ro): cu.append(('tài – quan – ấn đủ, quan không hỗn tạp', '"Tài quan ấn thụ, tất vượng phu" (Ch.15 VD1 tr.467); "quan sát không hỗn tạp lại có ấn phù là chồng sẽ…" (tr.466)'))
    if not nu and NHOM[bk[2]] == 'TA' and not any('xung' in c or 'hình' in c for c in cung): cu.append(('tài đóng ở chi ngày, không bị xung hình', '"Nhật đóng ở tài tinh, giàu lên nhờ vợ" (Ch.15 VD4 tr.468)'))
    if vt in ('dung', 'hy') and n_ro: cu.append(('sao phối ngẫu là dụng/hỷ và hiện rõ', 'C23-588 (vợ khá giả); VD13 (nhờ vợ mà giàu)'))
    B.append(('5. Phẩm chất người phối ngẫu (điều tốt)', ('; '.join(c for c, _ in cu) if cu else 'không có dấu hiệu nổi bật') + '.', (' · '.join(c for _, c in cu) or '—') + ' — Lưu ý: trong ví dụ, các điều này nói về NGƯỜI phối ngẫu (giàu, đẹp, hiền), KHÔNG xóa được dấu hiệu khắc (C23-623: vợ giàu đẹp nhưng vẫn "tất có hai vợ").'))
    rat_ro = (lo.count('Thương') + bk.count('Thương') >= 2) or sum(x in ('Tỉ','Kiếp') for x in ro) >= 4 or (nu and sum(1 for x in lo if x in ('Quan','Sát')) >= 2 and 'Quan' in lo and 'Sát' in lo)
    tong = len(kh) + (1 if so_cung else 0)
    if tong >= 2 or (kh and rat_ro):
        m = 0; ly = 'Có %d dấu hiệu khắc rõ%s%s. Tác giả kết luận khắc/kết hôn hai lần khi các tiêu chí cùng có hoặc một tiêu chí rất nặng ("thương quan trùng lặp mà vượng", "tỉ kiếp trùng trùng" – Ch.15 VD14, VD16).' % (len(kh), (' (' + ', '.join(k for k, _ in kh) + ')') if kh else '', ' và cung phối ngẫu bất lợi' if so_cung else '')
    elif kh:
        m = 1; ly = 'Có 1 dấu hiệu khắc (%s), không có tiêu chí khác đi kèm → có trắc trở; trong các ví dụ, dấu hiệu đơn lẻ thành biến cố khi gặp vận/năm kích hoạt (xem bước 8).' % kh[0][0]
    elif so_cung: m = 1; ly = 'Không có dấu hiệu khắc rõ, chỉ có cung phối ngẫu bất lợi (%s) → có trắc trở, tình cảm không êm.' % ', '.join(cung)
    else: m = 2; ly = 'Không có dấu hiệu khắc, cung phối ngẫu yên → hôn nhân thuận.'
    B.append(('6. Gộp các điều kiện', ly, 'Cách tác giả liệt kê "tiêu chí khắc vợ" rồi kết luận (Ch.15 VD9 tr.472, VD12 tr.473, VD14 tr.474).'))
    B.append(('7. KẾT LUẬN MỨC GỐC', MUC[m], 'Năm dễ có biến cố: vận/năm thương quan, tỉ kiếp, kình dương hoặc xung cung phối ngẫu (bước 8).'))
    return m, B
quy = lambda m: 'K' if m == 0 else ('V' if m == 1 else 'T')
if __name__ == '__main__':
    from gd_b_dac_trung import chay
    X = chay(lv_list=('HN',))
    c = collections.Counter((r['sach'], quy(luan(r)[0])) for r in X); tong = collections.Counter(r['sach'] for r in X)
    print('KIỂM TRA (không chỉnh ngưỡng): khớp %d/%d' % (sum(c[(s, s)] for s in 'KVT'), len(X)), {s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'})
    print('Phần mềm hiện tại:', sum(r['cong_cu'] == r['sach'] for r in X))
    if len(sys.argv) > 1:
        for r in X:
            m, B = luan(r)
            if quy(m) != r['sach']: print(r['la'][:14], r['gioi'], r['sach'], '→', quy(m), '|', B[3][1][:90], '|', B[4][1][:50], '||', r['cau'][:70])
