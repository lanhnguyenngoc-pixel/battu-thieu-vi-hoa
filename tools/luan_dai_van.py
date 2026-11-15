"""Nguyên mẫu bước 8: luận từng đại vận cho tài lộc, quan vận, hôn nhân – có lý giải và câu sách (chưa vào workbook)."""
import sys
sys.path.insert(0, '/home/claude/battu/tools'); sys.path.insert(0, '/home/claude/battu')
from tt_co_ban import *
from gd_b_mohinh import co_ban
KD = {'Giáp':'Mão','Bính':'Ngọ','Mậu':'Ngọ','Canh':'Dậu','Nhâm':'Tí'}
XUNG = {frozenset(p) for p in [('Tí','Ngọ'),('Sửu','Mùi'),('Dần','Thân'),('Mão','Dậu'),('Thìn','Tuất'),('Tị','Hợi')]}
def dai_van(E):
    out = []
    for c in 'CDEFGHIJ':
        cc = str(E.gia_tri('DAI_VAN', c + '35')).split()
        out.append(dict(can=cc[0], chi=cc[1], tuoi=int(E.gia_tri('DAI_VAN', c + '36') or 0), nam=int(E.gia_tri('DAI_VAN', c + '37') or 0),
                        xu_the=str(E.gia_tri('DAI_VAN', c + '54'))))
    return out
def nhom_van(nc, v):
    return NHOM[thap_than(nc, v['can'])], NHOM[thap_than(nc, TANG[v['chi']][0])], thap_than(nc, v['can']), thap_than(nc, TANG[v['chi']][0])
TEN = {'TK':'tỉ kiếp','TT':'thực thương','TA':'tài','QS':'quan sát','AN':'ấn'}
def luan_van(r, dv):
    a = co_ban(r); d = a['dem']; nc, ro, tat, lo, chi = a['nc'], a['ro'], a['tat'], a['lo'], a['chi']; nu = r['gioi'] == 'Nữ'
    nhuoc = a['nhuoc']; co_tai = d('TA', tat) > 0; tai_nhieu = d('TA', ro) >= 2; co_quan = d('QS', tat) > 0
    co_chinh_quan = 'Quan' in tat; co_thuc = d('TT', ro) > 0; co_an = d('AN', ro) > 0; tk_nhieu = d('TK', ro) >= 3
    kq = []
    for v in dv:
      for nua in ('can', 'chi'):
        G1, G2, T1, T2 = nhom_van(nc, v)
        if nua == 'can': gs = {G1}; t1 = t2 = T1; tu = v['tuoi']; ten = '%s %s – 5 năm đầu (can %s: %s), %d–%d tuổi' % (v['can'], v['chi'], v['can'], T1, tu, tu + 4)
        else: gs = {G2}; t1 = t2 = T2; tu = v['tuoi'] + 5; ten = '%s %s – 5 năm sau (chi %s: %s), %d–%d tuổi' % (v['can'], v['chi'], v['chi'], T2, tu, tu + 4)
        TL, SN, HN = [], [], []
        # --- TÀI LỘC ---
        if 'TA' in gs:
            if not nhuoc: TL.append(('+', 'vận tài, thân gánh được → phát tài', '"Trong trụ có tài, hành đến đất tài tất đại phát"; "tài đóng ở nhật nguyên, hành đến vận tài thì sẽ phát" (P2 tr.522, 526)'))
            else: TL.append(('−', 'vận tài mà thân nhược → tiền đến nhưng khó giữ, dễ vì của mà sinh họa', '"Thân nhược tài nhiều thì gặp tài là mang họa"; "trụ ngày yếu không thắng nổi tài thì vì của cải mà bị tai hoạ" (P2 tr.526–527)'))
        if 'TK' in gs:
            if nhuoc and (co_tai): TL.append(('+', 'vận tỉ kiếp giúp thân gánh tài → phát tài', '"Thân nhược, tài nhiều khi hành đến vận tỉ kiếp thì mới phát tài" (P2 tr.521; ứng P1-VD12 "39 tuổi dụng thần lên ngôi trở thành người giàu")'))
            elif not nhuoc: TL.append(('−', 'thân vượng gặp vận tỉ kiếp → phá tài, tranh tài', '"Thân vượng hành đến vận tỷ kiếp thì phá tài bại nghiệp"; "tỉ kiếp nhiều lại gặp tỉ kiếp là phá tài, ra cửa quan" (P2 tr.527; ứng VD17 sau 40 tuổi, Ch.18 VD10–11)'))
        if 'TT' in gs and co_tai and not nhuoc: TL.append(('+', 'vận thực thương sinh tài → thu nhập tăng', '"Thương quan sinh tài thì tự nhiên giàu có" (P2 tr.521; ứng VD5 "sau 31 hành vận thực thương tài mới phát")'))
        if 'AN' in gs and co_thuc and not nhuoc and t1 == 'Kiêu': TL.append(('−', 'vận kiêu gặp thực thần → tài vật hao tán', '"Thực thần gặp kiêu là tài vật hao tán" (P2 tr.527; ứng Ch.18 VD12)'))
        if 'QS' in gs and nhuoc and co_quan: TL.append(('−', 'thân nhược quan vượng gặp vận quan → phá tài', '"Quan vượng, thân nhược, hành đến quan vận thì hại vô cùng, phá tài là chắc chắn" (P2 tr.527)'))
        if 'QS' in gs and tk_nhieu and co_tai: TL.append(('+', 'vận quan sát chế tỉ kiếp, bảo vệ tài', '"Tài tinh gặp tỉ kiếp… được cứu trợ" (P2 tr.470); quan chế tỉ kiếp'))
        # --- QUAN VẬN ---
        if 'QS' in gs:
            if 'Thương' in lo and co_chinh_quan and not co_an: SN.append(('−', 'trụ có thương quan, vận lại gặp quan → tai họa vì chức vụ', '"Thương quan khi thương hết thì kị nhất lại gặp quan tinh" (P2 tr.542)'))
            elif not nhuoc: SN.append(('+', 'vận quan sát, thân gánh được → thăng chức', '"Quan vận gặp thân vượng tất sẽ thăng chức" (P2 tr.534; ứng VD25 9–48 tuổi)'))
            else: SN.append(('−', 'thân nhược gặp vận quan sát → chịu áp lực, bị đè nén', '"Ngày nhược gặp quan tinh thì thành đạt không lâu" (P2 tr.542; ứng Ch.19 VD7)'))
        if 'TT' in gs and t1 == 'Thương' or t2 == 'Thương':
            if co_chinh_quan: SN.append(('−', 'vận thương quan gặp quan trong trụ → dễ mất chức hoặc ốm', '"Trụ có quan có chức nhưng hành vận thương quan thì bị bãi miễn"; "thương quan gặp quan, không mất chức cũng bị ốm" (P2 tr.541–542; ứng Ch.19 VD8–9, VD25 sau 49)'))
        if 'AN' in gs and co_quan:
            if nhuoc: SN.append(('+', 'vận ấn: quan ấn tương sinh, thân được giúp → lợi công danh', '"Quan, ấn tương sinh lại gặp tuế vận là quan trong triều" (P2 tr.537; ứng VD37 "sau 43 vận chính ấn cầu quan thuận lợi nhất")'))
            elif 'Sát' in tat: SN.append(('−', 'thân vượng có sát lại gặp vận ấn → dễ mất chức', '"Thân vượng có sát, lại hành vận ấn thụ là mất chức" (P2 tr.534)'))
        if 'TA' in gs and co_quan and not nhuoc: SN.append(('+', 'vận tài sinh quan → thăng tiến', '"Tài vượng sinh quan thì đạt được vinh hiển" (P2 tr.534; ứng Ch.19 VD2 "40 tuổi hành tài vận, 1989 thăng chức")'))
        if 'TA' in gs and a['vai']['AN'] == 'dung': SN.append(('−', 'vận tài phá ấn (ấn là dụng) → dễ mất chức', '"Ấn gặp tài là mất chức"; "ấn gặp tài địa, không mất chức cũng nên rút lui" (P2 tr.536, 542)'))
        # --- HÔN NHÂN ---
        if t1 == 'Thương' or t2 == 'Thương': HN.append(('−', 'vận thương quan → %s' % ('khắc chồng' if nu else 'khắc vợ'), '"Gặp phải vận thương quan là nam khắc vợ, nữ khắc chồng" (P2 tr.471; ứng Ch.15 VD27)'))
        if not nu and 'TK' in gs: HN.append(('−', 'vận tỉ kiếp tranh tài (vợ)', '"Tài tinh đã bị tỉ kiếp, tuế vận lại còn gặp tỉ kiếp thì năm đó… " (P2 tr.469; ứng Ch.15 VD9, VD18)'))
        if nua == 'chi' and nc in KD and v['chi'] == KD[nc]: HN.append(('−', 'vận kình dương → %s' % ('khắc chồng, tái giá' if nu else 'khắc vợ, hao tài'), '"Người trong Tứ trụ không có kình dương, hành vận đến lúc gặp kình dương…" (P2 tr.468; ứng Ch.15 VD17)'))
        if nua == 'chi' and frozenset((v['chi'], chi[2])) in XUNG: HN.append(('−', 'chi vận xung cung phối ngẫu (chi ngày)', 'Cung phối ngẫu bị xung (P2 tr.462; ứng P1-VD1 "cung chồng bị xung thì sắp ly hôn")'))
        if (('QS' in gs) if nu else ('TA' in gs)): HN.append(('+', 'vận gặp sao phối ngẫu → dễ có duyên, kết hôn', 'Sao phối ngẫu đến vận (P1 tr.153–157)'))
        kq.append(dict(van=ten, xu_the_cong_cu=v['xu_the'], TL=TL, SN=SN, HN=HN))
    kq.insert(0, dict(ghi_chu='Can đại vận quản 5 năm đầu, chi quản 5 năm sau (P1 tr.120–124).'))
    return kq
def tom(ds):
    if not ds: return 'bình thường'
    s = sum(1 if x[0] == '+' else -1 for x in ds)
    return ('THUẬN' if s > 0 else 'NGHỊCH' if s < 0 else 'LẪN LỘN') + ': ' + '; '.join(x[1] for x in ds)
if __name__ == '__main__':
    from do_luan_giai import engine_tu_tru
    import json
    G = json.load(open('/home/claude/battu/tai_lieu/giai_doan_A/dap_an_tong.json'))
    for tt, gioi, ten in [('Canh Tuất,Tân Tị,Bính Ngọ,Nhâm Thìn', 'Nam', 'P2-410 VD17'), ('Canh Thìn,Mậu Tí,Canh Tuất,Kỷ Mão', 'Nữ', 'P2-417 VD25'),
                          ('Canh Thân,Quý Mùi,Mậu Thân,Quý Sửu', 'Nữ', 'P1-VD12 (nữ)')]:
        E = engine_tu_tru(tt, gioi)
        from gd_b_dac_trung import vai_tro
        r = dict(tu_tru=tt, gioi=gioi, than=str(E.gia_tri('DOI_CHIEU', 'C34')), role=vai_tro(E))
        print('\n=====', ten, tt, '| sách DV:', G.get(ten, {}).get('DV', ('',))[0])
        for k in luan_van(r, dai_van(E))[1:]:
            print(' ', k['van'], '| công cụ hiện tại:', k['xu_the_cong_cu'])
            print('     Tài lộc:', tom(k['TL'])); print('     Quan vận:', tom(k['SN'])); print('     Hôn nhân:', tom(k['HN']))
