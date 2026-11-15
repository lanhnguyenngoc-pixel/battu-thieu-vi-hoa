"""Hàm cơ bản tự viết để nghiên cứu (KHÔNG thay engine): thập thần, tàng can, trường sinh."""
CAN = ['Giáp','Ất','Bính','Đinh','Mậu','Kỷ','Canh','Tân','Nhâm','Quý']
CHI = ['Tí','Sửu','Dần','Mão','Thìn','Tị','Ngọ','Mùi','Thân','Dậu','Tuất','Hợi']
HANH_CAN = [0,0,1,1,2,2,3,3,4,4]  # Mộc Hỏa Thổ Kim Thủy
TANG = {'Tí':['Quý'],'Sửu':['Kỷ','Quý','Tân'],'Dần':['Giáp','Bính','Mậu'],'Mão':['Ất'],'Thìn':['Mậu','Ất','Quý'],'Tị':['Bính','Canh','Mậu'],
        'Ngọ':['Đinh','Kỷ'],'Mùi':['Kỷ','Đinh','Ất'],'Thân':['Canh','Nhâm','Mậu'],'Dậu':['Tân'],'Tuất':['Mậu','Tân','Đinh'],'Hợi':['Nhâm','Giáp']}
def thap_than(nc, c):
    a, b = CAN.index(nc), CAN.index(c); ha, hb = HANH_CAN[a], HANH_CAN[b]; cung = (a % 2) == (b % 2)
    d = (hb - ha) % 5
    return {0: ('Tỉ' if cung else 'Kiếp'), 1: ('Thực' if cung else 'Thương'), 2: ('TTài' if cung else 'CTài'),
            3: ('Sát' if cung else 'Quan'), 4: ('Kiêu' if cung else 'Ấn')}[d]
TS = ['Trường sinh','Mộc dục','Quan đới','Lâm quan','Đế vượng','Suy','Bệnh','Tử','Mộ','Tuyệt','Thai','Dưỡng']
TS_GOC = {'Giáp':11,'Bính':2,'Mậu':2,'Canh':5,'Nhâm':8,'Ất':6,'Đinh':9,'Kỷ':9,'Tân':0,'Quý':3}
def truong_sinh(c, chi):
    i = CAN.index(c); s = TS_GOC[c]; j = CHI.index(chi)
    return TS[(j - s) % 12] if i % 2 == 0 else TS[(s - j) % 12]
NHOM = {'Tỉ':'TK','Kiếp':'TK','Thực':'TT','Thương':'TT','TTài':'TA','CTài':'TA','Sát':'QS','Quan':'QS','Kiêu':'AN','Ấn':'AN'}
def phan_tich(tt):
    p = [x.strip().split() for x in tt.split(',')]; can = [a for a, b in p]; chi = [b for a, b in p]; nc = can[2]
    lo = [thap_than(nc, c) for i, c in enumerate(can) if i != 2]        # can lộ (năm, tháng, giờ)
    bk = [thap_than(nc, TANG[c][0]) for c in chi]                       # bản khí 4 chi
    tang = [thap_than(nc, x) for c in chi for x in TANG[c][1:]]         # trung, dư khí
    return dict(can=can, chi=chi, nc=nc, lo=lo, bk=bk, tang=tang, lenh=bk[1],
                ts_ngay=truong_sinh(nc, chi[2]), lo_pos={i: thap_than(nc, c) for i, c in enumerate(can) if i != 2})
