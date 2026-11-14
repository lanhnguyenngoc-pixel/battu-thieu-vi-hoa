# -*- coding: utf-8 -*-
"""Âm lịch Việt Nam (UTC+7) ↔ dương lịch.

Workbook 6.3 chỉ nhận ngày giờ DƯƠNG LỊCH hoặc can chi (YEU_CAU 5.1). Khi người dùng nhập âm lịch,
phần mềm đổi sang dương lịch rồi đưa vào NHAP_LIEU C26:G26 như nhập dương lịch.

Phần dương → âm (jd, newmoon, NM, sunlong, SL, m11, leapoff, s2l) CHÉP NGUYÊN VĂN từ
09_TAI_LIEU_BO_SUNG/ngu_hanh_thieu/lunar.py (đáp án đổi lịch của hồ sơ) – không sửa.
Phần âm → dương (l2s) dựng từ chính các hàm đó; kiểm khứ hồi với s2l trên mọi ngày 1899–2101
(tests/t_am_lich.py)."""
import math, datetime


# ===================== CHÉP NGUYÊN VĂN lunar.py =====================
def INT(x): return math.floor(x)
def jd(dd,mm,yy):
    a=INT((14-mm)/12); y=yy+4800-a; m=mm+12*a-3
    j=dd+INT((153*m+2)/5)+365*y+INT(y/4)-INT(y/100)+INT(y/400)-32045
    if j<2299161: j=dd+INT((153*m+2)/5)+365*y+INT(y/4)-32083
    return j
def newmoon(k):
    T=k/1236.85; T2=T*T; T3=T2*T; dr=math.pi/180
    Jd1=2415020.75933+29.53058868*k+0.0001178*T2-0.000000155*T3+0.00033*math.sin((166.56+132.87*T-0.009173*T2)*dr)
    M=359.2242+29.10535608*k-0.0000333*T2-0.00000347*T3
    Mpr=306.0253+385.81691806*k+0.0107306*T2+0.00001236*T3
    F=21.2964+390.67050646*k-0.0016528*T2-0.00000239*T3
    C1=(0.1734-0.000393*T)*math.sin(M*dr)+0.0021*math.sin(2*dr*M)-0.4068*math.sin(Mpr*dr)+0.0161*math.sin(dr*2*Mpr)-0.0004*math.sin(dr*3*Mpr)
    C1+=0.0104*math.sin(dr*2*F)-0.0051*math.sin(dr*(M+Mpr))-0.0074*math.sin(dr*(M-Mpr))+0.0004*math.sin(dr*(2*F+M))-0.0004*math.sin(dr*(2*F-M))-0.0006*math.sin(dr*(2*F+Mpr))+0.0010*math.sin(dr*(2*F-Mpr))+0.0005*math.sin(dr*(2*Mpr+M))
    dt=0.0465 if T<-11 else (-0.000278+0.000265*T+0.000262*T2 if T>=-11 else 0)
    if T<-11: dt=0.001+0.000839*T+0.0002261*T2-0.00000845*T3-0.000000081*T*T3
    return Jd1+C1-dt
def NM(k,tz): return INT(newmoon(k)+0.5+tz/24)
def sunlong(jdn):
    T=(jdn-2451545.0)/36525; T2=T*T; dr=math.pi/180
    M=357.52910+35999.05030*T-0.0001559*T2-0.00000048*T*T2
    L0=280.46645+36000.76983*T+0.0003032*T2
    DL=(1.914600-0.004817*T-0.000014*T2)*math.sin(dr*M)+(0.019993-0.000101*T)*math.sin(dr*2*M)+0.000290*math.sin(dr*3*M)
    L=(L0+DL)*dr; L=L-math.pi*2*INT(L/(math.pi*2)); return L
def SL(d,tz): return INT(sunlong(d-0.5-tz/24)/math.pi*6)
def m11(yy,tz):
    off=jd(31,12,yy)-2415021; k=INT(off/29.530588853); n=NM(k,tz)
    return NM(k-1,tz) if SL(n,tz)>=9 else n
def leapoff(a11,tz):
    k=INT((a11-2415021.076998695)/29.530588853+0.5); last=0; i=1; arc=SL(NM(k+i,tz),tz)
    while True:
        last=arc; i+=1; arc=SL(NM(k+i,tz),tz)
        if not(arc!=last and i<14): break
    return i-1
def s2l(dd,mm,yy,tz=7):
    d=jd(dd,mm,yy); k=INT((d-2415021.076998695)/29.530588853)
    ms=NM(k+1,tz)
    if ms>d: ms=NM(k,tz)
    a11=m11(yy,tz); b11=a11
    if a11>=ms: ly=yy; a11=m11(yy-1,tz)
    else: ly=yy+1; b11=m11(yy+1,tz)
    ld=d-ms+1; diff=INT((ms-a11)/29); leap=0; lm=diff+11
    if b11-a11>365:
        lo=leapoff(a11,tz)
        if diff>=lo:
            lm=diff+10
            if diff==lo: leap=1
    if lm>12: lm-=12
    if lm>=11 and diff<4: ly-=1
    return ld,lm,ly,leap
# ===================== HẾT PHẦN CHÉP NGUYÊN VĂN =====================


def jd_to_date(j):
    """Số ngày Julius → (ngày, tháng, năm) dương lịch (lịch Gregory)."""
    if j > 2299160:
        a = j + 32044; b = INT((4 * a + 3) / 146097); c = a - INT((b * 146097) / 4)
    else:
        b = 0; c = j + 32082
    d = INT((4 * c + 3) / 1461); e = c - INT((1461 * d) / 4); m = INT((5 * e + 2) / 153)
    day = e - INT((153 * m + 2) / 5) + 1; month = m + 3 - 12 * INT(m / 10); year = b * 100 + d - 4800 + INT(m / 10)
    return day, month, year


def thang_nhuan(ly, tz=7):
    """Tháng nhuận của năm âm ly (1–12), 0 nếu năm đó không nhuận."""
    for m in range(1, 13):
        if l2s(1, m, ly, 1, tz) is not None: return m
    return 0


def l2s(ld, lm, ly, leap=0, tz=7):
    """Âm lịch (ngày, tháng, năm, nhuận 0/1) → (ngày, tháng, năm) dương lịch. Trả None nếu ngày âm không tồn tại."""
    if lm < 11:
        a11 = m11(ly - 1, tz); b11 = m11(ly, tz)
    else:
        a11 = m11(ly, tz); b11 = m11(ly + 1, tz)
    k = INT(0.5 + (a11 - 2415021.076998695) / 29.530588853)
    off = lm - 11
    if off < 0: off += 12
    if b11 - a11 > 365:
        lo = leapoff(a11, tz); lpm = lo - 2
        if lpm < 0: lpm += 12
        if leap and lm != lpm: return None
        if leap or off >= lo: off += 1
    elif leap:
        return None
    ms = NM(k + off, tz)
    kq = jd_to_date(ms + ld - 1)
    # xác nhận khứ hồi bằng s2l (đúng quy ước lunar.py); sai thì coi là ngày không tồn tại
    if s2l(*kq, tz) != (ld, lm, ly, 1 if leap else 0): return None
    return kq


def so_ngay_thang_am(lm, ly, leap=0, tz=7):
    """Số ngày của tháng âm (29 hoặc 30)."""
    return 30 if l2s(30, lm, ly, leap, tz) is not None else 29
