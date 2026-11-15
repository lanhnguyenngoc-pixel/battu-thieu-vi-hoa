# -*- coding: utf-8 -*-
"""Module sheet DOI_CHIEU – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.6.xlsx (SHA-256 c1e0254833a8f522615199e1e0140460e912227d0b9ffc88b9164ebc26d3e0b3).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'DOI_CHIEU'
SI = 7

def f0(E, r, c):
    # Ô C6
    # = PP1_DINH_TINH!$C$35
    return E.v(9,35,3)

def f1(E, r, c):
    # Ô D6
    # = IF(PP2_DINH_LUONG!$C$43=0,"",IF(PP2_DINH_LUONG!$C$44=1,"tranh hợp, không hóa",IF(PP2_DINH_LUONG!$C$45=1,"cả hai can hư phù → không hợp",IF(PP2_DINH_LU
    #   ONG!$C$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$28+1)&" ("&CHOOSE(PP2_DINH_LUONG!$C$27,"chính hóa","phu tòng thê hóa","thê tòng phu hóa")&")","hợp
    #    không hóa"))))
    return ('' if X.truthy(X.eq(E.v(10,43,3),0.0)) else ('tranh hợp, không hóa' if X.truthy(X.eq(E.v(10,44,3),1.0)) else ('cả hai can hư phù → không hợp' if X.truthy(X.eq(E.v(10,45,3),1.0)) else (X.cat(X.cat(X.cat(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,28,3),1.0))),' ('),X.CHOOSE(E.v(10,27,3), lambda: 'chính hóa', lambda: 'phu tòng thê hóa', lambda: 'thê tòng phu hóa')),')') if X.truthy(X.eq(E.v(10,47,3),1.0)) else 'hợp không hóa'))))

def f2(E, r, c):
    # Ô E6
    # = IF(PP1_DINH_TINH!$C$33=0,"",IF(PP1_DINH_TINH!$C$34=PP2_DINH_LUONG!$C$47,"Giống","KHÁC"))
    return ('' if X.truthy(X.eq(E.v(9,33,3),0.0)) else ('Giống' if X.truthy(X.eq(E.v(9,34,3),E.v(10,47,3))) else 'KHÁC'))

def f3(E, r, c):
    # Ô F6, F7, F8
    # = IF(E6="KHÁC","PP1 (P1 tr.62): can ngày hợp can kề hóa khi chi tháng cùng hành hóa thần hoặc 3 chi còn lại thành cục hóa thần; can năm–tháng cần chi nă
    #   m có gốc hóa thần. PP2 (P2 tr.336-339): hóa thần phải cùng loại lệnh/nhân nguyên chi tháng, chi dưới cặp"&" can cùng loại/sinh hóa thần; cả hai can hư
    #    phù thì không hợp.","")
    return (X.cat('PP1 (P1 tr.62): can ngày hợp can kề hóa khi chi tháng cùng hành hóa thần hoặc 3 chi còn lại thành cục hóa thần; can năm–tháng cần chi năm có gốc hóa thần. PP2 (P2 tr.336-339): hóa thần phải cùng loại lệnh/nhân nguyên chi tháng, chi dưới cặp',' can cùng loại/sinh hóa thần; cả hai can hư phù thì không hợp.') if X.truthy(X.eq(E.v(7,r,c-1),'KHÁC')) else '')

def f4(E, r, c):
    # Ô C7
    # = PP1_DINH_TINH!$D$35
    return E.v(9,35,4)

def f5(E, r, c):
    # Ô D7
    # = IF(PP2_DINH_LUONG!$D$43=0,"",IF(PP2_DINH_LUONG!$D$44=1,"tranh hợp, không hóa",IF(PP2_DINH_LUONG!$D$45=1,"cả hai can hư phù → không hợp",IF(PP2_DINH_LU
    #   ONG!$D$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$28+1)&" ("&CHOOSE(PP2_DINH_LUONG!$D$27,"chính hóa","phu tòng thê hóa","thê tòng phu hóa")&")","hợp
    #    không hóa"))))
    return ('' if X.truthy(X.eq(E.v(10,43,4),0.0)) else ('tranh hợp, không hóa' if X.truthy(X.eq(E.v(10,44,4),1.0)) else ('cả hai can hư phù → không hợp' if X.truthy(X.eq(E.v(10,45,4),1.0)) else (X.cat(X.cat(X.cat(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,28,4),1.0))),' ('),X.CHOOSE(E.v(10,27,4), lambda: 'chính hóa', lambda: 'phu tòng thê hóa', lambda: 'thê tòng phu hóa')),')') if X.truthy(X.eq(E.v(10,47,4),1.0)) else 'hợp không hóa'))))

def f6(E, r, c):
    # Ô E7
    # = IF(PP1_DINH_TINH!$D$33=0,"",IF(PP1_DINH_TINH!$D$34=PP2_DINH_LUONG!$D$47,"Giống","KHÁC"))
    return ('' if X.truthy(X.eq(E.v(9,33,4),0.0)) else ('Giống' if X.truthy(X.eq(E.v(9,34,4),E.v(10,47,4))) else 'KHÁC'))

def f7(E, r, c):
    # Ô C8
    # = PP1_DINH_TINH!$E$35
    return E.v(9,35,5)

def f8(E, r, c):
    # Ô D8
    # = IF(PP2_DINH_LUONG!$E$43=0,"",IF(PP2_DINH_LUONG!$E$44=1,"tranh hợp, không hóa",IF(PP2_DINH_LUONG!$E$45=1,"cả hai can hư phù → không hợp",IF(PP2_DINH_LU
    #   ONG!$E$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$28+1)&" ("&CHOOSE(PP2_DINH_LUONG!$E$27,"chính hóa","phu tòng thê hóa","thê tòng phu hóa")&")","hợp
    #    không hóa"))))
    return ('' if X.truthy(X.eq(E.v(10,43,5),0.0)) else ('tranh hợp, không hóa' if X.truthy(X.eq(E.v(10,44,5),1.0)) else ('cả hai can hư phù → không hợp' if X.truthy(X.eq(E.v(10,45,5),1.0)) else (X.cat(X.cat(X.cat(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,28,5),1.0))),' ('),X.CHOOSE(E.v(10,27,5), lambda: 'chính hóa', lambda: 'phu tòng thê hóa', lambda: 'thê tòng phu hóa')),')') if X.truthy(X.eq(E.v(10,47,5),1.0)) else 'hợp không hóa'))))

def f9(E, r, c):
    # Ô E8
    # = IF(PP1_DINH_TINH!$E$33=0,"",IF(PP1_DINH_TINH!$E$34=PP2_DINH_LUONG!$E$47,"Giống","KHÁC"))
    return ('' if X.truthy(X.eq(E.v(9,33,5),0.0)) else ('Giống' if X.truthy(X.eq(E.v(9,34,5),E.v(10,47,5))) else 'KHÁC'))

def f10(E, r, c):
    # Ô C9
    # = IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),IF(PP1_DINH_TINH!$C$28=1,"hợp không hóa (đứng độc lập)",IF(PP1_DINH_TINH!$C$29=
    #   1,"bị xung","không hợp hóa")))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else ('hợp không hóa (đứng độc lập)' if X.truthy(X.eq(E.v(9,28,3),1.0)) else ('bị xung' if X.truthy(X.eq(E.v(9,29,3),1.0)) else 'không hợp hóa')))

def f11(E, r, c):
    # Ô D9
    # = IF(PP2_DINH_LUONG!$C$24>=0,"hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$24+1),IF(PP2_DINH_LUONG!$C$33=1,"khử bì (hợp không hóa / xung)","đủ tàng can"))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,3),1.0))) if X.truthy(X.ge(E.v(10,24,3),0.0)) else ('khử bì (hợp không hóa / xung)' if X.truthy(X.eq(E.v(10,33,3),1.0)) else 'đủ tàng can'))

def f12(E, r, c):
    # Ô E9
    # = IF(PP1_DINH_TINH!$C$27=PP2_DINH_LUONG!$C$24,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(9,27,3),E.v(10,24,3))) else 'KHÁC')

def f13(E, r, c):
    # Ô F9, F10, F11, F12
    # = IF(E9="KHÁC","PP1 (P1 tr.62, 67-68): cục/lục hợp hóa khi có can lộ hóa thần; tam hợp có chi vượng bị chi kề xung thì phá. PP2 (P2 tr.343-344): lục hợp
    #    còn cần bản khí chi tháng cùng loại/sinh hóa thần; can đã hóa cũng dẫn hóa chi; không xét phá cục bằng xung.","")
    return ('PP1 (P1 tr.62, 67-68): cục/lục hợp hóa khi có can lộ hóa thần; tam hợp có chi vượng bị chi kề xung thì phá. PP2 (P2 tr.343-344): lục hợp còn cần bản khí chi tháng cùng loại/sinh hóa thần; can đã hóa cũng dẫn hóa chi; không xét phá cục bằng xung.' if X.truthy(X.eq(E.v(7,r,c-1),'KHÁC')) else '')

def f14(E, r, c):
    # Ô C10
    # = IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),IF(PP1_DINH_TINH!$D$28=1,"hợp không hóa (đứng độc lập)",IF(PP1_DINH_TINH!$D$29=
    #   1,"bị xung","không hợp hóa")))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else ('hợp không hóa (đứng độc lập)' if X.truthy(X.eq(E.v(9,28,4),1.0)) else ('bị xung' if X.truthy(X.eq(E.v(9,29,4),1.0)) else 'không hợp hóa')))

def f15(E, r, c):
    # Ô D10
    # = IF(PP2_DINH_LUONG!$D$24>=0,"hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$24+1),IF(PP2_DINH_LUONG!$D$33=1,"khử bì (hợp không hóa / xung)","đủ tàng can"))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,4),1.0))) if X.truthy(X.ge(E.v(10,24,4),0.0)) else ('khử bì (hợp không hóa / xung)' if X.truthy(X.eq(E.v(10,33,4),1.0)) else 'đủ tàng can'))

def f16(E, r, c):
    # Ô E10
    # = IF(PP1_DINH_TINH!$D$27=PP2_DINH_LUONG!$D$24,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(9,27,4),E.v(10,24,4))) else 'KHÁC')

def f17(E, r, c):
    # Ô C11
    # = IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),IF(PP1_DINH_TINH!$E$28=1,"hợp không hóa (đứng độc lập)",IF(PP1_DINH_TINH!$E$29=
    #   1,"bị xung","không hợp hóa")))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else ('hợp không hóa (đứng độc lập)' if X.truthy(X.eq(E.v(9,28,5),1.0)) else ('bị xung' if X.truthy(X.eq(E.v(9,29,5),1.0)) else 'không hợp hóa')))

def f18(E, r, c):
    # Ô D11
    # = IF(PP2_DINH_LUONG!$E$24>=0,"hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$24+1),IF(PP2_DINH_LUONG!$E$33=1,"khử bì (hợp không hóa / xung)","đủ tàng can"))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,5),1.0))) if X.truthy(X.ge(E.v(10,24,5),0.0)) else ('khử bì (hợp không hóa / xung)' if X.truthy(X.eq(E.v(10,33,5),1.0)) else 'đủ tàng can'))

def f19(E, r, c):
    # Ô E11
    # = IF(PP1_DINH_TINH!$E$27=PP2_DINH_LUONG!$E$24,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(9,27,5),E.v(10,24,5))) else 'KHÁC')

def f20(E, r, c):
    # Ô C12
    # = IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),IF(PP1_DINH_TINH!$F$28=1,"hợp không hóa (đứng độc lập)",IF(PP1_DINH_TINH!$F$29=
    #   1,"bị xung","không hợp hóa")))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else ('hợp không hóa (đứng độc lập)' if X.truthy(X.eq(E.v(9,28,6),1.0)) else ('bị xung' if X.truthy(X.eq(E.v(9,29,6),1.0)) else 'không hợp hóa')))

def f21(E, r, c):
    # Ô D12
    # = IF(PP2_DINH_LUONG!$F$24>=0,"hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$F$24+1),IF(PP2_DINH_LUONG!$F$33=1,"khử bì (hợp không hóa / xung)","đủ tàng can"))
    return (X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,6),1.0))) if X.truthy(X.ge(E.v(10,24,6),0.0)) else ('khử bì (hợp không hóa / xung)' if X.truthy(X.eq(E.v(10,33,6),1.0)) else 'đủ tàng can'))

def f22(E, r, c):
    # Ô E12
    # = IF(PP1_DINH_TINH!$F$27=PP2_DINH_LUONG!$F$24,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(9,27,6),E.v(10,24,6))) else 'KHÁC')

def f23(E, r, c):
    # Ô C13
    # = IF(PP1_DINH_TINH!$C$43=1,"Nhật can được lệnh","Nhật can không được lệnh")&" ("&TRUNG_GIAN!$D$29&")"
    return X.cat(X.cat(X.cat(('Nhật can được lệnh' if X.truthy(X.eq(E.v(9,43,3),1.0)) else 'Nhật can không được lệnh'),' ('),E.v(23,29,4)),')')

def f24(E, r, c):
    # Ô D13
    # = "Hành nắm lệnh: "&PP2_DINH_LUONG!$C$97&"; hệ số nhật chủ ×"&TEXT(INDEX(PP2_DINH_LUONG!$E$91:$E$95,PP2_DINH_LUONG!$C$98+1),"0.0")
    return X.cat(X.cat(X.cat('Hành nắm lệnh: ',E.v(10,97,3)),'; hệ số nhật chủ ×'),X.TEXT(X.INDEX(E.rg(10,91,5,95,5), X.add(E.v(10,98,3),1.0)), '0.0'))

def f25(E, r, c):
    # Ô F13
    # = "PP1 (P1 tr.70): được lệnh là tiêu chí quyết định hàng đầu. PP2 (P2 tr.345-346): lệnh chỉ nhân hệ số ±1/5 vào tổng độ."
    return 'PP1 (P1 tr.70): được lệnh là tiêu chí quyết định hàng đầu. PP2 (P2 tr.345-346): lệnh chỉ nhân hệ số ±1/5 vào tổng độ.'

def f26(E, r, c):
    # Ô C14
    # = "Đắc địa "&PP1_DINH_TINH!$C$63&", được sinh "&PP1_DINH_TINH!$C$64&", được trợ "&PP1_DINH_TINH!$C$65&" (mạnh: "&PP1_DINH_TINH!$D$66&"/3)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đắc địa ',E.v(9,63,3)),', được sinh '),E.v(9,64,3)),', được trợ '),E.v(9,65,3)),' (mạnh: '),E.v(9,66,4)),'/3)')

def f27(E, r, c):
    # Ô D14
    # = TEXT(PP2_DINH_LUONG!$C$99,"0.0")&" / "&TEXT(PP2_DINH_LUONG!$C$100,"0.0")&" độ = "&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")
    return X.cat(X.cat(X.cat(X.cat(X.TEXT(E.v(10,99,3), '0.0'),' / '),X.TEXT(E.v(10,100,3), '0.0')),' độ = '),X.TEXT(E.v(10,101,3), '0.0%'))

def f28(E, r, c):
    # Ô F14
    # = "PP1 đếm điều kiện có/không và mạnh/yếu; PP2 cộng số độ của phe mình (hành nhật chủ + hành sinh nhật chủ)."
    return 'PP1 đếm điều kiện có/không và mạnh/yếu; PP2 cộng số độ của phe mình (hành nhật chủ + hành sinh nhật chủ).'

def f29(E, r, c):
    # Ô C15
    # = DUNG_HY_KY!$C$6
    return E.v(8,6,3)

def f30(E, r, c):
    # Ô D15
    # = DUNG_HY_KY!$D$6&" ("&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")&")"
    return X.cat(X.cat(X.cat(E.v(8,6,4),' ('),X.TEXT(E.v(10,101,3), '0.0%')),')')

def f31(E, r, c):
    # Ô E15
    # = IF(IF(OR(DUNG_HY_KY!$C$6="Quá vượng",DUNG_HY_KY!$C$6="Rất vượng",DUNG_HY_KY!$C$6="Thân vượng"),"vượng",IF(DUNG_HY_KY!$C$6="Trung hòa","trung hòa","như
    #   ợc"))=IF(DUNG_HY_KY!$D$6="Nhật chủ nhược","nhược","vượng"),"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(('vượng' if X.truthy(X.OR(X.eq(E.v(8,6,3),'Quá vượng'), X.eq(E.v(8,6,3),'Rất vượng'), X.eq(E.v(8,6,3),'Thân vượng'))) else ('trung hòa' if X.truthy(X.eq(E.v(8,6,3),'Trung hòa')) else 'nhược')),('nhược' if X.truthy(X.eq(E.v(8,6,4),'Nhật chủ nhược')) else 'vượng'))) else 'KHÁC')

def f32(E, r, c):
    # Ô F15
    # = IF(E15="Giống","Hai phương pháp thống nhất về vượng/nhược.",IF(COUNTIF(E6:E12,"KHÁC")>0,"(1) Hợp hóa khác nhau ở "&COUNTIF(E6:E12,"KHÁC")&" chỗ (xem c
    #   ác dòng trên) làm đổi lực lượng các hành. ","")&IF(ISNUMBER(SEARCH("a–d",PP1_DINH_TINH!$C$92)),"(2) PP1 rơi vào trường hợp nửa vượng nửa nhược, phải p
    #   hân xử a–d (P1 tr.71), tổng điểm "&PP1_DINH_TINH!$C$88&" – kết luận nhạy. ","")&IF(ABS(PP2_DINH_LUONG!$C$101-P_NGUONG)<0.05,"(3) PP2 sát ngưỡng 40% (t
    #   ỉ lệ "&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")&"). ","")&IF(AND(PP1_DINH_TINH!$C$43=1,IF(DUNG_HY_KY!$D$6="Nhật chủ nhược","nhược","vượng")="nhược"),"(4) PP
    #   1 coi được lệnh là quyết định (P1 tr.70), PP2 chỉ cộng 20% cho hành nắm lệnh nên phe khác nhiều độ vẫn thắng. ","")&IF(AND(PP1_DINH_TINH!$C$43=0,IF(DU
    #   NG_HY_KY!$D$6="Nhật chủ nhược","nhược","vượng")="vượng"),"(4) PP1 không được lệnh thì cần ≥2 điều mạnh (P1 tr.70-71); PP2 cộng dồn độ nên nhiều chữ cù
    #   ng phe vẫn đạt ≥40%. ","")&IF(DUNG_HY_KY!$C$6="Trung hòa","(5) PP1 cho trung hòa – PP2 không có khái niệm này, luôn chia vượng/nhược theo ngưỡng 40%. 
    #   ","")&"Bản chất: PP1 đánh giá định tính theo điều kiện, PP2 cân bằng số độ (P2 tr.331).")
    return ('Hai phương pháp thống nhất về vượng/nhược.' if X.truthy(X.eq(E.v(7,r,c-1),'Giống')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('(1) Hợp hóa khác nhau ở ',X.COUNTIF(E.rg(7,r-9,c-1,r-3,c-1), 'KHÁC')),' chỗ (xem các dòng trên) làm đổi lực lượng các hành. ') if X.truthy(X.gt(X.COUNTIF(E.rg(7,r-9,c-1,r-3,c-1), 'KHÁC'),0.0)) else ''),(X.cat(X.cat('(2) PP1 rơi vào trường hợp nửa vượng nửa nhược, phải phân xử a–d (P1 tr.71), tổng điểm ',E.v(9,88,3)),' – kết luận nhạy. ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('a–d', E.v(9,92,3)))) else '')),(X.cat(X.cat('(3) PP2 sát ngưỡng 40% (tỉ lệ ',X.TEXT(E.v(10,101,3), '0.0%')),'). ') if X.truthy(X.lt(X.ABS(X.sub(E.v(10,101,3),E.v(24,95,2))),0.05)) else '')),('(4) PP1 coi được lệnh là quyết định (P1 tr.70), PP2 chỉ cộng 20% cho hành nắm lệnh nên phe khác nhiều độ vẫn thắng. ' if X.truthy(X.AND(X.eq(E.v(9,43,3),1.0), X.eq(('nhược' if X.truthy(X.eq(E.v(8,6,4),'Nhật chủ nhược')) else 'vượng'),'nhược'))) else '')),('(4) PP1 không được lệnh thì cần ≥2 điều mạnh (P1 tr.70-71); PP2 cộng dồn độ nên nhiều chữ cùng phe vẫn đạt ≥40%. ' if X.truthy(X.AND(X.eq(E.v(9,43,3),0.0), X.eq(('nhược' if X.truthy(X.eq(E.v(8,6,4),'Nhật chủ nhược')) else 'vượng'),'vượng'))) else '')),('(5) PP1 cho trung hòa – PP2 không có khái niệm này, luôn chia vượng/nhược theo ngưỡng 40%. ' if X.truthy(X.eq(E.v(8,6,3),'Trung hòa')) else '')),'Bản chất: PP1 đánh giá định tính theo điều kiện, PP2 cân bằng số độ (P2 tr.331).'))

def f33(E, r, c):
    # Ô C16
    # = DUNG_HY_KY!$C$12&" chữ (hạng "&RANK(DUNG_HY_KY!$C$12,DUNG_HY_KY!$C$12:$C$16)&")"
    return X.cat(X.cat(X.cat(E.v(8,12,3),' chữ (hạng '),X.RANK(E.v(8,12,3), E.rg(8,12,3,16,3))),')')

def f34(E, r, c):
    # Ô D16
    # = TEXT(DUNG_HY_KY!$D$12,"0.0")&" độ (hạng "&RANK(DUNG_HY_KY!$D$12,DUNG_HY_KY!$D$12:$D$16)&")"
    return X.cat(X.cat(X.cat(X.TEXT(E.v(8,12,4), '0.0'),' độ (hạng '),X.RANK(E.v(8,12,4), E.rg(8,12,4,16,4))),')')

def f35(E, r, c):
    # Ô E16
    # = IF(RANK(DUNG_HY_KY!$C$12,DUNG_HY_KY!$C$12:$C$16)=RANK(DUNG_HY_KY!$D$12,DUNG_HY_KY!$D$12:$D$16),"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(X.RANK(E.v(8,12,3), E.rg(8,12,3,16,3)),X.RANK(E.v(8,12,4), E.rg(8,12,4,16,4)))) else 'KHÁC')

def f36(E, r, c):
    # Ô F16
    # = "PP1 đếm số chữ rõ (can lộ) + mờ (tàng can) (P2 tr.455-457); PP2 tính độ có trọng số theo vị trí, hợp hóa, xung, lệnh (P2 tr.331-355). So sánh theo th
    #   ứ hạng."
    return 'PP1 đếm số chữ rõ (can lộ) + mờ (tàng can) (P2 tr.455-457); PP2 tính độ có trọng số theo vị trí, hợp hóa, xung, lệnh (P2 tr.331-355). So sánh theo thứ hạng.'

def f37(E, r, c):
    # Ô C17
    # = DUNG_HY_KY!$C$13&" chữ (hạng "&RANK(DUNG_HY_KY!$C$13,DUNG_HY_KY!$C$12:$C$16)&")"
    return X.cat(X.cat(X.cat(E.v(8,13,3),' chữ (hạng '),X.RANK(E.v(8,13,3), E.rg(8,12,3,16,3))),')')

def f38(E, r, c):
    # Ô D17
    # = TEXT(DUNG_HY_KY!$D$13,"0.0")&" độ (hạng "&RANK(DUNG_HY_KY!$D$13,DUNG_HY_KY!$D$12:$D$16)&")"
    return X.cat(X.cat(X.cat(X.TEXT(E.v(8,13,4), '0.0'),' độ (hạng '),X.RANK(E.v(8,13,4), E.rg(8,12,4,16,4))),')')

def f39(E, r, c):
    # Ô E17
    # = IF(RANK(DUNG_HY_KY!$C$13,DUNG_HY_KY!$C$12:$C$16)=RANK(DUNG_HY_KY!$D$13,DUNG_HY_KY!$D$12:$D$16),"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(X.RANK(E.v(8,13,3), E.rg(8,12,3,16,3)),X.RANK(E.v(8,13,4), E.rg(8,12,4,16,4)))) else 'KHÁC')

def f40(E, r, c):
    # Ô F17, F18, F19, F20
    # = ""
    return ''

def f41(E, r, c):
    # Ô C18
    # = DUNG_HY_KY!$C$14&" chữ (hạng "&RANK(DUNG_HY_KY!$C$14,DUNG_HY_KY!$C$12:$C$16)&")"
    return X.cat(X.cat(X.cat(E.v(8,14,3),' chữ (hạng '),X.RANK(E.v(8,14,3), E.rg(8,12,3,16,3))),')')

def f42(E, r, c):
    # Ô D18
    # = TEXT(DUNG_HY_KY!$D$14,"0.0")&" độ (hạng "&RANK(DUNG_HY_KY!$D$14,DUNG_HY_KY!$D$12:$D$16)&")"
    return X.cat(X.cat(X.cat(X.TEXT(E.v(8,14,4), '0.0'),' độ (hạng '),X.RANK(E.v(8,14,4), E.rg(8,12,4,16,4))),')')

def f43(E, r, c):
    # Ô E18
    # = IF(RANK(DUNG_HY_KY!$C$14,DUNG_HY_KY!$C$12:$C$16)=RANK(DUNG_HY_KY!$D$14,DUNG_HY_KY!$D$12:$D$16),"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(X.RANK(E.v(8,14,3), E.rg(8,12,3,16,3)),X.RANK(E.v(8,14,4), E.rg(8,12,4,16,4)))) else 'KHÁC')

def f44(E, r, c):
    # Ô C19
    # = DUNG_HY_KY!$C$15&" chữ (hạng "&RANK(DUNG_HY_KY!$C$15,DUNG_HY_KY!$C$12:$C$16)&")"
    return X.cat(X.cat(X.cat(E.v(8,15,3),' chữ (hạng '),X.RANK(E.v(8,15,3), E.rg(8,12,3,16,3))),')')

def f45(E, r, c):
    # Ô D19
    # = TEXT(DUNG_HY_KY!$D$15,"0.0")&" độ (hạng "&RANK(DUNG_HY_KY!$D$15,DUNG_HY_KY!$D$12:$D$16)&")"
    return X.cat(X.cat(X.cat(X.TEXT(E.v(8,15,4), '0.0'),' độ (hạng '),X.RANK(E.v(8,15,4), E.rg(8,12,4,16,4))),')')

def f46(E, r, c):
    # Ô E19
    # = IF(RANK(DUNG_HY_KY!$C$15,DUNG_HY_KY!$C$12:$C$16)=RANK(DUNG_HY_KY!$D$15,DUNG_HY_KY!$D$12:$D$16),"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(X.RANK(E.v(8,15,3), E.rg(8,12,3,16,3)),X.RANK(E.v(8,15,4), E.rg(8,12,4,16,4)))) else 'KHÁC')

def f47(E, r, c):
    # Ô C20
    # = DUNG_HY_KY!$C$16&" chữ (hạng "&RANK(DUNG_HY_KY!$C$16,DUNG_HY_KY!$C$12:$C$16)&")"
    return X.cat(X.cat(X.cat(E.v(8,16,3),' chữ (hạng '),X.RANK(E.v(8,16,3), E.rg(8,12,3,16,3))),')')

def f48(E, r, c):
    # Ô D20
    # = TEXT(DUNG_HY_KY!$D$16,"0.0")&" độ (hạng "&RANK(DUNG_HY_KY!$D$16,DUNG_HY_KY!$D$12:$D$16)&")"
    return X.cat(X.cat(X.cat(X.TEXT(E.v(8,16,4), '0.0'),' độ (hạng '),X.RANK(E.v(8,16,4), E.rg(8,12,4,16,4))),')')

def f49(E, r, c):
    # Ô E20
    # = IF(RANK(DUNG_HY_KY!$C$16,DUNG_HY_KY!$C$12:$C$16)=RANK(DUNG_HY_KY!$D$16,DUNG_HY_KY!$D$12:$D$16),"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(X.RANK(E.v(8,16,3), E.rg(8,12,3,16,3)),X.RANK(E.v(8,16,4), E.rg(8,12,4,16,4)))) else 'KHÁC')

def f50(E, r, c):
    # Ô C21
    # = DUNG_HY_KY!$C$112
    return E.v(8,112,3)

def f51(E, r, c):
    # Ô D21
    # = DUNG_HY_KY!$D$112
    return E.v(8,112,4)

def f52(E, r, c):
    # Ô E21
    # = IF(DUNG_HY_KY!$C$112=DUNG_HY_KY!$D$112,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(8,112,3),E.v(8,112,4))) else 'KHÁC')

def f53(E, r, c):
    # Ô F21
    # = IF(E21="KHÁC","Nhóm nhiều nhất khác vì PP1 đếm số chữ, PP2 tính độ (xem thứ hạng các nhóm ở trên).","")
    return ('Nhóm nhiều nhất khác vì PP1 đếm số chữ, PP2 tính độ (xem thứ hạng các nhóm ở trên).' if X.truthy(X.eq(E.v(7,r,c-1),'KHÁC')) else '')

def f54(E, r, c):
    # Ô C22
    # = DUNG_HY_KY!$C$102
    return E.v(8,102,3)

def f55(E, r, c):
    # Ô D22
    # = DUNG_HY_KY!$D$102
    return E.v(8,102,4)

def f56(E, r, c):
    # Ô E22
    # = IF(DUNG_HY_KY!$C$102=DUNG_HY_KY!$D$102,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(8,102,3),E.v(8,102,4))) else 'KHÁC')

def f57(E, r, c):
    # Ô F22
    # = IF(E22="KHÁC","Điều kiện cách đặc biệt dùng lượng của từng phương pháp (hành nhiều nhất) và cục/hợp hóa của từng phương pháp (P2 tr.369-376); PP2 còn 
    #   có trường hợp độc vượng (P2 tr.353).","")
    return ('Điều kiện cách đặc biệt dùng lượng của từng phương pháp (hành nhiều nhất) và cục/hợp hóa của từng phương pháp (P2 tr.369-376); PP2 còn có trường hợp độc vượng (P2 tr.353).' if X.truthy(X.eq(E.v(7,r,c-1),'KHÁC')) else '')

def f58(E, r, c):
    # Ô C23
    # = DUNG_HY_KY!$C$106
    return E.v(8,106,3)

def f59(E, r, c):
    # Ô D23
    # = DUNG_HY_KY!$D$106
    return E.v(8,106,4)

def f60(E, r, c):
    # Ô E23
    # = IF(DUNG_HY_KY!$C$109=DUNG_HY_KY!$D$109,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(8,109,3),E.v(8,109,4))) else 'KHÁC')

def f61(E, r, c):
    # Ô F23
    # = IF(E23="Giống","",IF(E15="KHÁC","Trạng thái nhật chủ khác → dùng nhánh khác của bảng chọn dụng thần (P1 tr.108; P2 tr.380-384).",IF(E22="KHÁC","Cách c
    #   ục khác (đặc biệt/chính cách) → quy tắc chọn khác (P2 tr.381-386).",IF(DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5,"PP2 tính nhật chủ theo hóa thần (nhật can hợp
    #    hóa), PP1 giữ hành gốc.",IF(E21="KHÁC","Cùng trạng thái nhưng nhóm nhiều nhất khác (PP1 đếm chữ, PP2 tính độ) → thứ tự ưu tiên dụng khác (P1 tr.108).
    #   ","Cùng thứ tự ưu tiên nhưng nhóm có mặt khác (PP2: nhóm có độ = 0 sau hợp hóa/xung coi như không có) hoặc quy tắc ấn vượng/ấn thấu có gốc khác.")))))
    return ('' if X.truthy(X.eq(E.v(7,r,c-1),'Giống')) else ('Trạng thái nhật chủ khác → dùng nhánh khác của bảng chọn dụng thần (P1 tr.108; P2 tr.380-384).' if X.truthy(X.eq(E.v(7,r-8,c-1),'KHÁC')) else ('Cách cục khác (đặc biệt/chính cách) → quy tắc chọn khác (P2 tr.381-386).' if X.truthy(X.eq(E.v(7,r-1,c-1),'KHÁC')) else ('PP2 tính nhật chủ theo hóa thần (nhật can hợp hóa), PP1 giữ hành gốc.' if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else ('Cùng trạng thái nhưng nhóm nhiều nhất khác (PP1 đếm chữ, PP2 tính độ) → thứ tự ưu tiên dụng khác (P1 tr.108).' if X.truthy(X.eq(E.v(7,r-2,c-1),'KHÁC')) else 'Cùng thứ tự ưu tiên nhưng nhóm có mặt khác (PP2: nhóm có độ = 0 sau hợp hóa/xung coi như không có) hoặc quy tắc ấn vượng/ấn thấu có gốc khác.')))))

def f62(E, r, c):
    # Ô C24
    # = DUNG_HY_KY!$C$107
    return E.v(8,107,3)

def f63(E, r, c):
    # Ô D24
    # = DUNG_HY_KY!$D$107
    return E.v(8,107,4)

def f64(E, r, c):
    # Ô E24
    # = IF(DUNG_HY_KY!$C$110=DUNG_HY_KY!$D$110,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(8,110,3),E.v(8,110,4))) else 'KHÁC')

def f65(E, r, c):
    # Ô F24
    # = IF(E24="Giống","",IF(E15="KHÁC","Trạng thái nhật chủ khác → dùng nhánh khác của bảng chọn dụng thần (P1 tr.108; P2 tr.380-384).",IF(E22="KHÁC","Cách c
    #   ục khác (đặc biệt/chính cách) → quy tắc chọn khác (P2 tr.381-386).",IF(DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5,"PP2 tính nhật chủ theo hóa thần (nhật can hợp
    #    hóa), PP1 giữ hành gốc.",IF(E21="KHÁC","Cùng trạng thái nhưng nhóm nhiều nhất khác (PP1 đếm chữ, PP2 tính độ) → thứ tự ưu tiên dụng khác (P1 tr.108).
    #   ","Cùng thứ tự ưu tiên nhưng nhóm có mặt khác (PP2: nhóm có độ = 0 sau hợp hóa/xung coi như không có) hoặc quy tắc ấn vượng/ấn thấu có gốc khác.")))))
    return ('' if X.truthy(X.eq(E.v(7,r,c-1),'Giống')) else ('Trạng thái nhật chủ khác → dùng nhánh khác của bảng chọn dụng thần (P1 tr.108; P2 tr.380-384).' if X.truthy(X.eq(E.v(7,r-9,c-1),'KHÁC')) else ('Cách cục khác (đặc biệt/chính cách) → quy tắc chọn khác (P2 tr.381-386).' if X.truthy(X.eq(E.v(7,r-2,c-1),'KHÁC')) else ('PP2 tính nhật chủ theo hóa thần (nhật can hợp hóa), PP1 giữ hành gốc.' if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else ('Cùng trạng thái nhưng nhóm nhiều nhất khác (PP1 đếm chữ, PP2 tính độ) → thứ tự ưu tiên dụng khác (P1 tr.108).' if X.truthy(X.eq(E.v(7,r-3,c-1),'KHÁC')) else 'Cùng thứ tự ưu tiên nhưng nhóm có mặt khác (PP2: nhóm có độ = 0 sau hợp hóa/xung coi như không có) hoặc quy tắc ấn vượng/ấn thấu có gốc khác.')))))

def f66(E, r, c):
    # Ô C25
    # = DUNG_HY_KY!$C$108
    return E.v(8,108,3)

def f67(E, r, c):
    # Ô D25
    # = DUNG_HY_KY!$D$108
    return E.v(8,108,4)

def f68(E, r, c):
    # Ô E25
    # = IF(DUNG_HY_KY!$C$111=DUNG_HY_KY!$D$111,"Giống","KHÁC")
    return ('Giống' if X.truthy(X.eq(E.v(8,111,3),E.v(8,111,4))) else 'KHÁC')

def f69(E, r, c):
    # Ô F25
    # = IF(E25="Giống","",IF(E15="KHÁC","Trạng thái nhật chủ khác → dùng nhánh khác của bảng chọn dụng thần (P1 tr.108; P2 tr.380-384).",IF(E22="KHÁC","Cách c
    #   ục khác (đặc biệt/chính cách) → quy tắc chọn khác (P2 tr.381-386).",IF(DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5,"PP2 tính nhật chủ theo hóa thần (nhật can hợp
    #    hóa), PP1 giữ hành gốc.",IF(E21="KHÁC","Cùng trạng thái nhưng nhóm nhiều nhất khác (PP1 đếm chữ, PP2 tính độ) → thứ tự ưu tiên dụng khác (P1 tr.108).
    #   ","Cùng thứ tự ưu tiên nhưng nhóm có mặt khác (PP2: nhóm có độ = 0 sau hợp hóa/xung coi như không có) hoặc quy tắc ấn vượng/ấn thấu có gốc khác.")))))
    return ('' if X.truthy(X.eq(E.v(7,r,c-1),'Giống')) else ('Trạng thái nhật chủ khác → dùng nhánh khác của bảng chọn dụng thần (P1 tr.108; P2 tr.380-384).' if X.truthy(X.eq(E.v(7,r-10,c-1),'KHÁC')) else ('Cách cục khác (đặc biệt/chính cách) → quy tắc chọn khác (P2 tr.381-386).' if X.truthy(X.eq(E.v(7,r-3,c-1),'KHÁC')) else ('PP2 tính nhật chủ theo hóa thần (nhật can hợp hóa), PP1 giữ hành gốc.' if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else ('Cùng trạng thái nhưng nhóm nhiều nhất khác (PP1 đếm chữ, PP2 tính độ) → thứ tự ưu tiên dụng khác (P1 tr.108).' if X.truthy(X.eq(E.v(7,r-4,c-1),'KHÁC')) else 'Cùng thứ tự ưu tiên nhưng nhóm có mặt khác (PP2: nhóm có độ = 0 sau hợp hóa/xung coi như không có) hoặc quy tắc ấn vượng/ấn thấu có gốc khác.')))))

def f70(E, r, c):
    # Ô C31
    # = IF(OR(DUNG_HY_KY!$C$6="Quá vượng",DUNG_HY_KY!$C$6="Rất vượng",DUNG_HY_KY!$C$6="Thân vượng"),"vượng",IF(DUNG_HY_KY!$C$6="Trung hòa","trung hòa","nhược"
    #   ))
    return ('vượng' if X.truthy(X.OR(X.eq(E.v(8,6,3),'Quá vượng'), X.eq(E.v(8,6,3),'Rất vượng'), X.eq(E.v(8,6,3),'Thân vượng'))) else ('trung hòa' if X.truthy(X.eq(E.v(8,6,3),'Trung hòa')) else 'nhược'))

def f71(E, r, c):
    # Ô D31
    # = IF(PP1_DINH_TINH!$C$62=1,"KĐ",IF(OR(AND(PP1_DINH_TINH!$C$66>=2,PP1_DINH_TINH!$D$66>=2),AND(PP1_DINH_TINH!$C$66=1,PP1_DINH_TINH!$C$67=1)),"KĐ",IF(PP1_D
    #   INH_TINH!$C$66>=2,"PX","YN")))
    return ('KĐ' if X.truthy(X.eq(E.v(9,62,3),1.0)) else ('KĐ' if X.truthy(X.OR(X.AND(X.ge(E.v(9,66,3),2.0), X.ge(E.v(9,66,4),2.0)), X.AND(X.eq(E.v(9,66,3),1.0), X.eq(E.v(9,67,3),1.0)))) else ('PX' if X.truthy(X.ge(E.v(9,66,3),2.0)) else 'YN')))

def f72(E, r, c):
    # Ô E31
    # = CHOOSE(MATCH(D31,{"KĐ","PX","YN"},0),"quy tắc khẳng định","nửa vượng nửa nhược","không lệnh, ≤1 điều")
    return X.CHOOSE(X.MATCH(E.v(7,r,c-1), X.Arr([['KĐ', 'PX', 'YN']]), 0.0), lambda: 'quy tắc khẳng định', lambda: 'nửa vượng nửa nhược', lambda: 'không lệnh, ≤1 điều')

def f73(E, r, c):
    # Ô F31
    # = "PP1: "&DUNG_HY_KY!$C$6&" – "&PP1_DINH_TINH!$C$92&". Loại KĐ = quy tắc xác định của P1 tr.70–71 (được lệnh; hoặc không lệnh nhưng ≥2 điều mạnh; hoặc 1
    #    điều + ấn/tỉ cục). PX = ""nửa vượng nửa nhược … không dễ quyết định"" phải phân xử a–d (P1 tr.71). YN = không lệnh, ≤1 điều."&" (diễn giải từng chữ: 
    #   PP1_DINH_TINH dòng 104-110; KET_LUAN dòng 32)"
    return X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(8,6,3)),' – '),E.v(9,92,3)),'. Loại KĐ = quy tắc xác định của P1 tr.70–71 (được lệnh; hoặc không lệnh nhưng ≥2 điều mạnh; hoặc 1 điều + ấn/tỉ cục). PX = "nửa vượng nửa nhược … không dễ quyết định" phải phân xử a–d (P1 tr.71). YN = không lệnh, ≤1 điều.'),' (diễn giải từng chữ: PP1_DINH_TINH dòng 104-110; KET_LUAN dòng 32)')

def f74(E, r, c):
    # Ô C32
    # = IF(DUNG_HY_KY!$D$6="Nhật chủ nhược","nhược","vượng")
    return ('nhược' if X.truthy(X.eq(E.v(8,6,4),'Nhật chủ nhược')) else 'vượng')

def f75(E, r, c):
    # Ô D32
    # = PP2_DINH_LUONG!$C$101
    return E.v(10,101,3)

def f76(E, r, c):
    # Ô E32
    # = IF(PP2_DINH_LUONG!$C$101<$C$30,"xa ngưỡng (nhược rõ)",IF(PP2_DINH_LUONG!$C$101>=$D$30,"xa ngưỡng (vượng rõ)",IF(ABS(PP2_DINH_LUONG!$C$101-P_NGUONG)<$E
    #   $30,"SÁT NGƯỠNG 40%","vùng giữa")))
    return ('xa ngưỡng (nhược rõ)' if X.truthy(X.lt(E.v(10,101,3),E.v(7,30,3))) else ('xa ngưỡng (vượng rõ)' if X.truthy(X.ge(E.v(10,101,3),E.v(7,30,4))) else ('SÁT NGƯỠNG 40%' if X.truthy(X.lt(X.ABS(X.sub(E.v(10,101,3),E.v(24,95,2))),E.v(7,30,5))) else 'vùng giữa')))

def f77(E, r, c):
    # Ô F32
    # = "PP2: phe mình "&TEXT(PP2_DINH_LUONG!$C$99,"0.0")&"/"&TEXT(PP2_DINH_LUONG!$C$100,"0.0")&" độ = "&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")&"; ngưỡng 40% (P2 
    #   tr.352)."&IF((DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5)," PP2 đã đổi hành nhật chủ theo hợp hóa thành "&INDEX(HANH_TEN,DUNG_HY_KY!$D$5+1)&" – tỉ lệ này là của
    #    hóa thần.","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2: phe mình ',X.TEXT(E.v(10,99,3), '0.0')),'/'),X.TEXT(E.v(10,100,3), '0.0')),' độ = '),X.TEXT(E.v(10,101,3), '0.0%')),'; ngưỡng 40% (P2 tr.352).'),(X.cat(X.cat(' PP2 đã đổi hành nhật chủ theo hợp hóa thành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,5,4),1.0))),' – tỉ lệ này là của hóa thần.') if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else ''))

def f78(E, r, c):
    # Ô C33
    # = IF(NHAP_LIEU!$C$14="Theo PP1",IF(AND(C31=C32,DUNG_HY_KY!$C$5=DUNG_HY_KY!$D$5),0,1),IF(NHAP_LIEU!$C$14="Theo PP2",IF(AND(C31=C32,DUNG_HY_KY!$C$5=DUNG_H
    #   Y_KY!$D$5),0,2),IF(C31=C32,0,IF((DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5),1,IF(OR(PP2_DINH_LUONG!$C$101<$C$30,PP2_DINH_LUONG!$C$101>=$D$30),2,1)))))
    return ((0.0 if X.truthy(X.AND(X.eq(E.v(7,r-2,c),E.v(7,r-1,c)), X.eq(E.v(8,5,3),E.v(8,5,4)))) else 1.0) if X.truthy(X.eq(E.v(2,14,3),'Theo PP1')) else ((0.0 if X.truthy(X.AND(X.eq(E.v(7,r-2,c),E.v(7,r-1,c)), X.eq(E.v(8,5,3),E.v(8,5,4)))) else 2.0) if X.truthy(X.eq(E.v(2,14,3),'Theo PP2')) else (0.0 if X.truthy(X.eq(E.v(7,r-2,c),E.v(7,r-1,c))) else (1.0 if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else (2.0 if X.truthy(X.OR(X.lt(E.v(10,101,3),E.v(7,30,3)), X.ge(E.v(10,101,3),E.v(7,30,4)))) else 1.0)))))

def f79(E, r, c):
    # Ô D33
    # = CHOOSE(C33+1,"Cả hai (thống nhất)","Theo PP1","Theo PP2")&IF(AND(C33<>0,NHAP_LIEU!$C$14<>"Theo khuyến nghị")," (NGƯỜI DÙNG CHỌN ở NHAP_LIEU C14)","")
    return X.cat(X.CHOOSE(X.add(E.v(7,r,c-1),1.0), lambda: 'Cả hai (thống nhất)', lambda: 'Theo PP1', lambda: 'Theo PP2'),(' (NGƯỜI DÙNG CHỌN ở NHAP_LIEU C14)' if X.truthy(X.AND(X.ne(E.v(7,r,c-1),0.0), X.ne(E.v(2,14,3),'Theo khuyến nghị'))) else ''))

def f80(E, r, c):
    # Ô F33
    # = IF(AND(C33<>0,NHAP_LIEU!$C$14<>"Theo khuyến nghị"),"Người dùng chọn "&NHAP_LIEU!$C$14&" ở NHAP_LIEU C14 (thay cho khuyến nghị tự động). ","")&IF(C33=0
    #   ,"Hai phương pháp cùng kết luận – không cần chọn.",IF((DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5),"PP2 tính theo hóa thần, PP1 giữ hành gốc → hai bên không cùn
    #   g đối tượng; lấy PP1 cho vượng suy của nhật can.",IF(D31="PX","PP1 ở vùng ""nửa vượng nửa nhược"" (P1 tr.71) đã phân xử a–d (hòa thì đếm chữ) → lấy PP
    #   1 (G11, chốt 20/10/2026: trên các lá có lời giải, nhóm này PP1 đúng 7/12 lá).",IF(C33=2,"PP1 vượng nhờ quy tắc 3 (không lệnh, ≥2 điều mạnh) nhưng PP2 
    #   cho phe mình dưới mức nhược rõ → các điều kiện không ""giúp ích nhiều"" như P1 tr.70 đòi hỏi → lấy PP2.","PP1 dựa trên quy tắc "&IF(D31="KĐ","khẳng đị
    #   nh","mất lệnh – thiếu điều kiện")&" của sách → lấy PP1."))))
    return X.cat((X.cat(X.cat('Người dùng chọn ',E.v(2,14,3)),' ở NHAP_LIEU C14 (thay cho khuyến nghị tự động). ') if X.truthy(X.AND(X.ne(E.v(7,r,c-3),0.0), X.ne(E.v(2,14,3),'Theo khuyến nghị'))) else ''),('Hai phương pháp cùng kết luận – không cần chọn.' if X.truthy(X.eq(E.v(7,r,c-3),0.0)) else ('PP2 tính theo hóa thần, PP1 giữ hành gốc → hai bên không cùng đối tượng; lấy PP1 cho vượng suy của nhật can.' if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else ('PP1 ở vùng "nửa vượng nửa nhược" (P1 tr.71) đã phân xử a–d (hòa thì đếm chữ) → lấy PP1 (G11, chốt 20/10/2026: trên các lá có lời giải, nhóm này PP1 đúng 7/12 lá).' if X.truthy(X.eq(E.v(7,r-2,c-2),'PX')) else ('PP1 vượng nhờ quy tắc 3 (không lệnh, ≥2 điều mạnh) nhưng PP2 cho phe mình dưới mức nhược rõ → các điều kiện không "giúp ích nhiều" như P1 tr.70 đòi hỏi → lấy PP2.' if X.truthy(X.eq(E.v(7,r,c-3),2.0)) else X.cat(X.cat('PP1 dựa trên quy tắc ',('khẳng định' if X.truthy(X.eq(E.v(7,r-2,c-2),'KĐ')) else 'mất lệnh – thiếu điều kiện')),' của sách → lấy PP1.'))))))

def f81(E, r, c):
    # Ô C34
    # = UPPER(IF(C33=2,C32,C31))&" ("&D33&")"
    return X.cat(X.cat(X.cat(X.UPPER((E.v(7,r-2,c) if X.truthy(X.eq(E.v(7,r-1,c),2.0)) else E.v(7,r-3,c))),' ('),E.v(7,r-1,c+1)),')')

def f82(E, r, c):
    # Ô C35
    # = IF((IF(C33=2,C32,C31)="nhược")<>(C32="nhược"),"Thấp",IF(OR(PP1_DINH_TINH!$C$117=2,AND(PP2_DINH_LUONG!$C$101>=0.35,PP2_DINH_LUONG!$C$101<0.45)),"Vừa","
    #   Cao"))
    return ('Thấp' if X.truthy(X.ne(X.eq((E.v(7,r-3,c) if X.truthy(X.eq(E.v(7,r-2,c),2.0)) else E.v(7,r-4,c)),'nhược'),X.eq(E.v(7,r-3,c),'nhược'))) else ('Vừa' if X.truthy(X.OR(X.eq(E.v(9,117,3),2.0), X.AND(X.ge(E.v(10,101,3),0.35), X.lt(E.v(10,101,3),0.45)))) else 'Cao'))

def f83(E, r, c):
    # Ô F35
    # = CHOOSE(MATCH(C35,{"Cao","Vừa","Thấp"},0),"Cao: hai phương pháp cùng kết luận, không cận ranh giới. Trên 68 lá số của sách: đúng 34/35.","Vừa: hai phươ
    #   ng pháp cùng kết luận nhưng vượng cận ranh giới hoặc PP2 sát ngưỡng (35–45%) – nên đối chiếu thêm với sự việc đã xảy ra.","Thấp: hai phương pháp khác 
    #   kết luận – kết luận chính theo khuyến nghị, xem thêm cách đọc còn lại; nên kiểm nghiệm bằng sự việc đã qua.")
    return X.CHOOSE(X.MATCH(E.v(7,r,c-3), X.Arr([['Cao', 'Vừa', 'Thấp']]), 0.0), lambda: 'Cao: hai phương pháp cùng kết luận, không cận ranh giới. Trên 68 lá số của sách: đúng 34/35.', lambda: 'Vừa: hai phương pháp cùng kết luận nhưng vượng cận ranh giới hoặc PP2 sát ngưỡng (35–45%) – nên đối chiếu thêm với sự việc đã xảy ra.', lambda: 'Thấp: hai phương pháp khác kết luận – kết luận chính theo khuyến nghị, xem thêm cách đọc còn lại; nên kiểm nghiệm bằng sự việc đã qua.')

def f84(E, r, c):
    # Ô C36
    # = "DỮ KIỆN: PP1 = "&DUNG_HY_KY!$C$6&" ("&PP1_DINH_TINH!$C$92&"); PP2 = phe mình "&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")&" → "&DUNG_HY_KY!$D$6&". "&IF(C33=0
    #   ,"QUY TẮC: hai phương pháp cùng kết luận nên không phải chọn. ",IF((DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5),"QUY TẮC: PP2 coi nhật can đã hợp hóa (P2 tr.336
    #   -338) nên tỉ lệ là của hóa thần; trong chính cách sách vẫn luận nhật can theo hành gốc (ví dụ P2 tr.415: Đinh hợp Nhâm nhưng sách kết luận ""nhật trụ 
    #   vẫn yếu"") → theo PP1; nếu đủ điều kiện các"&"h hóa thì xem DUNG_HY_KY. ",IF(C33=2,"QUY TẮC 6.6: PP2 đo được phe mình "&TEXT(PP2_DINH_LUONG!$C$101,"0.
    #   0%")&" – xa ngưỡng 40% (dưới "&TEXT($C$30,"0%")&" là nhược rõ, từ "&TEXT($D$30,"0%")&" là vượng rõ), lực đo được đủ rõ nên theo PP2. Kiểm chứng 15/11/
    #   2026 trên 431 lá có lời giải vượng suy: 65 lá Thiệu Vĩ Hoa giữ nguyên 57/65; 366 lá các sách khác tăng từ 242 lên 288 (66% → 79%). Sách định nghĩa ""c
    #   an ngày mất lệnh … Nếu quan sát nặng quá … Thực thương nhiều quá thì can ngày hao tổn cũng gọi là suy nhược"" (P2 tr.330). ",IF(FALSE,"QUY TẮC: PP1 kế
    #   t luận vượng theo quy tắc 3 của P1 tr.70: không được lệnh ""nhưng lại có được hai điều trở lên … những điều đó lại mạnh, giúp ích nhiều"". PP2 đo được
    #    phe mình chỉ "&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")&" (dưới mức nhược rõ "&TEXT($C$30,"0%")&"), tức các điều đó không giúp ích nhiều; sách định nghĩa "
    #   "can ngày mất lệnh … Nếu quan sát nặng quá … Thực thương nhiều quá thì can ngày hao tổn cũng gọi là suy nhược"" (P2 tr.330). Trên các lá của sách thuộ
    #   c loại này, 4/7 lá tác giả kết l"&"uận nhược → theo PP2. ",IF(D31="KĐ","QUY TẮC: P1 tr.70 viết ""Đầu tiên can ngày được lệnh, đó là điều quan trọng nh
    #   ất để phán đoán thân vượng … được một trong ba điều … thì có thể khẳng định là thân vượng""; PP1 ở đây rơi vào một quy tắc khẳng định như vậy, còn tro
    #   ng PP2 lệnh "&"tháng chỉ làm tăng/giảm 1/5 độ (P2 tr.345) nên sức của lệnh có thể chưa đủ → theo PP1. ",IF(D31="PX","QUY TẮC: PP1 rơi vào trường hợp "
    #   "nửa vượng, nửa nhược cân bằng nhau"" (P1 tr.71); công cụ phân xử theo a–d (tổng "&PP1_DINH_TINH!$C$88&"), nếu hòa thì ""so sánh mức độ"" bằng đếm chữ
    #    (P1 tr.71; P2 tr.455-457) → theo PP1 (kiểm chứng trên lá có lời giải: nhóm này PP1 đúng 7/12 lá). ","QUY TẮC: nhật can mất lệnh và có không quá 1 tro
    #   ng 3 điều (đắc địa, được sinh, được trợ) – đúng mô tả ""Can ngày suy nhược: tức là can ngày mất lệnh, không đắc địa, không có trợ giúp, không được sin
    #   h"" (P2 tr.330); PP2 cộng dồn độ của nhiều"&" chữ yếu nên có thể vượt 40% → theo PP1 nhưng độ tin thấp. "))))))&"KẾT LUẬN: nhật chủ "&IF(C33=2,C32,C31
    #   )&" – độ tin "&LOWER(C35)&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('DỮ KIỆN: PP1 = ',E.v(8,6,3)),' ('),E.v(9,92,3)),'); PP2 = phe mình '),X.TEXT(E.v(10,101,3), '0.0%')),' → '),E.v(8,6,4)),'. '),('QUY TẮC: hai phương pháp cùng kết luận nên không phải chọn. ' if X.truthy(X.eq(E.v(7,r-3,c),0.0)) else (X.cat('QUY TẮC: PP2 coi nhật can đã hợp hóa (P2 tr.336-338) nên tỉ lệ là của hóa thần; trong chính cách sách vẫn luận nhật can theo hành gốc (ví dụ P2 tr.415: Đinh hợp Nhâm nhưng sách kết luận "nhật trụ vẫn yếu") → theo PP1; nếu đủ điều kiện các','h hóa thì xem DUNG_HY_KY. ') if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('QUY TẮC 6.6: PP2 đo được phe mình ',X.TEXT(E.v(10,101,3), '0.0%')),' – xa ngưỡng 40% (dưới '),X.TEXT(E.v(7,30,3), '0%')),' là nhược rõ, từ '),X.TEXT(E.v(7,30,4), '0%')),' là vượng rõ), lực đo được đủ rõ nên theo PP2. Kiểm chứng 15/11/2026 trên 431 lá có lời giải vượng suy: 65 lá Thiệu Vĩ Hoa giữ nguyên 57/65; 366 lá các sách khác tăng từ 242 lên 288 (66% → 79%). Sách định nghĩa "can ngày mất lệnh … Nếu quan sát nặng quá … Thực thương nhiều quá thì can ngày hao tổn cũng gọi là suy nhược" (P2 tr.330). ') if X.truthy(X.eq(E.v(7,r-3,c),2.0)) else (X.cat(X.cat(X.cat(X.cat(X.cat('QUY TẮC: PP1 kết luận vượng theo quy tắc 3 của P1 tr.70: không được lệnh "nhưng lại có được hai điều trở lên … những điều đó lại mạnh, giúp ích nhiều". PP2 đo được phe mình chỉ ',X.TEXT(E.v(10,101,3), '0.0%')),' (dưới mức nhược rõ '),X.TEXT(E.v(7,30,3), '0%')),'), tức các điều đó không giúp ích nhiều; sách định nghĩa "can ngày mất lệnh … Nếu quan sát nặng quá … Thực thương nhiều quá thì can ngày hao tổn cũng gọi là suy nhược" (P2 tr.330). Trên các lá của sách thuộc loại này, 4/7 lá tác giả kết l'),'uận nhược → theo PP2. ') if X.truthy(False) else (X.cat('QUY TẮC: P1 tr.70 viết "Đầu tiên can ngày được lệnh, đó là điều quan trọng nhất để phán đoán thân vượng … được một trong ba điều … thì có thể khẳng định là thân vượng"; PP1 ở đây rơi vào một quy tắc khẳng định như vậy, còn trong PP2 lệnh ','tháng chỉ làm tăng/giảm 1/5 độ (P2 tr.345) nên sức của lệnh có thể chưa đủ → theo PP1. ') if X.truthy(X.eq(E.v(7,r-5,c+1),'KĐ')) else (X.cat(X.cat('QUY TẮC: PP1 rơi vào trường hợp "nửa vượng, nửa nhược cân bằng nhau" (P1 tr.71); công cụ phân xử theo a–d (tổng ',E.v(9,88,3)),'), nếu hòa thì "so sánh mức độ" bằng đếm chữ (P1 tr.71; P2 tr.455-457) → theo PP1 (kiểm chứng trên lá có lời giải: nhóm này PP1 đúng 7/12 lá). ') if X.truthy(X.eq(E.v(7,r-5,c+1),'PX')) else X.cat('QUY TẮC: nhật can mất lệnh và có không quá 1 trong 3 điều (đắc địa, được sinh, được trợ) – đúng mô tả "Can ngày suy nhược: tức là can ngày mất lệnh, không đắc địa, không có trợ giúp, không được sinh" (P2 tr.330); PP2 cộng dồn độ của nhiều',' chữ yếu nên có thể vượt 40% → theo PP1 nhưng độ tin thấp. ')))))))),'KẾT LUẬN: nhật chủ '),(E.v(7,r-4,c) if X.truthy(X.eq(E.v(7,r-3,c),2.0)) else E.v(7,r-5,c))),' – độ tin '),X.LOWER(E.v(7,r-1,c))),'.')

def f85(E, r, c):
    # Ô C37
    # = IF(C33=2,"Dụng "&DUNG_HY_KY!$D$106&" | Hỷ "&DUNG_HY_KY!$D$107&" | Kỵ "&DUNG_HY_KY!$D$108&" (theo PP2)","Dụng "&DUNG_HY_KY!$C$106&" | Hỷ "&DUNG_HY_KY!$
    #   C$107&" | Kỵ "&DUNG_HY_KY!$C$108&IF(C33=1," (theo PP1)",IF(DUNG_HY_KY!$C$109=DUNG_HY_KY!$D$109," (hai PP thống nhất)"," (PP1; PP2 cùng trạng thái nhưn
    #   g dụng "&DUNG_HY_KY!$D$106&" – nhóm nhiều nhất khác nhau, xem dòng DỤNG THẦN)")))
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dụng ',E.v(8,106,4)),' | Hỷ '),E.v(8,107,4)),' | Kỵ '),E.v(8,108,4)),' (theo PP2)') if X.truthy(X.eq(E.v(7,r-4,c),2.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dụng ',E.v(8,106,3)),' | Hỷ '),E.v(8,107,3)),' | Kỵ '),E.v(8,108,3)),(' (theo PP1)' if X.truthy(X.eq(E.v(7,r-4,c),1.0)) else (' (hai PP thống nhất)' if X.truthy(X.eq(E.v(8,109,3),E.v(8,109,4))) else X.cat(X.cat(' (PP1; PP2 cùng trạng thái nhưng dụng ',E.v(8,106,4)),' – nhóm nhiều nhất khác nhau, xem dòng DỤNG THẦN)')))))

def f86(E, r, c):
    # Ô C42
    # = IF(C31=C32,"THỐNG NHẤT: cả hai cho nhật chủ "&UPPER(C31)&" – không cần chọn; các dòng dưới chỉ để hiểu cách hai phương pháp đi đến cùng kết luận.","KH
    #   ÁC NHAU: PP1 cho nhật chủ "&UPPER(C31)&", PP2 cho "&UPPER(C32)&". Các dòng 2–4 chỉ ra nguyên nhân, dòng 5–7 là khuyến nghị và cách kiểm chứng.")
    return (X.cat(X.cat('THỐNG NHẤT: cả hai cho nhật chủ ',X.UPPER(E.v(7,r-11,c))),' – không cần chọn; các dòng dưới chỉ để hiểu cách hai phương pháp đi đến cùng kết luận.') if X.truthy(X.eq(E.v(7,r-11,c),E.v(7,r-10,c))) else X.cat(X.cat(X.cat(X.cat('KHÁC NHAU: PP1 cho nhật chủ ',X.UPPER(E.v(7,r-11,c))),', PP2 cho '),X.UPPER(E.v(7,r-10,c))),'. Các dòng 2–4 chỉ ra nguyên nhân, dòng 5–7 là khuyến nghị và cách kiểm chứng.'))

def f87(E, r, c):
    # Ô C43
    # = "PP1: "&IF(PP1_DINH_TINH!$C$43=1,"nhật can ĐƯỢC LỆNH ("&TRUNG_GIAN!$D$29&" tại tháng "&TRUNG_GIAN!$D$11&") – sách coi lệnh là tiêu chí quyết định hàng
    #    đầu: được lệnh + 1 điều là vượng (P1 tr.70).","nhật can KHÔNG được lệnh ("&TRUNG_GIAN!$D$29&" tại tháng "&TRUNG_GIAN!$D$11&") – muốn vượng phải có ≥ 
    #   2 điều mạnh (P1 tr.70-71).")&" PP2: lệnh chỉ là hệ số nhân – hành nhật chủ ×"&INDEX(PP2_DINH_LUONG!$E$91:$E$95,TRUNG_GIAN!$E$7+1)&", hành ấn ×"&INDEX(
    #   PP2_DINH_LUONG!$E$91:$E$95,MOD(TRUNG_GIAN!$E$7+4,5)+1)&" (P2 tr.345-346), nên lệnh không quyết định một mình."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',(X.cat(X.cat(X.cat(X.cat('nhật can ĐƯỢC LỆNH (',E.v(23,29,4)),' tại tháng '),E.v(23,11,4)),') – sách coi lệnh là tiêu chí quyết định hàng đầu: được lệnh + 1 điều là vượng (P1 tr.70).') if X.truthy(X.eq(E.v(9,43,3),1.0)) else X.cat(X.cat(X.cat(X.cat('nhật can KHÔNG được lệnh (',E.v(23,29,4)),' tại tháng '),E.v(23,11,4)),') – muốn vượng phải có ≥ 2 điều mạnh (P1 tr.70-71).'))),' PP2: lệnh chỉ là hệ số nhân – hành nhật chủ ×'),X.INDEX(E.rg(10,91,5,95,5), X.add(E.v(23,7,5),1.0))),', hành ấn ×'),X.INDEX(E.rg(10,91,5,95,5), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))),' (P2 tr.345-346), nên lệnh không quyết định một mình.')

def f88(E, r, c):
    # Ô C44
    # = "PP1 đếm có/không và mạnh/yếu: "&C14&". PP2 cộng độ mọi chữ: "&D14&" (ngưỡng 40%). "&IF(IF(C31="nhược",0,IF(C31="trung hòa",1,2))<IF(C32="nhược",0,IF(
    #   C32="trung hòa",1,2)),"Ở lá số này PP1 THẤP hơn PP2 vì trong "&PP1_DINH_TINH!$C$66&" điều phe mình có mặt "&IF(PP1_DINH_TINH!$D$66=0,"không điều nào "
    #   "mạnh""","chỉ "&PP1_DINH_TINH!$D$66&" điều ""mạnh""")&" (bị xung, ở xa, trung/dư khí, không vượng tại lệnh – xem PP1_DINH_TINH dòng 105-107) nên PP1 g
    #   ần như không tính; PP2 vẫn cộng đủ số độ của chúng.",IF(IF(C31="nhược",0,IF(C31="trung hòa",1,2))>IF(C32="nhược",0,IF(C32="trung hòa",1,2)),"Ở lá số n
    #   ày PP1 CAO hơn PP2 vì "&IF(PP1_DINH_TINH!$C$43=1,"nhật can được lệnh – PP1 coi lệnh là quyết định","có điều mạnh")&", trong khi tổng độ phe khắc – tiế
    #   t – hao lớn nên PP2 chỉ được "&TEXT(D32,"0.0%")&".","Hai cách đo cho cùng kết luận."))&""&CHAR(10)&"Về đắc địa: "&PP1_DINH_TINH!$C$47
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1 đếm có/không và mạnh/yếu: ',E.v(7,r-30,c)),'. PP2 cộng độ mọi chữ: '),E.v(7,r-30,c+1)),' (ngưỡng 40%). '),(X.cat(X.cat(X.cat(X.cat('Ở lá số này PP1 THẤP hơn PP2 vì trong ',E.v(9,66,3)),' điều phe mình có mặt '),('không điều nào "mạnh"' if X.truthy(X.eq(E.v(9,66,4),0.0)) else X.cat(X.cat('chỉ ',E.v(9,66,4)),' điều "mạnh"'))),' (bị xung, ở xa, trung/dư khí, không vượng tại lệnh – xem PP1_DINH_TINH dòng 105-107) nên PP1 gần như không tính; PP2 vẫn cộng đủ số độ của chúng.') if X.truthy(X.lt((0.0 if X.truthy(X.eq(E.v(7,r-13,c),'nhược')) else (1.0 if X.truthy(X.eq(E.v(7,r-13,c),'trung hòa')) else 2.0)),(0.0 if X.truthy(X.eq(E.v(7,r-12,c),'nhược')) else (1.0 if X.truthy(X.eq(E.v(7,r-12,c),'trung hòa')) else 2.0)))) else (X.cat(X.cat(X.cat(X.cat('Ở lá số này PP1 CAO hơn PP2 vì ',('nhật can được lệnh – PP1 coi lệnh là quyết định' if X.truthy(X.eq(E.v(9,43,3),1.0)) else 'có điều mạnh')),', trong khi tổng độ phe khắc – tiết – hao lớn nên PP2 chỉ được '),X.TEXT(E.v(7,r-12,c+1), '0.0%')),'.') if X.truthy(X.gt((0.0 if X.truthy(X.eq(E.v(7,r-13,c),'nhược')) else (1.0 if X.truthy(X.eq(E.v(7,r-13,c),'trung hòa')) else 2.0)),(0.0 if X.truthy(X.eq(E.v(7,r-12,c),'nhược')) else (1.0 if X.truthy(X.eq(E.v(7,r-12,c),'trung hòa')) else 2.0)))) else 'Hai cách đo cho cùng kết luận.'))),''),X.CHAR(10.0)),'Về đắc địa: '),E.v(9,47,3))

def f89(E, r, c):
    # Ô C45
    # = IF(TRIM(IF(E6="KHÁC",B6&": PP1 "&C6&" / PP2 "&D6&"; ","")&IF(E7="KHÁC",B7&": PP1 "&C7&" / PP2 "&D7&"; ","")&IF(E8="KHÁC",B8&": PP1 "&C8&" / PP2 "&D8&"
    #   ; ","")&IF(E9="KHÁC",B9&": PP1 "&C9&" / PP2 "&D9&"; ","")&IF(E10="KHÁC",B10&": PP1 "&C10&" / PP2 "&D10&"; ","")&IF(E11="KHÁC",B11&": PP1 "&C11&" / PP2
    #    "&D11&"; ","")&IF(E12="KHÁC",B12&": PP1 "&C12&" / PP2 "&D12&"; ",""))="","Hợp, hóa, xung được hai phương pháp xử lý giống nhau ở lá số này (dòng 6-12
    #   ) → không phải nguyên nhân khác biệt.","Khác nhau ở: "&TRIM(IF(E6="KHÁC",B6&": PP1 "&C6&" / PP2 "&D6&"; ","")&IF(E7="KHÁC",B7&": PP1 "&C7&" / PP2 "&D7
    #   &"; ","")&IF(E8="KHÁC",B8&": PP1 "&C8&" / PP2 "&D8&"; ","")&IF(E9="KHÁC",B9&": PP1 "&C9&" / PP2 "&D9&"; ","")&IF(E10="KHÁC",B10&": PP1 "&C10&" / PP2 "
    #   &D10&"; ","")&IF(E11="KHÁC",B11&": PP1 "&C11&" / PP2 "&D11&"; ","")&IF(E12="KHÁC",B12&": PP1 "&C12&" / PP2 "&D12&"; ",""))&" (Phần 1 cần can lộ hóa th
    #   ần mới hóa – P1 tr.62; Phần 2 xét hóa thần cùng loại, khử bì – P2 tr.336-344; xem MAU_THUAN).")
    return ('Hợp, hóa, xung được hai phương pháp xử lý giống nhau ở lá số này (dòng 6-12) → không phải nguyên nhân khác biệt.' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-39,c-1),': PP1 '),E.v(7,r-39,c)),' / PP2 '),E.v(7,r-39,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-39,c+2),'KHÁC')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-38,c-1),': PP1 '),E.v(7,r-38,c)),' / PP2 '),E.v(7,r-38,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-38,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-37,c-1),': PP1 '),E.v(7,r-37,c)),' / PP2 '),E.v(7,r-37,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-37,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-36,c-1),': PP1 '),E.v(7,r-36,c)),' / PP2 '),E.v(7,r-36,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-36,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-35,c-1),': PP1 '),E.v(7,r-35,c)),' / PP2 '),E.v(7,r-35,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-35,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-34,c-1),': PP1 '),E.v(7,r-34,c)),' / PP2 '),E.v(7,r-34,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-34,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-33,c-1),': PP1 '),E.v(7,r-33,c)),' / PP2 '),E.v(7,r-33,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-33,c+2),'KHÁC')) else ''))),'')) else X.cat(X.cat('Khác nhau ở: ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-39,c-1),': PP1 '),E.v(7,r-39,c)),' / PP2 '),E.v(7,r-39,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-39,c+2),'KHÁC')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-38,c-1),': PP1 '),E.v(7,r-38,c)),' / PP2 '),E.v(7,r-38,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-38,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-37,c-1),': PP1 '),E.v(7,r-37,c)),' / PP2 '),E.v(7,r-37,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-37,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-36,c-1),': PP1 '),E.v(7,r-36,c)),' / PP2 '),E.v(7,r-36,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-36,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-35,c-1),': PP1 '),E.v(7,r-35,c)),' / PP2 '),E.v(7,r-35,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-35,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-34,c-1),': PP1 '),E.v(7,r-34,c)),' / PP2 '),E.v(7,r-34,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-34,c+2),'KHÁC')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(7,r-33,c-1),': PP1 '),E.v(7,r-33,c)),' / PP2 '),E.v(7,r-33,c+1)),'; ') if X.truthy(X.eq(E.v(7,r-33,c+2),'KHÁC')) else '')))),' (Phần 1 cần can lộ hóa thần mới hóa – P1 tr.62; Phần 2 xét hóa thần cùng loại, khử bì – P2 tr.336-344; xem MAU_THUAN).'))

def f90(E, r, c):
    # Ô C46
    # = "PP1: "&IF(PP1_DINH_TINH!$C$111<>"","được lệnh nên vượng, nhưng phân xử a–d và đếm chữ đều nghiêng về phe khắc (vượng cận ranh giới) → KÉM chắc chắn",
    #   IF(D31="PX","kết luận nằm trong vùng ""nửa vượng nửa nhược"" phải phân xử a–d → KÉM chắc chắn","kết luận theo quy tắc trực tiếp ("&E31&") → chắc chắn 
    #   hơn"))&". PP2: phe mình "&TEXT(D32,"0.0%")&IF(E32="SÁT NGƯỠNG 40%"," – sát ngưỡng 40% → KÉM chắc chắn"," – cách ngưỡng 40% rõ → chắc chắn hơn")&"."&IF
    #   (PP1_DINH_TINH!$C$111<>"",CHAR(10)&PP1_DINH_TINH!$C$111,"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',('được lệnh nên vượng, nhưng phân xử a–d và đếm chữ đều nghiêng về phe khắc (vượng cận ranh giới) → KÉM chắc chắn' if X.truthy(X.ne(E.v(9,111,3),'')) else ('kết luận nằm trong vùng "nửa vượng nửa nhược" phải phân xử a–d → KÉM chắc chắn' if X.truthy(X.eq(E.v(7,r-15,c+1),'PX')) else X.cat(X.cat('kết luận theo quy tắc trực tiếp (',E.v(7,r-15,c+2)),') → chắc chắn hơn')))),'. PP2: phe mình '),X.TEXT(E.v(7,r-14,c+1), '0.0%')),(' – sát ngưỡng 40% → KÉM chắc chắn' if X.truthy(X.eq(E.v(7,r-14,c+2),'SÁT NGƯỠNG 40%')) else ' – cách ngưỡng 40% rõ → chắc chắn hơn')),'.'),(X.cat(X.CHAR(10.0),E.v(9,111,3)) if X.truthy(X.ne(E.v(9,111,3),'')) else ''))

def f91(E, r, c):
    # Ô C47
    # = "Khuyến nghị: "&C34&". Lý do: "&F33&" Độ tin: "&C35&". "&IF(DUNG_HY_KY!$C$106=DUNG_HY_KY!$D$106,"Dụng thần hai phương pháp GIỐNG nhau ("&DUNG_HY_KY!$C
    #   $106&") nên dù chọn bên nào thì các kết luận ứng dụng (phương vị, màu, nghề, vận tốt xấu) không đổi.","Dụng thần KHÁC nhau (PP1 "&DUNG_HY_KY!$C$106&",
    #    PP2 "&DUNG_HY_KY!$D$106&") → mọi kết luận về vận, năm, phương vị đổi theo lựa chọn; toàn bộ công cụ đang dùng "&D33&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Khuyến nghị: ',E.v(7,r-13,c)),'. Lý do: '),E.v(7,r-14,c+3)),' Độ tin: '),E.v(7,r-12,c)),'. '),(X.cat(X.cat('Dụng thần hai phương pháp GIỐNG nhau (',E.v(8,106,3)),') nên dù chọn bên nào thì các kết luận ứng dụng (phương vị, màu, nghề, vận tốt xấu) không đổi.') if X.truthy(X.eq(E.v(8,106,3),E.v(8,106,4))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dụng thần KHÁC nhau (PP1 ',E.v(8,106,3)),', PP2 '),E.v(8,106,4)),') → mọi kết luận về vận, năm, phương vị đổi theo lựa chọn; toàn bộ công cụ đang dùng '),E.v(7,r-14,c+1)),'.')))

def f92(E, r, c):
    # Ô C48
    # = IF(TRIM(IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(1,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(2,LUU_NIEN!$CD$
    #   30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(3,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH
    #   (4,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(5,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$
    #   30:$CC$129,MATCH(6,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(7,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(IND
    #   EX(LUU_NIEN!$CC$30:$CC$129,MATCH(8,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(9,LUU_NIEN!$CD$30:$CD$129,0))&"; "
    #   ,"")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(10,LUU_NIEN!$CD$30:$CD$129,0))&"; ",""))="","Không có năm nào đã qua mà hai phương pháp chấm năm khác
    #    nhau → lựa chọn không ảnh hưởng chấm năm.","Các năm đã qua hai phương pháp chấm KHÁC nhau (tối đa 10 năm đầu): "&TRIM(IFERROR(INDEX(LUU_NIEN!$CC$30:$
    #   CC$129,MATCH(1,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(2,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(L
    #   UU_NIEN!$CC$30:$CC$129,MATCH(3,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(4,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")
    #   &IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(5,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(6,LUU_NIEN!$CD$30:$CD$
    #   129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(7,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(8,LUU_
    #   NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$129,MATCH(9,LUU_NIEN!$CD$30:$CD$129,0))&"; ","")&IFERROR(INDEX(LUU_NIEN!$CC$30:$CC$
    #   129,MATCH(10,LUU_NIEN!$CD$30:$CD$129,0))&"; ",""))&" Hãy nhớ lại năm đó thực tế tốt hay xấu: khớp PP1 nhiều hơn thì theo PP1, khớp PP2 thì theo PP2 (đ
    #   ặt ô DOI_CHIEU!C33 bằng tay nếu cần – xem giả định G11).")
    return ('Không có năm nào đã qua mà hai phương pháp chấm năm khác nhau → lựa chọn không ảnh hưởng chấm năm.' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(1.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: ''),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(2.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(3.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(4.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(5.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(6.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(7.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(8.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(9.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(10.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: ''))),'')) else X.cat(X.cat('Các năm đã qua hai phương pháp chấm KHÁC nhau (tối đa 10 năm đầu): ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(1.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: ''),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(2.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(3.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(4.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(5.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(6.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(7.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(8.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(9.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,81,129,81), X.MATCH(10.0, E.rg(15,30,82,129,82), 0.0)),'; '), lambda: '')))),' Hãy nhớ lại năm đó thực tế tốt hay xấu: khớp PP1 nhiều hơn thì theo PP1, khớp PP2 thì theo PP2 (đặt ô DOI_CHIEU!C33 bằng tay nếu cần – xem giả định G11).'))

def f93(E, r, c):
    # Ô C49
    # = IF(C31=C32,"Hai phương pháp thống nhất: nhật chủ "&UPPER(C31)&"; dùng kết luận chung.","Hai phương pháp khác nhau; công cụ khuyến nghị "&D33&" (nhật c
    #   hủ "&C34&", độ tin "&LOWER(C35)&"). Nếu các năm đã qua ở dòng 7 khớp phương pháp còn lại nhiều hơn thì nên đổi theo phương pháp đó.")&IF(COUNTIF(SU_KI
    #   EN!$V$6:$V$35,1)+COUNTIF(SU_KIEN!$V$6:$V$35,-1)>=2,CHAR(10)&"Sự việc đã nhập (SU_KIEN): "&SU_KIEN!$B$38,"")
    return X.cat((X.cat(X.cat('Hai phương pháp thống nhất: nhật chủ ',X.UPPER(E.v(7,r-18,c))),'; dùng kết luận chung.') if X.truthy(X.eq(E.v(7,r-18,c),E.v(7,r-17,c))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Hai phương pháp khác nhau; công cụ khuyến nghị ',E.v(7,r-16,c+1)),' (nhật chủ '),E.v(7,r-15,c)),', độ tin '),X.LOWER(E.v(7,r-14,c))),'). Nếu các năm đã qua ở dòng 7 khớp phương pháp còn lại nhiều hơn thì nên đổi theo phương pháp đó.')),(X.cat(X.cat(X.CHAR(10.0),'Sự việc đã nhập (SU_KIEN): '),E.v(32,38,2)) if X.truthy(X.ge(X.add(X.COUNTIF(E.rg(32,6,22,35,22), 1.0),X.COUNTIF(E.rg(32,6,22,35,22), X.neg(1.0))),2.0)) else ''))

def f94(E, r, c):
    # Ô C50
    # = PP1_DINH_TINH!$C$112
    return E.v(9,112,3)

def f95(E, r, c):
    # Ô C51
    # = IF(C31=C32,"Hai phương pháp thống nhất – không phải chọn.","HAI PHƯƠNG PHÁP CHO KẾT QUẢ KHÁC NHAU (PP1: "&UPPER(C31)&"; PP2: "&UPPER(C32)&"). Lựa chọn
    #    "&D33&" ở dòng 33-34 là QUY ƯỚC KỸ THUẬT (giả định G11 – MAU_THUAN), không phải quy tắc của sách; muốn dùng một phương pháp cố định thì chọn Theo PP1
    #    / Theo PP2 ở NHAP_LIEU C14.")
    return ('Hai phương pháp thống nhất – không phải chọn.' if X.truthy(X.eq(E.v(7,r-20,c),E.v(7,r-19,c))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('HAI PHƯƠNG PHÁP CHO KẾT QUẢ KHÁC NHAU (PP1: ',X.UPPER(E.v(7,r-20,c))),'; PP2: '),X.UPPER(E.v(7,r-19,c))),'). Lựa chọn '),E.v(7,r-18,c+1)),' ở dòng 33-34 là QUY ƯỚC KỸ THUẬT (giả định G11 – MAU_THUAN), không phải quy tắc của sách; muốn dùng một phương pháp cố định thì chọn Theo PP1 / Theo PP2 ở NHAP_LIEU C14.'))

def f96(E, r, c):
    # Ô C52
    # = IF(C35<>"Thấp","","Tham khảo thêm – nếu đọc theo "&IF(C33=2,"PP1 (xét lệnh trước)","PP2 (tính lực)")&": thân "&UPPER(IF(C33=2,C31,C32))&", dụng thần "
    #   &IF(C33=2,DUNG_HY_KY!$C$106,DUNG_HY_KY!$D$106)&", hỷ thần "&IF(C33=2,DUNG_HY_KY!$C$107,DUNG_HY_KY!$D$107)&"; các đại vận thuận theo cách đọc này: "&IF
    #   (IF(C33=2,TRIM(IF(DAI_VAN!C$54="Tốt",DAI_VAN!C$35&" ("&DAI_VAN!C$36&"–"&(DAI_VAN!C$36+9)&") ","")&IF(DAI_VAN!D$54="Tốt",DAI_VAN!D$35&" ("&DAI_VAN!D$36
    #   &"–"&(DAI_VAN!D$36+9)&") ","")&IF(DAI_VAN!E$54="Tốt",DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&") ","")&IF(DAI_VAN!F$54="Tốt",DAI_VAN!F$35&"
    #    ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&") ","")&IF(DAI_VAN!G$54="Tốt",DAI_VAN!G$35&" ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&") ","")&IF(DAI_VAN!H$54="Tốt
    #   ",DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI_VAN!H$36+9)&") ","")&IF(DAI_VAN!I$54="Tốt",DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&") ","")&IF(D
    #   AI_VAN!J$54="Tốt",DAI_VAN!J$35&" ("&DAI_VAN!J$36&"–"&(DAI_VAN!J$36+9)&") ","")),TRIM(IF(DAI_VAN!C$69="Tốt",DAI_VAN!C$35&" ("&DAI_VAN!C$36&"–"&(DAI_VAN
    #   !C$36+9)&") ","")&IF(DAI_VAN!D$69="Tốt",DAI_VAN!D$35&" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&") ","")&IF(DAI_VAN!E$69="Tốt",DAI_VAN!E$35&" ("&DAI_VAN!E
    #   $36&"–"&(DAI_VAN!E$36+9)&") ","")&IF(DAI_VAN!F$69="Tốt",DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&") ","")&IF(DAI_VAN!G$69="Tốt",DAI_VAN!G$3
    #   5&" ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&") ","")&IF(DAI_VAN!H$69="Tốt",DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI_VAN!H$36+9)&") ","")&IF(DAI_VAN!I$69="
    #   Tốt",DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&") ","")&IF(DAI_VAN!J$69="Tốt",DAI_VAN!J$35&" ("&DAI_VAN!J$36&"–"&(DAI_VAN!J$36+9)&") ","")))
    #   ="","không có",IF(C33=2,TRIM(IF(DAI_VAN!C$54="Tốt",DAI_VAN!C$35&" ("&DAI_VAN!C$36&"–"&(DAI_VAN!C$36+9)&") ","")&IF(DAI_VAN!D$54="Tốt",DAI_VAN!D$35&" (
    #   "&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&") ","")&IF(DAI_VAN!E$54="Tốt",DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&") ","")&IF(DAI_VAN!F$54="Tốt",
    #   DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&") ","")&IF(DAI_VAN!G$54="Tốt",DAI_VAN!G$35&" ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&") ","")&IF(DAI
    #   _VAN!H$54="Tốt",DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI_VAN!H$36+9)&") ","")&IF(DAI_VAN!I$54="Tốt",DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)
    #   &") ","")&IF(DAI_VAN!J$54="Tốt",DAI_VAN!J$35&" ("&DAI_VAN!J$36&"–"&(DAI_VAN!J$36+9)&") ","")),TRIM(IF(DAI_VAN!C$69="Tốt",DAI_VAN!C$35&" ("&DAI_VAN!C$3
    #   6&"–"&(DAI_VAN!C$36+9)&") ","")&IF(DAI_VAN!D$69="Tốt",DAI_VAN!D$35&" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&") ","")&IF(DAI_VAN!E$69="Tốt",DAI_VAN!E$35&
    #   " ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&") ","")&IF(DAI_VAN!F$69="Tốt",DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&") ","")&IF(DAI_VAN!G$69="Tố
    #   t",DAI_VAN!G$35&" ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&") ","")&IF(DAI_VAN!H$69="Tốt",DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI_VAN!H$36+9)&") ","")&IF(
    #   DAI_VAN!I$69="Tốt",DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&") ","")&IF(DAI_VAN!J$69="Tốt",DAI_VAN!J$35&" ("&DAI_VAN!J$36&"–"&(DAI_VAN!J$36
    #   +9)&") ",""))))&". Kết luận chính vẫn theo "&D33&".")
    return ('' if X.truthy(X.ne(E.v(7,r-17,c),'Thấp')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tham khảo thêm – nếu đọc theo ',('PP1 (xét lệnh trước)' if X.truthy(X.eq(E.v(7,r-19,c),2.0)) else 'PP2 (tính lực)')),': thân '),X.UPPER((E.v(7,r-21,c) if X.truthy(X.eq(E.v(7,r-19,c),2.0)) else E.v(7,r-20,c)))),', dụng thần '),(E.v(8,106,3) if X.truthy(X.eq(E.v(7,r-19,c),2.0)) else E.v(8,106,4))),', hỷ thần '),(E.v(8,107,3) if X.truthy(X.eq(E.v(7,r-19,c),2.0)) else E.v(8,107,4))),'; các đại vận thuận theo cách đọc này: '),('không có' if X.truthy(X.eq((X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c),'Tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+1),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+2),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+3),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+4),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+5),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+6),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+7),'Tốt')) else ''))) if X.truthy(X.eq(E.v(7,r-19,c),2.0)) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c),'Tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+1),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+2),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+3),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+4),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+5),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+6),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+7),'Tốt')) else '')))),'')) else (X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c),'Tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+1),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+2),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+3),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+4),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+5),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+6),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,54,c+7),'Tốt')) else ''))) if X.truthy(X.eq(E.v(7,r-19,c),2.0)) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c),'Tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+1),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+2),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+3),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+4),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+5),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+6),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,69,c+7),'Tốt')) else '')))))),'. Kết luận chính vẫn theo '),E.v(7,r-19,c+1)),'.'))

CELLS = {
    'C6': (f0, 6, 3),
    'D6': (f1, 6, 4),
    'E6': (f2, 6, 5),
    'F6': (f3, 6, 6),
    'C7': (f4, 7, 3),
    'D7': (f5, 7, 4),
    'E7': (f6, 7, 5),
    'F7': (f3, 7, 6),
    'C8': (f7, 8, 3),
    'D8': (f8, 8, 4),
    'E8': (f9, 8, 5),
    'F8': (f3, 8, 6),
    'C9': (f10, 9, 3),
    'D9': (f11, 9, 4),
    'E9': (f12, 9, 5),
    'F9': (f13, 9, 6),
    'C10': (f14, 10, 3),
    'D10': (f15, 10, 4),
    'E10': (f16, 10, 5),
    'F10': (f13, 10, 6),
    'C11': (f17, 11, 3),
    'D11': (f18, 11, 4),
    'E11': (f19, 11, 5),
    'F11': (f13, 11, 6),
    'C12': (f20, 12, 3),
    'D12': (f21, 12, 4),
    'E12': (f22, 12, 5),
    'F12': (f13, 12, 6),
    'C13': (f23, 13, 3),
    'D13': (f24, 13, 4),
    'F13': (f25, 13, 6),
    'C14': (f26, 14, 3),
    'D14': (f27, 14, 4),
    'F14': (f28, 14, 6),
    'C15': (f29, 15, 3),
    'D15': (f30, 15, 4),
    'E15': (f31, 15, 5),
    'F15': (f32, 15, 6),
    'C16': (f33, 16, 3),
    'D16': (f34, 16, 4),
    'E16': (f35, 16, 5),
    'F16': (f36, 16, 6),
    'C17': (f37, 17, 3),
    'D17': (f38, 17, 4),
    'E17': (f39, 17, 5),
    'F17': (f40, 17, 6),
    'C18': (f41, 18, 3),
    'D18': (f42, 18, 4),
    'E18': (f43, 18, 5),
    'F18': (f40, 18, 6),
    'C19': (f44, 19, 3),
    'D19': (f45, 19, 4),
    'E19': (f46, 19, 5),
    'F19': (f40, 19, 6),
    'C20': (f47, 20, 3),
    'D20': (f48, 20, 4),
    'E20': (f49, 20, 5),
    'F20': (f40, 20, 6),
    'C21': (f50, 21, 3),
    'D21': (f51, 21, 4),
    'E21': (f52, 21, 5),
    'F21': (f53, 21, 6),
    'C22': (f54, 22, 3),
    'D22': (f55, 22, 4),
    'E22': (f56, 22, 5),
    'F22': (f57, 22, 6),
    'C23': (f58, 23, 3),
    'D23': (f59, 23, 4),
    'E23': (f60, 23, 5),
    'F23': (f61, 23, 6),
    'C24': (f62, 24, 3),
    'D24': (f63, 24, 4),
    'E24': (f64, 24, 5),
    'F24': (f65, 24, 6),
    'C25': (f66, 25, 3),
    'D25': (f67, 25, 4),
    'E25': (f68, 25, 5),
    'F25': (f69, 25, 6),
    'C31': (f70, 31, 3),
    'D31': (f71, 31, 4),
    'E31': (f72, 31, 5),
    'F31': (f73, 31, 6),
    'C32': (f74, 32, 3),
    'D32': (f75, 32, 4),
    'E32': (f76, 32, 5),
    'F32': (f77, 32, 6),
    'C33': (f78, 33, 3),
    'D33': (f79, 33, 4),
    'F33': (f80, 33, 6),
    'C34': (f81, 34, 3),
    'C35': (f82, 35, 3),
    'F35': (f83, 35, 6),
    'C36': (f84, 36, 3),
    'C37': (f85, 37, 3),
    'C42': (f86, 42, 3),
    'C43': (f87, 43, 3),
    'C44': (f88, 44, 3),
    'C45': (f89, 45, 3),
    'C46': (f90, 46, 3),
    'C47': (f91, 47, 3),
    'C48': (f92, 48, 3),
    'C49': (f93, 49, 3),
    'C50': (f94, 50, 3),
    'C51': (f95, 51, 3),
    'C52': (f96, 52, 3),
}
