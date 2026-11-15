# -*- coding: utf-8 -*-
"""Module sheet TO_CHAT – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.6.xlsx (SHA-256 ef3858360c0984a0528d265328dcd5caaf8d951a2feed88333eec94b8b316638).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'TO_CHAT'
SI = 20

def f0(E, r, c):
    # Ô C5
    # = INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+1)&" – "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$6,DUNG_HY_KY!$C$6)&" (theo "&DOI_CHIE
    #   U!$D$33&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),1.0)),' – '),(E.v(8,6,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,6,3))),' (theo '),E.v(7,33,4)),')')

def f1(E, r, c):
    # Ô H5
    # = "Mộc"
    return 'Mộc'

def f2(E, r, c):
    # Ô I5
    # = "Hỏa"
    return 'Hỏa'

def f3(E, r, c):
    # Ô J5
    # = "Thổ"
    return 'Thổ'

def f4(E, r, c):
    # Ô K5
    # = "Kim"
    return 'Kim'

def f5(E, r, c):
    # Ô L5
    # = "Thủy"
    return 'Thủy'

def f6(E, r, c):
    # Ô H6
    # = IF(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5),-1,IF(DOI_CHIEU!$C$33=2,INDEX(DUNG_HY_KY!$D$24:$D$28,1),INDEX(DUNG_HY_KY!$C$24:$C$28,1)))
    return (X.neg(1.0) if X.truthy(X.eq(0.0,(E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)))) else (X.INDEX(E.rg(8,24,4,28,4), 1.0) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(8,24,3,28,3), 1.0)))

def f7(E, r, c):
    # Ô I6
    # = IF(1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5),-1,IF(DOI_CHIEU!$C$33=2,INDEX(DUNG_HY_KY!$D$24:$D$28,2),INDEX(DUNG_HY_KY!$C$24:$C$28,2)))
    return (X.neg(1.0) if X.truthy(X.eq(1.0,(E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)))) else (X.INDEX(E.rg(8,24,4,28,4), 2.0) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(8,24,3,28,3), 2.0)))

def f8(E, r, c):
    # Ô J6
    # = IF(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5),-1,IF(DOI_CHIEU!$C$33=2,INDEX(DUNG_HY_KY!$D$24:$D$28,3),INDEX(DUNG_HY_KY!$C$24:$C$28,3)))
    return (X.neg(1.0) if X.truthy(X.eq(2.0,(E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)))) else (X.INDEX(E.rg(8,24,4,28,4), 3.0) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(8,24,3,28,3), 3.0)))

def f9(E, r, c):
    # Ô K6
    # = IF(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5),-1,IF(DOI_CHIEU!$C$33=2,INDEX(DUNG_HY_KY!$D$24:$D$28,4),INDEX(DUNG_HY_KY!$C$24:$C$28,4)))
    return (X.neg(1.0) if X.truthy(X.eq(3.0,(E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)))) else (X.INDEX(E.rg(8,24,4,28,4), 4.0) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(8,24,3,28,3), 4.0)))

def f10(E, r, c):
    # Ô L6
    # = IF(4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5),-1,IF(DOI_CHIEU!$C$33=2,INDEX(DUNG_HY_KY!$D$24:$D$28,5),INDEX(DUNG_HY_KY!$C$24:$C$28,5)))
    return (X.neg(1.0) if X.truthy(X.eq(4.0,(E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)))) else (X.INDEX(E.rg(8,24,4,28,4), 5.0) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(8,24,3,28,3), 5.0)))

def f11(E, r, c):
    # Ô C7
    # = IF(DUNG_HY_KY!$C$72<>"",-1,IF(DOI_CHIEU!$C$33=2,IF(DUNG_HY_KY!$D$6="Độc vượng",0,IF(DUNG_HY_KY!$D$6="Nhật chủ vượng",1,3+(MATCH(MAX($H$6:$L$6),$H$6:$L
    #   $6,0)-1))),IF(DUNG_HY_KY!$C$6="Quá vượng",0,IF(OR(DUNG_HY_KY!$C$6="Rất vượng",DUNG_HY_KY!$C$6="Thân vượng"),1,IF(DUNG_HY_KY!$C$6="Trung hòa",2,IF(DUNG
    #   _HY_KY!$C$6="Quá nhược",8,3+(MATCH(MAX($H$6:$L$6),$H$6:$L$6,0)-1)))))))
    return (X.neg(1.0) if X.truthy(X.ne(E.v(8,72,3),'')) else ((0.0 if X.truthy(X.eq(E.v(8,6,4),'Độc vượng')) else (1.0 if X.truthy(X.eq(E.v(8,6,4),'Nhật chủ vượng')) else X.add(3.0,X.sub(X.MATCH(X.MAX(E.rg(20,6,8,6,12)), E.rg(20,6,8,6,12), 0.0),1.0)))) if X.truthy(X.eq(E.v(7,33,3),2.0)) else (0.0 if X.truthy(X.eq(E.v(8,6,3),'Quá vượng')) else (1.0 if X.truthy(X.OR(X.eq(E.v(8,6,3),'Rất vượng'), X.eq(E.v(8,6,3),'Thân vượng'))) else (2.0 if X.truthy(X.eq(E.v(8,6,3),'Trung hòa')) else (8.0 if X.truthy(X.eq(E.v(8,6,3),'Quá nhược')) else X.add(3.0,X.sub(X.MATCH(X.MAX(E.rg(20,6,8,6,12)), E.rg(20,6,8,6,12), 0.0),1.0))))))))

def f12(E, r, c):
    # Ô D7
    # = IF(C7<0,"cách đặc biệt – xem tính cách ngoại cách",CHOOSE(C7+1,"vượng quá","vượng vừa","vừa phải","nhược – Mộc nhiều","nhược – Hỏa nhiều","nhược – Thổ
    #    nhiều","nhược – Kim nhiều","nhược – Thủy nhiều","nhược quá"))
    return ('cách đặc biệt – xem tính cách ngoại cách' if X.truthy(X.lt(E.v(20,r,c-1),0.0)) else X.CHOOSE(X.add(E.v(20,r,c-1),1.0), lambda: 'vượng quá', lambda: 'vượng vừa', lambda: 'vừa phải', lambda: 'nhược – Mộc nhiều', lambda: 'nhược – Hỏa nhiều', lambda: 'nhược – Thổ nhiều', lambda: 'nhược – Kim nhiều', lambda: 'nhược – Thủy nhiều', lambda: 'nhược quá'))

def f13(E, r, c):
    # Ô C8
    # = IF(C7<0,"Lá số thuộc cách đặc biệt "&DUNG_HY_KY!$C$72&" – sách xem cá tính theo ngoại cách (mục D).","Nhật chủ "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,D
    #   UNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+1)&" "&D7&": "&IF(C7<0,"",INDEX($N$4:$V$8,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+1,C7+1))&IF(IF(C7<0,""
    #   ,INDEX($N$4:$V$8,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+1,C7+1))="","(sách không có dòng này)","")&" ("&INDEX($W$4:$W$8,IF(DOI_CHIEU!$C
    #   $33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+1)&").")
    return (X.cat(X.cat('Lá số thuộc cách đặc biệt ',E.v(8,72,3)),' – sách xem cá tính theo ngoại cách (mục D).') if X.truthy(X.lt(E.v(20,r-1,c),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),1.0))),' '),E.v(20,r-1,c+1)),': '),('' if X.truthy(X.lt(E.v(20,r-1,c),0.0)) else X.INDEX(E.rg(20,4,14,8,22), X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),1.0), X.add(E.v(20,r-1,c),1.0)))),('(sách không có dòng này)' if X.truthy(X.eq(('' if X.truthy(X.lt(E.v(20,r-1,c),0.0)) else X.INDEX(E.rg(20,4,14,8,22), X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),1.0), X.add(E.v(20,r-1,c),1.0))),'')) else '')),' ('),X.INDEX(E.rg(20,4,23,8,23), X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),1.0))),').'))

def f14(E, r, c):
    # Ô C9
    # = "Dữ kiện: hành nhật chủ "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+1)&", trạng thái "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$6
    #   ,DUNG_HY_KY!$C$6)&"; lượng các hành khác (PP khuyến nghị): "&TRIM(IF(H6<0,"","Mộc "&TEXT(H6,"0.0")&"; ")&IF(I6<0,"","Hỏa "&TEXT(I6,"0.0")&"; ")&IF(J6<
    #   0,"","Thổ "&TEXT(J6,"0.0")&"; ")&IF(K6<0,"","Kim "&TEXT(K6,"0.0")&"; ")&IF(L6<0,"","Thủy "&TEXT(L6,"0.0")&"; "))&". Quy tắc: P2 tr.578-581 chia mỗi hà
    #   nh nhật chủ thành vượng quá / vượng vừa / vừa phải / nhược (theo hành khác nhiều nhất) / nhược quá; ánh xạ: Quá vượng (PP2: độc vượng) → vượng quá; Rấ
    #   t vượng, Thân vượng (PP2: vượng) → vượng vừa; Trung hò"&"a → vừa phải; Thân nhược → nhược + hành nhiều nhất; Quá nhược → nhược quá (giả định G13). Kết
    #    luận: "&D7&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: hành nhật chủ ',X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),1.0))),', trạng thái '),(E.v(8,6,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,6,3))),'; lượng các hành khác (PP khuyến nghị): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(('' if X.truthy(X.lt(E.v(20,r-3,c+5),0.0)) else X.cat(X.cat('Mộc ',X.TEXT(E.v(20,r-3,c+5), '0.0')),'; ')),('' if X.truthy(X.lt(E.v(20,r-3,c+6),0.0)) else X.cat(X.cat('Hỏa ',X.TEXT(E.v(20,r-3,c+6), '0.0')),'; '))),('' if X.truthy(X.lt(E.v(20,r-3,c+7),0.0)) else X.cat(X.cat('Thổ ',X.TEXT(E.v(20,r-3,c+7), '0.0')),'; '))),('' if X.truthy(X.lt(E.v(20,r-3,c+8),0.0)) else X.cat(X.cat('Kim ',X.TEXT(E.v(20,r-3,c+8), '0.0')),'; '))),('' if X.truthy(X.lt(E.v(20,r-3,c+9),0.0)) else X.cat(X.cat('Thủy ',X.TEXT(E.v(20,r-3,c+9), '0.0')),'; '))))),'. Quy tắc: P2 tr.578-581 chia mỗi hành nhật chủ thành vượng quá / vượng vừa / vừa phải / nhược (theo hành khác nhiều nhất) / nhược quá; ánh xạ: Quá vượng (PP2: độc vượng) → vượng quá; Rất vượng, Thân vượng (PP2: vượng) → vượng vừa; Trung hò'),'a → vừa phải; Thân nhược → nhược + hành nhiều nhất; Quá nhược → nhược quá (giả định G13). Kết luận: '),E.v(20,r-2,c+1)),'.')

def f15(E, r, c):
    # Ô C12
    # = TRIM(IF(TO_CHAT!$L$42="ĐỦ","• "&TO_CHAT!$G$42&" (P1 tr.140) ","")&IF(TO_CHAT!$L$43="ĐỦ","• "&TO_CHAT!$G$43&" (P1 tr.139) ","")&IF(TO_CHAT!$L$44="ĐỦ","
    #   • "&TO_CHAT!$G$44&" (P1 tr.88; P2 tr.575) ","")&IF(TO_CHAT!$L$45="ĐỦ","• "&TO_CHAT!$G$45&" (P2 tr.577) ","")&IF(TO_CHAT!$L$47="ĐỦ","• "&TO_CHAT!$G$47&
    #   " (P2 tr.573) ","")&IF(TO_CHAT!$L$48="ĐỦ","• "&TO_CHAT!$G$48&" (P2 tr.574) ","")&IF(TO_CHAT!$L$49="ĐỦ","• "&TO_CHAT!$G$49&" (P2 tr.574) ","")&IF(TO_CH
    #   AT!$L$50="ĐỦ","• "&TO_CHAT!$G$50&" (P2 tr.574) ","")&IF(TO_CHAT!$L$51="ĐỦ","• "&TO_CHAT!$G$51&" (P2 tr.574) ","")&IF(TO_CHAT!$L$52="ĐỦ","• "&TO_CHAT!$
    #   G$52&" (P2 tr.301, 575) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(20,42,7)),' (P1 tr.140) ') if X.truthy(X.eq(E.v(20,42,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,43,7)),' (P1 tr.139) ') if X.truthy(X.eq(E.v(20,43,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,44,7)),' (P1 tr.88; P2 tr.575) ') if X.truthy(X.eq(E.v(20,44,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,45,7)),' (P2 tr.577) ') if X.truthy(X.eq(E.v(20,45,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,47,7)),' (P2 tr.573) ') if X.truthy(X.eq(E.v(20,47,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,48,7)),' (P2 tr.574) ') if X.truthy(X.eq(E.v(20,48,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,49,7)),' (P2 tr.574) ') if X.truthy(X.eq(E.v(20,49,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,50,7)),' (P2 tr.574) ') if X.truthy(X.eq(E.v(20,50,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,51,7)),' (P2 tr.574) ') if X.truthy(X.eq(E.v(20,51,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,52,7)),' (P2 tr.301, 575) ') if X.truthy(X.eq(E.v(20,52,12),'ĐỦ')) else '')))

def f16(E, r, c):
    # Ô C13
    # = TRIM(IF(TO_CHAT!$L$46="ĐỦ","• "&TO_CHAT!$G$46&" (P2 tr.577) ","")&IF(TO_CHAT!$L$53="ĐỦ","• "&TO_CHAT!$G$53&" (P2 tr.577) ","")&IF(TO_CHAT!$L$54="ĐỦ","
    #   • "&TO_CHAT!$G$54&" (P2 tr.577) ",""))
    return X.TRIM(X.cat(X.cat((X.cat(X.cat('• ',E.v(20,46,7)),' (P2 tr.577) ') if X.truthy(X.eq(E.v(20,46,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,53,7)),' (P2 tr.577) ') if X.truthy(X.eq(E.v(20,53,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,54,7)),' (P2 tr.577) ') if X.truthy(X.eq(E.v(20,54,12),'ĐỦ')) else '')))

def f17(E, r, c):
    # Ô C16
    # = TRIM(IF(AND(TO_CHAT!$L$55="ĐỦ",TO_CHAT!$I$55<>"Bất lợi"),"• "&TO_CHAT!$G$55&" ("&TO_CHAT!$H$55&") ","")&IF(AND(TO_CHAT!$L$56="ĐỦ",TO_CHAT!$I$56<>"Bất 
    #   lợi"),"• "&TO_CHAT!$G$56&" ("&TO_CHAT!$H$56&") ","")&IF(AND(TO_CHAT!$L$57="ĐỦ",TO_CHAT!$I$57<>"Bất lợi"),"• "&TO_CHAT!$G$57&" ("&TO_CHAT!$H$57&") ",""
    #   )&IF(AND(TO_CHAT!$L$58="ĐỦ",TO_CHAT!$I$58<>"Bất lợi"),"• "&TO_CHAT!$G$58&" ("&TO_CHAT!$H$58&") ","")&IF(AND(TO_CHAT!$L$59="ĐỦ",TO_CHAT!$I$59<>"Bất lợi
    #   "),"• "&TO_CHAT!$G$59&" ("&TO_CHAT!$H$59&") ","")&IF(AND(TO_CHAT!$L$60="ĐỦ",TO_CHAT!$I$60<>"Bất lợi"),"• "&TO_CHAT!$G$60&" ("&TO_CHAT!$H$60&") ","")&I
    #   F(AND(TO_CHAT!$L$61="ĐỦ",TO_CHAT!$I$61<>"Bất lợi"),"• "&TO_CHAT!$G$61&" ("&TO_CHAT!$H$61&") ","")&IF(AND(TO_CHAT!$L$62="ĐỦ",TO_CHAT!$I$62<>"Bất lợi"),
    #   "• "&TO_CHAT!$G$62&" ("&TO_CHAT!$H$62&") ","")&IF(AND(TO_CHAT!$L$63="ĐỦ",TO_CHAT!$I$63<>"Bất lợi"),"• "&TO_CHAT!$G$63&" ("&TO_CHAT!$H$63&") ","")&IF(T
    #   O_CHAT!$M$16="","",IF(TRIM(IF(AND(TO_CHAT!$L$55="ĐỦ",TO_CHAT!$I$55<>"Bất lợi"),"• "&TO_CHAT!$G$55&" ("&TO_CHAT!$H$55&") ","")&IF(AND(TO_CHAT!$L$56="ĐỦ
    #   ",TO_CHAT!$I$56<>"Bất lợi"),"• "&TO_CHAT!$G$56&" ("&TO_CHAT!$H$56&") ","")&IF(AND(TO_CHAT!$L$57="ĐỦ",TO_CHAT!$I$57<>"Bất lợi"),"• "&TO_CHAT!$G$57&" ("
    #   &TO_CHAT!$H$57&") ","")&IF(AND(TO_CHAT!$L$58="ĐỦ",TO_CHAT!$I$58<>"Bất lợi"),"• "&TO_CHAT!$G$58&" ("&TO_CHAT!$H$58&") ","")&IF(AND(TO_CHAT!$L$59="ĐỦ",T
    #   O_CHAT!$I$59<>"Bất lợi"),"• "&TO_CHAT!$G$59&" ("&TO_CHAT!$H$59&") ","")&IF(AND(TO_CHAT!$L$60="ĐỦ",TO_CHAT!$I$60<>"Bất lợi"),"• "&TO_CHAT!$G$60&" ("&TO
    #   _CHAT!$H$60&") ","")&IF(AND(TO_CHAT!$L$61="ĐỦ",TO_CHAT!$I$61<>"Bất lợi"),"• "&TO_CHAT!$G$61&" ("&TO_CHAT!$H$61&") ","")&IF(AND(TO_CHAT!$L$62="ĐỦ",TO_C
    #   HAT!$I$62<>"Bất lợi"),"• "&TO_CHAT!$G$62&" ("&TO_CHAT!$H$62&") ","")&IF(AND(TO_CHAT!$L$63="ĐỦ",TO_CHAT!$I$63<>"Bất lợi"),"• "&TO_CHAT!$G$63&" ("&TO_CH
    #   AT!$H$63&") ",""))="",""," ")&"⚠ LƯU Ý – sách nêu điểm bất lợi: "&TO_CHAT!$M$16))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('• ',E.v(20,55,7)),' ('),E.v(20,55,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,55,12),'ĐỦ'), X.ne(E.v(20,55,9),'Bất lợi'))) else ''),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,56,7)),' ('),E.v(20,56,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,56,12),'ĐỦ'), X.ne(E.v(20,56,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,57,7)),' ('),E.v(20,57,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,57,12),'ĐỦ'), X.ne(E.v(20,57,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,58,7)),' ('),E.v(20,58,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,58,12),'ĐỦ'), X.ne(E.v(20,58,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,59,7)),' ('),E.v(20,59,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,59,12),'ĐỦ'), X.ne(E.v(20,59,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,60,7)),' ('),E.v(20,60,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,60,12),'ĐỦ'), X.ne(E.v(20,60,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,61,7)),' ('),E.v(20,61,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,61,12),'ĐỦ'), X.ne(E.v(20,61,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,62,7)),' ('),E.v(20,62,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,62,12),'ĐỦ'), X.ne(E.v(20,62,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,63,7)),' ('),E.v(20,63,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,63,12),'ĐỦ'), X.ne(E.v(20,63,9),'Bất lợi'))) else '')),('' if X.truthy(X.eq(E.v(20,16,13),'')) else X.cat(X.cat(('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('• ',E.v(20,55,7)),' ('),E.v(20,55,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,55,12),'ĐỦ'), X.ne(E.v(20,55,9),'Bất lợi'))) else ''),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,56,7)),' ('),E.v(20,56,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,56,12),'ĐỦ'), X.ne(E.v(20,56,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,57,7)),' ('),E.v(20,57,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,57,12),'ĐỦ'), X.ne(E.v(20,57,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,58,7)),' ('),E.v(20,58,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,58,12),'ĐỦ'), X.ne(E.v(20,58,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,59,7)),' ('),E.v(20,59,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,59,12),'ĐỦ'), X.ne(E.v(20,59,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,60,7)),' ('),E.v(20,60,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,60,12),'ĐỦ'), X.ne(E.v(20,60,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,61,7)),' ('),E.v(20,61,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,61,12),'ĐỦ'), X.ne(E.v(20,61,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,62,7)),' ('),E.v(20,62,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,62,12),'ĐỦ'), X.ne(E.v(20,62,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,63,7)),' ('),E.v(20,63,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,63,12),'ĐỦ'), X.ne(E.v(20,63,9),'Bất lợi'))) else ''))),'')) else ' '),'⚠ LƯU Ý – sách nêu điểm bất lợi: '),E.v(20,16,13)))))

def f18(E, r, c):
    # Ô M16
    # = TRIM(IF(AND(TO_CHAT!$L$55="ĐỦ",TO_CHAT!$I$55="Bất lợi"),"• "&TO_CHAT!$G$55&" ("&TO_CHAT!$H$55&") ","")&IF(AND(TO_CHAT!$L$56="ĐỦ",TO_CHAT!$I$56="Bất lợ
    #   i"),"• "&TO_CHAT!$G$56&" ("&TO_CHAT!$H$56&") ","")&IF(AND(TO_CHAT!$L$57="ĐỦ",TO_CHAT!$I$57="Bất lợi"),"• "&TO_CHAT!$G$57&" ("&TO_CHAT!$H$57&") ","")&I
    #   F(AND(TO_CHAT!$L$58="ĐỦ",TO_CHAT!$I$58="Bất lợi"),"• "&TO_CHAT!$G$58&" ("&TO_CHAT!$H$58&") ","")&IF(AND(TO_CHAT!$L$59="ĐỦ",TO_CHAT!$I$59="Bất lợi"),"•
    #    "&TO_CHAT!$G$59&" ("&TO_CHAT!$H$59&") ","")&IF(AND(TO_CHAT!$L$60="ĐỦ",TO_CHAT!$I$60="Bất lợi"),"• "&TO_CHAT!$G$60&" ("&TO_CHAT!$H$60&") ","")&IF(AND(
    #   TO_CHAT!$L$61="ĐỦ",TO_CHAT!$I$61="Bất lợi"),"• "&TO_CHAT!$G$61&" ("&TO_CHAT!$H$61&") ","")&IF(AND(TO_CHAT!$L$62="ĐỦ",TO_CHAT!$I$62="Bất lợi"),"• "&TO_
    #   CHAT!$G$62&" ("&TO_CHAT!$H$62&") ","")&IF(AND(TO_CHAT!$L$63="ĐỦ",TO_CHAT!$I$63="Bất lợi"),"• "&TO_CHAT!$G$63&" ("&TO_CHAT!$H$63&") ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('• ',E.v(20,55,7)),' ('),E.v(20,55,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,55,12),'ĐỦ'), X.eq(E.v(20,55,9),'Bất lợi'))) else ''),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,56,7)),' ('),E.v(20,56,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,56,12),'ĐỦ'), X.eq(E.v(20,56,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,57,7)),' ('),E.v(20,57,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,57,12),'ĐỦ'), X.eq(E.v(20,57,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,58,7)),' ('),E.v(20,58,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,58,12),'ĐỦ'), X.eq(E.v(20,58,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,59,7)),' ('),E.v(20,59,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,59,12),'ĐỦ'), X.eq(E.v(20,59,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,60,7)),' ('),E.v(20,60,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,60,12),'ĐỦ'), X.eq(E.v(20,60,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,61,7)),' ('),E.v(20,61,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,61,12),'ĐỦ'), X.eq(E.v(20,61,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,62,7)),' ('),E.v(20,62,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,62,12),'ĐỦ'), X.eq(E.v(20,62,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,63,7)),' ('),E.v(20,63,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,63,12),'ĐỦ'), X.eq(E.v(20,63,9),'Bất lợi'))) else '')))

def f19(E, r, c):
    # Ô C17
    # = IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)<0,"Không xác định được dụng thần",INDEX($X$4:$X$8,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,
    #   DUNG_HY_KY!$C$109)+1))
    return ('Không xác định được dụng thần' if X.truthy(X.lt((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0)) else X.INDEX(E.rg(20,4,24,8,24), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0)))

def f20(E, r, c):
    # Ô C18
    # = IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)<0,"—",INDEX($X$4:$X$8,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1))
    return ('—' if X.truthy(X.lt((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else X.INDEX(E.rg(20,4,24,8,24), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0)))

def f21(E, r, c):
    # Ô C19
    # = "Sách: chọn nghề nghiệp là một cách ""bổ"" để cân bằng tổng hợp cho Tứ trụ (P1 tr.24); bảng ngũ hành hợp với ngành nghề và phương vị (P1 tr.26). Hành 
    #   bổ cho tứ trụ là dụng thần, sau đó hỷ thần (P2 tr.380-382) – theo phương pháp được khuyến nghị: "&DOI_CHIEU!$D$33&"."
    return X.cat(X.cat('Sách: chọn nghề nghiệp là một cách "bổ" để cân bằng tổng hợp cho Tứ trụ (P1 tr.24); bảng ngũ hành hợp với ngành nghề và phương vị (P1 tr.26). Hành bổ cho tứ trụ là dụng thần, sau đó hỷ thần (P2 tr.380-382) – theo phương pháp được khuyến nghị: ',E.v(7,33,4)),'.')

def f22(E, r, c):
    # Ô C22
    # = TRIM(IF(TO_CHAT!$L$64="ĐỦ","• "&TO_CHAT!$G$64&" (P2 tr.582) ","")&IF(TO_CHAT!$L$65="ĐỦ","• "&TO_CHAT!$G$65&" (P2 tr.582) ","")&IF(TO_CHAT!$L$66="ĐỦ","
    #   • "&TO_CHAT!$G$66&" (P2 tr.582) ","")&IF(TO_CHAT!$L$67="ĐỦ","• "&TO_CHAT!$G$67&" (P2 tr.582) ",""))
    return X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(20,64,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,64,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,65,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,65,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,66,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,66,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,67,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,67,12),'ĐỦ')) else '')))

def f23(E, r, c):
    # Ô C23
    # = TRIM(IF(TO_CHAT!$L$68="ĐỦ","• "&TO_CHAT!$G$68&" (P2 tr.581) ","")&IF(TO_CHAT!$L$69="ĐỦ","• "&TO_CHAT!$G$69&" (P2 tr.581) ","")&IF(TO_CHAT!$L$70="ĐỦ","
    #   • "&TO_CHAT!$G$70&" (P2 tr.581) ","")&IF(TO_CHAT!$L$71="ĐỦ","• "&TO_CHAT!$G$71&" (P2 tr.581) ","")&IF(TO_CHAT!$L$72="ĐỦ","• "&TO_CHAT!$G$72&" (P2 tr.5
    #   81) ","")&IF(TO_CHAT!$L$73="ĐỦ","• "&TO_CHAT!$G$73&" (P2 tr.581) ","")&IF(TO_CHAT!$L$74="ĐỦ","• "&TO_CHAT!$G$74&" (P2 tr.581) ","")&IF(TO_CHAT!$L$75="
    #   ĐỦ","• "&TO_CHAT!$G$75&" (P2 tr.581) ","")&IF(TO_CHAT!$L$76="ĐỦ","• "&TO_CHAT!$G$76&" (P2 tr.581) ","")&IF(TO_CHAT!$L$77="ĐỦ","• "&TO_CHAT!$G$77&" (P2
    #    tr.581) ","")&IF(TO_CHAT!$L$78="ĐỦ","• "&TO_CHAT!$G$78&" (P2 tr.581) ","")&IF(TO_CHAT!$L$79="ĐỦ","• "&TO_CHAT!$G$79&" (P2 tr.581) ","")&IF(TO_CHAT!$L
    #   $80="ĐỦ","• "&TO_CHAT!$G$80&" (P2 tr.581) ","")&IF(TO_CHAT!$L$81="ĐỦ","• "&TO_CHAT!$G$81&" (P2 tr.581) ","")&IF(TO_CHAT!$L$82="ĐỦ","• "&TO_CHAT!$G$82&
    #   " (P2 tr.581) ","")&IF(TO_CHAT!$L$83="ĐỦ","• "&TO_CHAT!$G$83&" (P2 tr.581) ","")&IF(TO_CHAT!$L$84="ĐỦ","• "&TO_CHAT!$G$84&" (P2 tr.581) ","")&IF(TO_CH
    #   AT!$L$85="ĐỦ","• "&TO_CHAT!$G$85&" (P2 tr.581) ","")&IF(TO_CHAT!$L$86="ĐỦ","• "&TO_CHAT!$G$86&" (P2 tr.581) ","")&IF(TO_CHAT!$L$87="ĐỦ","• "&TO_CHAT!$
    #   G$87&" (P2 tr.581) ","")&IF(TO_CHAT!$L$88="ĐỦ","• "&TO_CHAT!$G$88&" (P2 tr.582) ","")&IF(TO_CHAT!$L$89="ĐỦ","• "&TO_CHAT!$G$89&" (P2 tr.582) ","")&IF(
    #   TO_CHAT!$L$90="ĐỦ","• "&TO_CHAT!$G$90&" (P2 tr.582) ","")&IF(TO_CHAT!$L$91="ĐỦ","• "&TO_CHAT!$G$91&" (P2 tr.582) ","")&IF(TO_CHAT!$L$92="ĐỦ","• "&TO_C
    #   HAT!$G$92&" (P2 tr.582) ","")&IF(TO_CHAT!$L$93="ĐỦ","• "&TO_CHAT!$G$93&" (P2 tr.582) ","")&IF(TO_CHAT!$L$94="ĐỦ","• "&TO_CHAT!$G$94&" (P2 tr.582) ",""
    #   )&IF(TO_CHAT!$L$95="ĐỦ","• "&TO_CHAT!$G$95&" (P2 tr.582) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(20,68,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,68,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,69,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,69,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,70,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,70,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,71,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,71,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,72,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,72,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,73,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,73,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,74,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,74,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,75,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,75,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,76,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,76,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,77,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,77,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,78,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,78,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,79,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,79,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,80,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,80,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,81,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,81,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,82,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,82,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,83,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,83,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,84,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,84,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,85,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,85,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,86,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,86,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,87,7)),' (P2 tr.581) ') if X.truthy(X.eq(E.v(20,87,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,88,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,88,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,89,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,89,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,90,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,90,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,91,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,91,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,92,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,92,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,93,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,93,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,94,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,94,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,95,7)),' (P2 tr.582) ') if X.truthy(X.eq(E.v(20,95,12),'ĐỦ')) else '')))

def f24(E, r, c):
    # Ô C24
    # = TRIM(IF(TO_CHAT!$L$96="ĐỦ","• "&TO_CHAT!$G$96&" (P1 tr.87-88) ","")&IF(TO_CHAT!$L$97="ĐỦ","• "&TO_CHAT!$G$97&" (P1 tr.87-88) ","")&IF(TO_CHAT!$L$98="Đ
    #   Ủ","• "&TO_CHAT!$G$98&" (P1 tr.87-88) ","")&IF(TO_CHAT!$L$99="ĐỦ","• "&TO_CHAT!$G$99&" (P1 tr.87-88) ","")&IF(TO_CHAT!$L$100="ĐỦ","• "&TO_CHAT!$G$100&
    #   " (P1 tr.87-88) ","")&IF(TO_CHAT!$L$101="ĐỦ","• "&TO_CHAT!$G$101&" (P1 tr.87-88) ","")&IF(TO_CHAT!$L$102="ĐỦ","• "&TO_CHAT!$G$102&" (P1 tr.87-88) ",""
    #   )&IF(TO_CHAT!$L$103="ĐỦ","• "&TO_CHAT!$G$103&" (P1 tr.87-88) ","")&IF(TO_CHAT!$L$104="ĐỦ","• "&TO_CHAT!$G$104&" (P1 tr.87-88) ","")&IF(TO_CHAT!$L$105=
    #   "ĐỦ","• "&TO_CHAT!$G$105&" (P1 tr.87-88) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(20,96,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,96,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,97,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,97,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,98,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,98,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,99,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,99,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,100,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,100,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,101,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,101,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,102,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,102,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,103,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,103,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,104,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,104,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,105,7)),' (P1 tr.87-88) ') if X.truthy(X.eq(E.v(20,105,12),'ĐỦ')) else '')))

def f25(E, r, c):
    # Ô C25
    # = "Điểm tốt: "&IF(TRIM(AB25&" "&AC25)="","không có câu nào đủ điều kiện",TRIM(AB25&" "&AC25))&CHAR(10)&"Điểm cần lưu ý: "&IF(TRIM(AD25&" "&AE25)="","khô
    #   ng có",TRIM(AD25&" "&AE25))&CHAR(10)&"Đặc điểm khác: "&IF(TRIM(AF25&" "&AG25)="","không có",TRIM(AF25&" "&AG25))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Điểm tốt: ',('không có câu nào đủ điều kiện' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(E.v(20,r,c+25),' '),E.v(20,r,c+26))),'')) else X.TRIM(X.cat(X.cat(E.v(20,r,c+25),' '),E.v(20,r,c+26))))),X.CHAR(10.0)),'Điểm cần lưu ý: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(E.v(20,r,c+27),' '),E.v(20,r,c+28))),'')) else X.TRIM(X.cat(X.cat(E.v(20,r,c+27),' '),E.v(20,r,c+28))))),X.CHAR(10.0)),'Đặc điểm khác: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(E.v(20,r,c+29),' '),E.v(20,r,c+30))),'')) else X.TRIM(X.cat(X.cat(E.v(20,r,c+29),' '),E.v(20,r,c+30)))))

def f26(E, r, c):
    # Ô AB25
    # = TRIM(IF(L110="ĐỦ","• "&G110&" ["&B110&", "&H110&"] ","")&IF(L111="ĐỦ","• "&G111&" ["&B111&", "&H111&"] ","")&IF(L113="ĐỦ","• "&G113&" ["&B113&", "&H11
    #   3&"] ","")&IF(L114="ĐỦ","• "&G114&" ["&B114&", "&H114&"] ","")&IF(L116="ĐỦ","• "&G116&" ["&B116&", "&H116&"] ","")&IF(L117="ĐỦ","• "&G117&" ["&B117&",
    #    "&H117&"] ","")&IF(L118="ĐỦ","• "&G118&" ["&B118&", "&H118&"] ","")&IF(L119="ĐỦ","• "&G119&" ["&B119&", "&H119&"] ","")&IF(L120="ĐỦ","• "&G120&" ["&B
    #   120&", "&H120&"] ","")&IF(L122="ĐỦ","• "&G122&" ["&B122&", "&H122&"] ","")&IF(L130="ĐỦ","• "&G130&" ["&B130&", "&H130&"] ","")&IF(L131="ĐỦ","• "&G131&
    #   " ["&B131&", "&H131&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+85,c-21)),' ['),E.v(20,r+85,c-26)),', '),E.v(20,r+85,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+85,c-16),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+86,c-21)),' ['),E.v(20,r+86,c-26)),', '),E.v(20,r+86,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+86,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+88,c-21)),' ['),E.v(20,r+88,c-26)),', '),E.v(20,r+88,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+88,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+89,c-21)),' ['),E.v(20,r+89,c-26)),', '),E.v(20,r+89,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+89,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+91,c-21)),' ['),E.v(20,r+91,c-26)),', '),E.v(20,r+91,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+91,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+92,c-21)),' ['),E.v(20,r+92,c-26)),', '),E.v(20,r+92,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+92,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+93,c-21)),' ['),E.v(20,r+93,c-26)),', '),E.v(20,r+93,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+93,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+94,c-21)),' ['),E.v(20,r+94,c-26)),', '),E.v(20,r+94,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+94,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+95,c-21)),' ['),E.v(20,r+95,c-26)),', '),E.v(20,r+95,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+95,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+97,c-21)),' ['),E.v(20,r+97,c-26)),', '),E.v(20,r+97,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+97,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+105,c-21)),' ['),E.v(20,r+105,c-26)),', '),E.v(20,r+105,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+105,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+106,c-21)),' ['),E.v(20,r+106,c-26)),', '),E.v(20,r+106,c-20)),'] ') if X.truthy(X.eq(E.v(20,r+106,c-16),'ĐỦ')) else '')))

def f27(E, r, c):
    # Ô AC25
    # = TRIM(IF(L138="ĐỦ","• "&G138&" ["&B138&", "&H138&"] ","")&IF(L139="ĐỦ","• "&G139&" ["&B139&", "&H139&"] ","")&IF(L140="ĐỦ","• "&G140&" ["&B140&", "&H14
    #   0&"] ","")&IF(L141="ĐỦ","• "&G141&" ["&B141&", "&H141&"] ","")&IF(L142="ĐỦ","• "&G142&" ["&B142&", "&H142&"] ","")&IF(L143="ĐỦ","• "&G143&" ["&B143&",
    #    "&H143&"] ","")&IF(L146="ĐỦ","• "&G146&" ["&B146&", "&H146&"] ","")&IF(L148="ĐỦ","• "&G148&" ["&B148&", "&H148&"] ","")&IF(L150="ĐỦ","• "&G150&" ["&B
    #   150&", "&H150&"] ","")&IF(L152="ĐỦ","• "&G152&" ["&B152&", "&H152&"] ","")&IF(L153="ĐỦ","• "&G153&" ["&B153&", "&H153&"] ","")&IF(L154="ĐỦ","• "&G154&
    #   " ["&B154&", "&H154&"] ","")&IF(L178="ĐỦ","• "&G178&" ["&B178&", "&H178&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+113,c-22)),' ['),E.v(20,r+113,c-27)),', '),E.v(20,r+113,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+113,c-17),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+114,c-22)),' ['),E.v(20,r+114,c-27)),', '),E.v(20,r+114,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+114,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+115,c-22)),' ['),E.v(20,r+115,c-27)),', '),E.v(20,r+115,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+115,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+116,c-22)),' ['),E.v(20,r+116,c-27)),', '),E.v(20,r+116,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+116,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+117,c-22)),' ['),E.v(20,r+117,c-27)),', '),E.v(20,r+117,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+117,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+118,c-22)),' ['),E.v(20,r+118,c-27)),', '),E.v(20,r+118,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+118,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+121,c-22)),' ['),E.v(20,r+121,c-27)),', '),E.v(20,r+121,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+121,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+123,c-22)),' ['),E.v(20,r+123,c-27)),', '),E.v(20,r+123,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+123,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+125,c-22)),' ['),E.v(20,r+125,c-27)),', '),E.v(20,r+125,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+125,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+127,c-22)),' ['),E.v(20,r+127,c-27)),', '),E.v(20,r+127,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+127,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+128,c-22)),' ['),E.v(20,r+128,c-27)),', '),E.v(20,r+128,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+128,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+129,c-22)),' ['),E.v(20,r+129,c-27)),', '),E.v(20,r+129,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+129,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+153,c-22)),' ['),E.v(20,r+153,c-27)),', '),E.v(20,r+153,c-21)),'] ') if X.truthy(X.eq(E.v(20,r+153,c-17),'ĐỦ')) else '')))

def f28(E, r, c):
    # Ô AD25
    # = TRIM(IF(L121="ĐỦ","• "&G121&" ["&B121&", "&H121&"] ","")&IF(L123="ĐỦ","• "&G123&" ["&B123&", "&H123&"] ","")&IF(L124="ĐỦ","• "&G124&" ["&B124&", "&H12
    #   4&"] ","")&IF(L127="ĐỦ","• "&G127&" ["&B127&", "&H127&"] ","")&IF(L128="ĐỦ","• "&G128&" ["&B128&", "&H128&"] ","")&IF(L132="ĐỦ","• "&G132&" ["&B132&",
    #    "&H132&"] ","")&IF(L133="ĐỦ","• "&G133&" ["&B133&", "&H133&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+96,c-23)),' ['),E.v(20,r+96,c-28)),', '),E.v(20,r+96,c-22)),'] ') if X.truthy(X.eq(E.v(20,r+96,c-18),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+98,c-23)),' ['),E.v(20,r+98,c-28)),', '),E.v(20,r+98,c-22)),'] ') if X.truthy(X.eq(E.v(20,r+98,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+99,c-23)),' ['),E.v(20,r+99,c-28)),', '),E.v(20,r+99,c-22)),'] ') if X.truthy(X.eq(E.v(20,r+99,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+102,c-23)),' ['),E.v(20,r+102,c-28)),', '),E.v(20,r+102,c-22)),'] ') if X.truthy(X.eq(E.v(20,r+102,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+103,c-23)),' ['),E.v(20,r+103,c-28)),', '),E.v(20,r+103,c-22)),'] ') if X.truthy(X.eq(E.v(20,r+103,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+107,c-23)),' ['),E.v(20,r+107,c-28)),', '),E.v(20,r+107,c-22)),'] ') if X.truthy(X.eq(E.v(20,r+107,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+108,c-23)),' ['),E.v(20,r+108,c-28)),', '),E.v(20,r+108,c-22)),'] ') if X.truthy(X.eq(E.v(20,r+108,c-18),'ĐỦ')) else '')))

def f29(E, r, c):
    # Ô AE25
    # = TRIM(IF(L144="ĐỦ","• "&G144&" ["&B144&", "&H144&"] ","")&IF(L149="ĐỦ","• "&G149&" ["&B149&", "&H149&"] ","")&IF(L151="ĐỦ","• "&G151&" ["&B151&", "&H15
    #   1&"] ","")&IF(L155="ĐỦ","• "&G155&" ["&B155&", "&H155&"] ","")&IF(L156="ĐỦ","• "&G156&" ["&B156&", "&H156&"] ","")&IF(L158="ĐỦ","• "&G158&" ["&B158&",
    #    "&H158&"] ","")&IF(L159="ĐỦ","• "&G159&" ["&B159&", "&H159&"] ","")&IF(L160="ĐỦ","• "&G160&" ["&B160&", "&H160&"] ","")&IF(L161="ĐỦ","• "&G161&" ["&B
    #   161&", "&H161&"] ","")&IF(L162="ĐỦ","• "&G162&" ["&B162&", "&H162&"] ","")&IF(L163="ĐỦ","• "&G163&" ["&B163&", "&H163&"] ","")&IF(L164="ĐỦ","• "&G164&
    #   " ["&B164&", "&H164&"] ","")&IF(L165="ĐỦ","• "&G165&" ["&B165&", "&H165&"] ","")&IF(L166="ĐỦ","• "&G166&" ["&B166&", "&H166&"] ","")&IF(L169="ĐỦ","• "
    #   &G169&" ["&B169&", "&H169&"] ","")&IF(L170="ĐỦ","• "&G170&" ["&B170&", "&H170&"] ","")&IF(L171="ĐỦ","• "&G171&" ["&B171&", "&H171&"] ","")&IF(L172="ĐỦ
    #   ","• "&G172&" ["&B172&", "&H172&"] ","")&IF(L173="ĐỦ","• "&G173&" ["&B173&", "&H173&"] ","")&IF(L174="ĐỦ","• "&G174&" ["&B174&", "&H174&"] ","")&IF(L1
    #   75="ĐỦ","• "&G175&" ["&B175&", "&H175&"] ","")&IF(L176="ĐỦ","• "&G176&" ["&B176&", "&H176&"] ","")&IF(L177="ĐỦ","• "&G177&" ["&B177&", "&H177&"] ","")
    #   &IF(L179="ĐỦ","• "&G179&" ["&B179&", "&H179&"] ","")&IF(L180="ĐỦ","• "&G180&" ["&B180&", "&H180&"] ","")&IF(L183="ĐỦ","• "&G183&" ["&B183&", "&H183&"]
    #    ","")&IF(L184="ĐỦ","• "&G184&" ["&B184&", "&H184&"] ","")&IF(L185="ĐỦ","• "&G185&" ["&B185&", "&H185&"] ","")&IF(L186="ĐỦ","• "&G186&" ["&B186&", "&H
    #   186&"] ","")&IF(L187="ĐỦ","• "&G187&" ["&B187&", "&H187&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+119,c-24)),' ['),E.v(20,r+119,c-29)),', '),E.v(20,r+119,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+119,c-19),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+124,c-24)),' ['),E.v(20,r+124,c-29)),', '),E.v(20,r+124,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+124,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+126,c-24)),' ['),E.v(20,r+126,c-29)),', '),E.v(20,r+126,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+126,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+130,c-24)),' ['),E.v(20,r+130,c-29)),', '),E.v(20,r+130,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+130,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+131,c-24)),' ['),E.v(20,r+131,c-29)),', '),E.v(20,r+131,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+131,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+133,c-24)),' ['),E.v(20,r+133,c-29)),', '),E.v(20,r+133,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+133,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+134,c-24)),' ['),E.v(20,r+134,c-29)),', '),E.v(20,r+134,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+134,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+135,c-24)),' ['),E.v(20,r+135,c-29)),', '),E.v(20,r+135,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+135,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+136,c-24)),' ['),E.v(20,r+136,c-29)),', '),E.v(20,r+136,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+136,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+137,c-24)),' ['),E.v(20,r+137,c-29)),', '),E.v(20,r+137,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+137,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+138,c-24)),' ['),E.v(20,r+138,c-29)),', '),E.v(20,r+138,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+138,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+139,c-24)),' ['),E.v(20,r+139,c-29)),', '),E.v(20,r+139,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+139,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+140,c-24)),' ['),E.v(20,r+140,c-29)),', '),E.v(20,r+140,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+140,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+141,c-24)),' ['),E.v(20,r+141,c-29)),', '),E.v(20,r+141,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+141,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+144,c-24)),' ['),E.v(20,r+144,c-29)),', '),E.v(20,r+144,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+144,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+145,c-24)),' ['),E.v(20,r+145,c-29)),', '),E.v(20,r+145,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+145,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+146,c-24)),' ['),E.v(20,r+146,c-29)),', '),E.v(20,r+146,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+146,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+147,c-24)),' ['),E.v(20,r+147,c-29)),', '),E.v(20,r+147,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+147,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+148,c-24)),' ['),E.v(20,r+148,c-29)),', '),E.v(20,r+148,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+148,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+149,c-24)),' ['),E.v(20,r+149,c-29)),', '),E.v(20,r+149,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+149,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+150,c-24)),' ['),E.v(20,r+150,c-29)),', '),E.v(20,r+150,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+150,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+151,c-24)),' ['),E.v(20,r+151,c-29)),', '),E.v(20,r+151,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+151,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+152,c-24)),' ['),E.v(20,r+152,c-29)),', '),E.v(20,r+152,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+152,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+154,c-24)),' ['),E.v(20,r+154,c-29)),', '),E.v(20,r+154,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+154,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+155,c-24)),' ['),E.v(20,r+155,c-29)),', '),E.v(20,r+155,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+155,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+158,c-24)),' ['),E.v(20,r+158,c-29)),', '),E.v(20,r+158,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+158,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+159,c-24)),' ['),E.v(20,r+159,c-29)),', '),E.v(20,r+159,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+159,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+160,c-24)),' ['),E.v(20,r+160,c-29)),', '),E.v(20,r+160,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+160,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+161,c-24)),' ['),E.v(20,r+161,c-29)),', '),E.v(20,r+161,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+161,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+162,c-24)),' ['),E.v(20,r+162,c-29)),', '),E.v(20,r+162,c-23)),'] ') if X.truthy(X.eq(E.v(20,r+162,c-19),'ĐỦ')) else '')))

def f30(E, r, c):
    # Ô AF25
    # = TRIM(IF(L112="ĐỦ","• "&G112&" ["&B112&", "&H112&"] ","")&IF(L115="ĐỦ","• "&G115&" ["&B115&", "&H115&"] ","")&IF(L125="ĐỦ","• "&G125&" ["&B125&", "&H12
    #   5&"] ","")&IF(L126="ĐỦ","• "&G126&" ["&B126&", "&H126&"] ","")&IF(L129="ĐỦ","• "&G129&" ["&B129&", "&H129&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+87,c-25)),' ['),E.v(20,r+87,c-30)),', '),E.v(20,r+87,c-24)),'] ') if X.truthy(X.eq(E.v(20,r+87,c-20),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+90,c-25)),' ['),E.v(20,r+90,c-30)),', '),E.v(20,r+90,c-24)),'] ') if X.truthy(X.eq(E.v(20,r+90,c-20),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+100,c-25)),' ['),E.v(20,r+100,c-30)),', '),E.v(20,r+100,c-24)),'] ') if X.truthy(X.eq(E.v(20,r+100,c-20),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+101,c-25)),' ['),E.v(20,r+101,c-30)),', '),E.v(20,r+101,c-24)),'] ') if X.truthy(X.eq(E.v(20,r+101,c-20),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+104,c-25)),' ['),E.v(20,r+104,c-30)),', '),E.v(20,r+104,c-24)),'] ') if X.truthy(X.eq(E.v(20,r+104,c-20),'ĐỦ')) else '')))

def f31(E, r, c):
    # Ô AG25
    # = TRIM(IF(L145="ĐỦ","• "&G145&" ["&B145&", "&H145&"] ","")&IF(L147="ĐỦ","• "&G147&" ["&B147&", "&H147&"] ","")&IF(L157="ĐỦ","• "&G157&" ["&B157&", "&H15
    #   7&"] ","")&IF(L167="ĐỦ","• "&G167&" ["&B167&", "&H167&"] ","")&IF(L168="ĐỦ","• "&G168&" ["&B168&", "&H168&"] ","")&IF(L181="ĐỦ","• "&G181&" ["&B181&",
    #    "&H181&"] ","")&IF(L182="ĐỦ","• "&G182&" ["&B182&", "&H182&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+120,c-26)),' ['),E.v(20,r+120,c-31)),', '),E.v(20,r+120,c-25)),'] ') if X.truthy(X.eq(E.v(20,r+120,c-21),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+122,c-26)),' ['),E.v(20,r+122,c-31)),', '),E.v(20,r+122,c-25)),'] ') if X.truthy(X.eq(E.v(20,r+122,c-21),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+132,c-26)),' ['),E.v(20,r+132,c-31)),', '),E.v(20,r+132,c-25)),'] ') if X.truthy(X.eq(E.v(20,r+132,c-21),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+142,c-26)),' ['),E.v(20,r+142,c-31)),', '),E.v(20,r+142,c-25)),'] ') if X.truthy(X.eq(E.v(20,r+142,c-21),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+143,c-26)),' ['),E.v(20,r+143,c-31)),', '),E.v(20,r+143,c-25)),'] ') if X.truthy(X.eq(E.v(20,r+143,c-21),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+156,c-26)),' ['),E.v(20,r+156,c-31)),', '),E.v(20,r+156,c-25)),'] ') if X.truthy(X.eq(E.v(20,r+156,c-21),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r+157,c-26)),' ['),E.v(20,r+157,c-31)),', '),E.v(20,r+157,c-25)),'] ') if X.truthy(X.eq(E.v(20,r+157,c-21),'ĐỦ')) else '')))

def f32(E, r, c):
    # Ô C29
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+0,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),1.0))

def f33(E, r, c):
    # Ô E29
    # = 5*SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+0,5))*{1,1,0,1})+COUNTIF(TRUNG_GIAN!$C$21:$F$23,MOD(TRUNG_GIAN!$E$7+0,5))
    return X.add(X.mul(5.0,X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),0.0), 5.0)),X.Arr([[1.0, 1.0, 0.0, 1.0]])))),X.COUNTIF(E.rg(23,21,3,23,6), X.MOD(X.add(E.v(23,7,5),0.0), 5.0)))

def f34(E, r, c):
    # Ô F29, F30, F31, F32, F33
    # = IF(SUM($E$29:$E$33)=0,0,E29/SUM($E$29:$E$33))
    return (0.0 if X.truthy(X.eq(X.SUM(E.rg(20,29,5,33,5)),0.0)) else X.div(E.v(20,r,c-1),X.SUM(E.rg(20,29,5,33,5))))

def f35(E, r, c):
    # Ô G29
    # = DUNG_HY_KY!$C$12
    return E.v(8,12,3)

def f36(E, r, c):
    # Ô H29, H30, H31, H32, H33
    # = IF(SUM($G$29:$G$33)=0,0,G29/SUM($G$29:$G$33))
    return (0.0 if X.truthy(X.eq(X.SUM(E.rg(20,29,7,33,7)),0.0)) else X.div(E.v(20,r,c-1),X.SUM(E.rg(20,29,7,33,7))))

def f37(E, r, c):
    # Ô I29
    # = ROUND(DUNG_HY_KY!$D$12,1)
    return X.ROUND(E.v(8,12,4), 1.0)

def f38(E, r, c):
    # Ô J29, J30, J31, J32, J33
    # = IF(SUM($I$29:$I$33)=0,0,I29/SUM($I$29:$I$33))
    return (0.0 if X.truthy(X.eq(X.SUM(E.rg(20,29,9,33,9)),0.0)) else X.div(E.v(20,r,c-1),X.SUM(E.rg(20,29,9,33,9))))

def f39(E, r, c):
    # Ô K29, K30, K31, K32, K33
    # = ROUND(100*(F29-J29),0)
    return X.ROUND(X.mul(100.0,X.sub(E.v(20,r,c-5),E.v(20,r,c-1))), 0.0)

def f40(E, r, c):
    # Ô C30
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))

def f41(E, r, c):
    # Ô E30
    # = 5*SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+1,5))*{1,1,0,1})+COUNTIF(TRUNG_GIAN!$C$21:$F$23,MOD(TRUNG_GIAN!$E$7+1,5))
    return X.add(X.mul(5.0,X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),1.0), 5.0)),X.Arr([[1.0, 1.0, 0.0, 1.0]])))),X.COUNTIF(E.rg(23,21,3,23,6), X.MOD(X.add(E.v(23,7,5),1.0), 5.0)))

def f42(E, r, c):
    # Ô G30
    # = DUNG_HY_KY!$C$13
    return E.v(8,13,3)

def f43(E, r, c):
    # Ô I30
    # = ROUND(DUNG_HY_KY!$D$13,1)
    return X.ROUND(E.v(8,13,4), 1.0)

def f44(E, r, c):
    # Ô C31
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))

def f45(E, r, c):
    # Ô E31
    # = 5*SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+2,5))*{1,1,0,1})+COUNTIF(TRUNG_GIAN!$C$21:$F$23,MOD(TRUNG_GIAN!$E$7+2,5))
    return X.add(X.mul(5.0,X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),2.0), 5.0)),X.Arr([[1.0, 1.0, 0.0, 1.0]])))),X.COUNTIF(E.rg(23,21,3,23,6), X.MOD(X.add(E.v(23,7,5),2.0), 5.0)))

def f46(E, r, c):
    # Ô G31
    # = DUNG_HY_KY!$C$14
    return E.v(8,14,3)

def f47(E, r, c):
    # Ô I31
    # = ROUND(DUNG_HY_KY!$D$14,1)
    return X.ROUND(E.v(8,14,4), 1.0)

def f48(E, r, c):
    # Ô C32
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+3,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0))

def f49(E, r, c):
    # Ô E32
    # = 5*SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+3,5))*{1,1,0,1})+COUNTIF(TRUNG_GIAN!$C$21:$F$23,MOD(TRUNG_GIAN!$E$7+3,5))
    return X.add(X.mul(5.0,X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),3.0), 5.0)),X.Arr([[1.0, 1.0, 0.0, 1.0]])))),X.COUNTIF(E.rg(23,21,3,23,6), X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))

def f50(E, r, c):
    # Ô G32
    # = DUNG_HY_KY!$C$15
    return E.v(8,15,3)

def f51(E, r, c):
    # Ô I32
    # = ROUND(DUNG_HY_KY!$D$15,1)
    return X.ROUND(E.v(8,15,4), 1.0)

def f52(E, r, c):
    # Ô C33
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+4,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))

def f53(E, r, c):
    # Ô E33
    # = 5*SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+4,5))*{1,1,0,1})+COUNTIF(TRUNG_GIAN!$C$21:$F$23,MOD(TRUNG_GIAN!$E$7+4,5))
    return X.add(X.mul(5.0,X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),4.0), 5.0)),X.Arr([[1.0, 1.0, 0.0, 1.0]])))),X.COUNTIF(E.rg(23,21,3,23,6), X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))

def f54(E, r, c):
    # Ô G33
    # = DUNG_HY_KY!$C$16
    return E.v(8,16,3)

def f55(E, r, c):
    # Ô I33
    # = ROUND(DUNG_HY_KY!$D$16,1)
    return X.ROUND(E.v(8,16,4), 1.0)

def f56(E, r, c):
    # Ô E34, F34, G34, H34, I34, J34
    # = SUM(E29:E33)
    return X.SUM(E.rg(20,r-5,c,r-1,c))

def f57(E, r, c):
    # Ô C35
    # = "Cách 1: "&TRIM(IF(F29=MAX(F$29:F$33),$B$29&" ("&$D$29&") ","")&IF(F30=MAX(F$29:F$33),$B$30&" ("&$D$30&") ","")&IF(F31=MAX(F$29:F$33),$B$31&" ("&$D$31
    #   &") ","")&IF(F32=MAX(F$29:F$33),$B$32&" ("&$D$32&") ","")&IF(F33=MAX(F$29:F$33),$B$33&" ("&$D$33&") ",""))&" / "&TRIM(IF(F29=MIN(F$29:F$33),$B$29&" ("
    #   &$D$29&") ","")&IF(F30=MIN(F$29:F$33),$B$30&" ("&$D$30&") ","")&IF(F31=MIN(F$29:F$33),$B$31&" ("&$D$31&") ","")&IF(F32=MIN(F$29:F$33),$B$32&" ("&$D$32
    #   &") ","")&IF(F33=MIN(F$29:F$33),$B$33&" ("&$D$33&") ",""))&CHAR(10)&"Cách 2: "&TRIM(IF(H29=MAX(H$29:H$33),$B$29&" ("&$D$29&") ","")&IF(H30=MAX(H$29:H$
    #   33),$B$30&" ("&$D$30&") ","")&IF(H31=MAX(H$29:H$33),$B$31&" ("&$D$31&") ","")&IF(H32=MAX(H$29:H$33),$B$32&" ("&$D$32&") ","")&IF(H33=MAX(H$29:H$33),$B
    #   $33&" ("&$D$33&") ",""))&" / "&TRIM(IF(H29=MIN(H$29:H$33),$B$29&" ("&$D$29&") ","")&IF(H30=MIN(H$29:H$33),$B$30&" ("&$D$30&") ","")&IF(H31=MIN(H$29:H$
    #   33),$B$31&" ("&$D$31&") ","")&IF(H32=MIN(H$29:H$33),$B$32&" ("&$D$32&") ","")&IF(H33=MIN(H$29:H$33),$B$33&" ("&$D$33&") ",""))&CHAR(10)&"Cách 3: "&TRI
    #   M(IF(J29=MAX(J$29:J$33),$B$29&" ("&$D$29&") ","")&IF(J30=MAX(J$29:J$33),$B$30&" ("&$D$30&") ","")&IF(J31=MAX(J$29:J$33),$B$31&" ("&$D$31&") ","")&IF(J
    #   32=MAX(J$29:J$33),$B$32&" ("&$D$32&") ","")&IF(J33=MAX(J$29:J$33),$B$33&" ("&$D$33&") ",""))&" / "&TRIM(IF(J29=MIN(J$29:J$33),$B$29&" ("&$D$29&") ",""
    #   )&IF(J30=MIN(J$29:J$33),$B$30&" ("&$D$30&") ","")&IF(J31=MIN(J$29:J$33),$B$31&" ("&$D$31&") ","")&IF(J32=MIN(J$29:J$33),$B$32&" ("&$D$32&") ","")&IF(J
    #   33=MIN(J$29:J$33),$B$33&" ("&$D$33&") ",""))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Cách 1: ',X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(20,29,2),' ('),E.v(20,29,4)),') ') if X.truthy(X.eq(E.v(20,r-6,c+3),X.MAX(E.rg(20,29,c+3,33,c+3)))) else ''),(X.cat(X.cat(X.cat(E.v(20,30,2),' ('),E.v(20,30,4)),') ') if X.truthy(X.eq(E.v(20,r-5,c+3),X.MAX(E.rg(20,29,c+3,33,c+3)))) else '')),(X.cat(X.cat(X.cat(E.v(20,31,2),' ('),E.v(20,31,4)),') ') if X.truthy(X.eq(E.v(20,r-4,c+3),X.MAX(E.rg(20,29,c+3,33,c+3)))) else '')),(X.cat(X.cat(X.cat(E.v(20,32,2),' ('),E.v(20,32,4)),') ') if X.truthy(X.eq(E.v(20,r-3,c+3),X.MAX(E.rg(20,29,c+3,33,c+3)))) else '')),(X.cat(X.cat(X.cat(E.v(20,33,2),' ('),E.v(20,33,4)),') ') if X.truthy(X.eq(E.v(20,r-2,c+3),X.MAX(E.rg(20,29,c+3,33,c+3)))) else '')))),' / '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(20,29,2),' ('),E.v(20,29,4)),') ') if X.truthy(X.eq(E.v(20,r-6,c+3),X.MIN(E.rg(20,29,c+3,33,c+3)))) else ''),(X.cat(X.cat(X.cat(E.v(20,30,2),' ('),E.v(20,30,4)),') ') if X.truthy(X.eq(E.v(20,r-5,c+3),X.MIN(E.rg(20,29,c+3,33,c+3)))) else '')),(X.cat(X.cat(X.cat(E.v(20,31,2),' ('),E.v(20,31,4)),') ') if X.truthy(X.eq(E.v(20,r-4,c+3),X.MIN(E.rg(20,29,c+3,33,c+3)))) else '')),(X.cat(X.cat(X.cat(E.v(20,32,2),' ('),E.v(20,32,4)),') ') if X.truthy(X.eq(E.v(20,r-3,c+3),X.MIN(E.rg(20,29,c+3,33,c+3)))) else '')),(X.cat(X.cat(X.cat(E.v(20,33,2),' ('),E.v(20,33,4)),') ') if X.truthy(X.eq(E.v(20,r-2,c+3),X.MIN(E.rg(20,29,c+3,33,c+3)))) else '')))),X.CHAR(10.0)),'Cách 2: '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(20,29,2),' ('),E.v(20,29,4)),') ') if X.truthy(X.eq(E.v(20,r-6,c+5),X.MAX(E.rg(20,29,c+5,33,c+5)))) else ''),(X.cat(X.cat(X.cat(E.v(20,30,2),' ('),E.v(20,30,4)),') ') if X.truthy(X.eq(E.v(20,r-5,c+5),X.MAX(E.rg(20,29,c+5,33,c+5)))) else '')),(X.cat(X.cat(X.cat(E.v(20,31,2),' ('),E.v(20,31,4)),') ') if X.truthy(X.eq(E.v(20,r-4,c+5),X.MAX(E.rg(20,29,c+5,33,c+5)))) else '')),(X.cat(X.cat(X.cat(E.v(20,32,2),' ('),E.v(20,32,4)),') ') if X.truthy(X.eq(E.v(20,r-3,c+5),X.MAX(E.rg(20,29,c+5,33,c+5)))) else '')),(X.cat(X.cat(X.cat(E.v(20,33,2),' ('),E.v(20,33,4)),') ') if X.truthy(X.eq(E.v(20,r-2,c+5),X.MAX(E.rg(20,29,c+5,33,c+5)))) else '')))),' / '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(20,29,2),' ('),E.v(20,29,4)),') ') if X.truthy(X.eq(E.v(20,r-6,c+5),X.MIN(E.rg(20,29,c+5,33,c+5)))) else ''),(X.cat(X.cat(X.cat(E.v(20,30,2),' ('),E.v(20,30,4)),') ') if X.truthy(X.eq(E.v(20,r-5,c+5),X.MIN(E.rg(20,29,c+5,33,c+5)))) else '')),(X.cat(X.cat(X.cat(E.v(20,31,2),' ('),E.v(20,31,4)),') ') if X.truthy(X.eq(E.v(20,r-4,c+5),X.MIN(E.rg(20,29,c+5,33,c+5)))) else '')),(X.cat(X.cat(X.cat(E.v(20,32,2),' ('),E.v(20,32,4)),') ') if X.truthy(X.eq(E.v(20,r-3,c+5),X.MIN(E.rg(20,29,c+5,33,c+5)))) else '')),(X.cat(X.cat(X.cat(E.v(20,33,2),' ('),E.v(20,33,4)),') ') if X.truthy(X.eq(E.v(20,r-2,c+5),X.MIN(E.rg(20,29,c+5,33,c+5)))) else '')))),X.CHAR(10.0)),'Cách 3: '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(20,29,2),' ('),E.v(20,29,4)),') ') if X.truthy(X.eq(E.v(20,r-6,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else ''),(X.cat(X.cat(X.cat(E.v(20,30,2),' ('),E.v(20,30,4)),') ') if X.truthy(X.eq(E.v(20,r-5,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,31,2),' ('),E.v(20,31,4)),') ') if X.truthy(X.eq(E.v(20,r-4,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,32,2),' ('),E.v(20,32,4)),') ') if X.truthy(X.eq(E.v(20,r-3,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,33,2),' ('),E.v(20,33,4)),') ') if X.truthy(X.eq(E.v(20,r-2,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')))),' / '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(20,29,2),' ('),E.v(20,29,4)),') ') if X.truthy(X.eq(E.v(20,r-6,c+7),X.MIN(E.rg(20,29,c+7,33,c+7)))) else ''),(X.cat(X.cat(X.cat(E.v(20,30,2),' ('),E.v(20,30,4)),') ') if X.truthy(X.eq(E.v(20,r-5,c+7),X.MIN(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,31,2),' ('),E.v(20,31,4)),') ') if X.truthy(X.eq(E.v(20,r-4,c+7),X.MIN(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,32,2),' ('),E.v(20,32,4)),') ') if X.truthy(X.eq(E.v(20,r-3,c+7),X.MIN(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,33,2),' ('),E.v(20,33,4)),') ') if X.truthy(X.eq(E.v(20,r-2,c+7),X.MIN(E.rg(20,29,c+7,33,c+7)))) else ''))))

def f58(E, r, c):
    # Ô C36
    # = "Số nhóm dưới 50%: Cách 1 = "&COUNTIF(F29:F33,"<0.5")&", Cách 2 = "&COUNTIF(H29:H33,"<0.5")&", Cách 3 = "&COUNTIF(J29:J33,"<0.5")&". Năm phần cộng lại
    #    luôn bằng 100% nên nhiều nhất chỉ một nhóm vượt 50% – ngưỡng này gần như luôn kết luận ""cần hóa giải"" và không có trong sách."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Số nhóm dưới 50%: Cách 1 = ',X.COUNTIF(E.rg(20,r-7,c+3,r-3,c+3), '<0.5')),', Cách 2 = '),X.COUNTIF(E.rg(20,r-7,c+5,r-3,c+5), '<0.5')),', Cách 3 = '),X.COUNTIF(E.rg(20,r-7,c+7,r-3,c+7), '<0.5')),'. Năm phần cộng lại luôn bằng 100% nên nhiều nhất chỉ một nhóm vượt 50% – ngưỡng này gần như luôn kết luận "cần hóa giải" và không có trong sách.')

def f59(E, r, c):
    # Ô C37
    # = "Cách 3 (có lệnh tháng, xung, hợp) là số đo gần sách nhất; nhóm nhiều nhất theo Cách 3 là "&TRIM(IF(J29=MAX(J$29:J$33),$B$29&" ("&$D$29&") ","")&IF(J3
    #   0=MAX(J$29:J$33),$B$30&" ("&$D$30&") ","")&IF(J31=MAX(J$29:J$33),$B$31&" ("&$D$31&") ","")&IF(J32=MAX(J$29:J$33),$B$32&" ("&$D$32&") ","")&IF(J33=MAX(
    #   J$29:J$33),$B$33&" ("&$D$33&") ",""))&". Theo sách, nhóm nhiều nhất là căn cứ chọn dụng thần (P1 tr.99, 104-108), không phải điểm năng khiếu: tố chất 
    #   đọc ở mục A–D. Chênh lệch cột K cho thấy Cách 1 đánh giá cao/thấp nhóm nào so với cách tính độ của sách."
    return X.cat(X.cat('Cách 3 (có lệnh tháng, xung, hợp) là số đo gần sách nhất; nhóm nhiều nhất theo Cách 3 là ',X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(20,29,2),' ('),E.v(20,29,4)),') ') if X.truthy(X.eq(E.v(20,r-8,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else ''),(X.cat(X.cat(X.cat(E.v(20,30,2),' ('),E.v(20,30,4)),') ') if X.truthy(X.eq(E.v(20,r-7,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,31,2),' ('),E.v(20,31,4)),') ') if X.truthy(X.eq(E.v(20,r-6,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,32,2),' ('),E.v(20,32,4)),') ') if X.truthy(X.eq(E.v(20,r-5,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')),(X.cat(X.cat(X.cat(E.v(20,33,2),' ('),E.v(20,33,4)),') ') if X.truthy(X.eq(E.v(20,r-4,c+7),X.MAX(E.rg(20,29,c+7,33,c+7)))) else '')))),'. Theo sách, nhóm nhiều nhất là căn cứ chọn dụng thần (P1 tr.99, 104-108), không phải điểm năng khiếu: tố chất đọc ở mục A–D. Chênh lệch cột K cho thấy Cách 1 đánh giá cao/thấp nhóm nào so với cách tính độ của sách.')

def f60(E, r, c):
    # Ô E42, F42
    # = IF((COUNTIF(THAN_SAT!$F$27:$F$30,"Có tác dụng")+COUNTIF(THAN_SAT!$F$39:$F$42,"Có tác dụng")>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIF(E.rg(12,27,6,30,6), 'Có tác dụng'),X.COUNTIF(E.rg(12,39,6,42,6), 'Có tác dụng')),0.0)) else '—')

def f61(E, r, c):
    # Ô J42, J43, J44, J45, J46, J47, J48, J49, J50, J51, J52, J53 … (161 ô)
    # = IF(AND(E42="ĐỦ",F42="ĐỦ"),"cả hai",IF(E42="ĐỦ","PP1",IF(F42="ĐỦ","PP2","")))
    return ('cả hai' if X.truthy(X.AND(X.eq(E.v(20,r,c-5),'ĐỦ'), X.eq(E.v(20,r,c-4),'ĐỦ'))) else ('PP1' if X.truthy(X.eq(E.v(20,r,c-5),'ĐỦ')) else ('PP2' if X.truthy(X.eq(E.v(20,r,c-4),'ĐỦ')) else '')))

def f62(E, r, c):
    # Ô K42
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P1 tr.140): Có Văn Xương còn tác dụng (không bị hình xung khắc hại) → "&IF(J42="","chưa đủ điều kiện ở cả h
    #   ai phương pháp – không áp dụng.","ĐỦ ("&J42&") ⇒ "&G42&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P1 tr.140): Có Văn Xương còn tác dụng (không bị hình xung khắc hại) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f63(E, r, c):
    # Ô L42, L43, L44, L45, L46, L47, L48, L49, L50, L51, L52, L53 … (161 ô)
    # = IF(DOI_CHIEU!$C$33=2,F42,E42)
    return (E.v(20,r,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(20,r,c-7))

def f64(E, r, c):
    # Ô E43, F43
    # = IF((THAN_SAT!$H$96="Có"),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(12,96,8),'Có')) else '—')

def f65(E, r, c):
    # Ô K43
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P1 tr.139): Có Thiên ất quý nhân còn tác dụng → "&IF(J43="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J43&") ⇒ "&G43&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P1 tr.139): Có Thiên ất quý nhân còn tác dụng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f66(E, r, c):
    # Ô E44, F44
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f67(E, r, c):
    # Ô K44
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.88; P2 tr.575): Có thương quan → "&
    #   IF(J44="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J44&") ⇒ "&G44&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P1 tr.88; P2 tr.575): Có thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f68(E, r, c):
    # Ô E45, F45, E130, F130
    # = IF((PP1_DINH_TINH!$D$100>=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,100,4),1.0)) else '—')

def f69(E, r, c):
    # Ô K45
    # = "Dữ kiện: "&"Số chữ Ấn lộ rõ: "&PP1_DINH_TINH!$D$100&" → Quy tắc (P2 tr.577): Ấn (chính/thiên ấn) lộ ở thiên can → "&IF(J45="","chưa đủ điều kiện ở cả
    #    hai phương pháp – không áp dụng.","ĐỦ ("&J45&") ⇒ "&G45&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn lộ rõ: '),E.v(9,100,4)),' → Quy tắc (P2 tr.577): Ấn (chính/thiên ấn) lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f70(E, r, c):
    # Ô E46, F46
    # = IF(AND((PP1_DINH_TINH!$F$100>=1),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,
    #   "Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,100,6),1.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f71(E, r, c):
    # Ô K46
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính ấn vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"C
    #   hính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $J$5:$J$20,1)>0,"Thiên ấn vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Q
    #   uy tắc (P2 tr.577): Có ấn nhưng không chữ ấn nào vượng tại lệnh (ấn hưu tù) → "&IF(J46="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J46&") ⇒ "&G46&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.577): Có ấn nhưng không chữ ấn nào vượng tại lệnh (ấn hưu tù) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f72(E, r, c):
    # Ô E47, F47, E110, F110
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))) else '—')

def f73(E, r, c):
    # Ô K47
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:
    #   $Y$20,0)>0,"Chính ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.573): Chính ấn nhiều (≥ P_NHIEU chữ) → "
    #   &IF(J47="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J47&") ⇒ "&G47&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.573): Chính ấn nhiều (≥ P_NHIEU chữ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f74(E, r, c):
    # Ô E48, F48, E116, F116
    # = IF((INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))) else '—')

def f75(E, r, c):
    # Ô K48
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.574): Hành Thủy nhiều (≥ P_NHIEU chữ) → "&IF(J48="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J48&") ⇒ "&G48&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.574): Hành Thủy nhiều (≥ P_NHIEU chữ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f76(E, r, c):
    # Ô E49, F49
    # = IF(OR(AND(OR(THAN_SAT!$F$27="Có tác dụng",THAN_SAT!$F$39="Có tác dụng"),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B
    #   $5:$B$20,0))>0)),AND(OR(THAN_SAT!$F$28="Có tác dụng",THAN_SAT!$F$40="Có tác dụng"),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$B$5:$B$20,1))>0)),AND(OR(THAN_SAT!$F$29="Có tác dụng",THAN_SAT!$F$41="Có tác dụng"),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$B$5:$B$20,2))>0)),AND(OR(THAN_SAT!$F$30="Có tác dụng",THAN_SAT!$F$42="Có tác dụng"),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(12,27,6),'Có tác dụng'), X.eq(E.v(12,39,6),'Có tác dụng')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)), X.AND(X.OR(X.eq(E.v(12,28,6),'Có tác dụng'), X.eq(E.v(12,40,6),'Có tác dụng')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)), X.AND(X.OR(X.eq(E.v(12,29,6),'Có tác dụng'), X.eq(E.v(12,41,6),'Có tác dụng')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)), X.AND(X.OR(X.eq(E.v(12,30,6),'Có tác dụng'), X.eq(E.v(12,42,6),'Có tác dụng')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)))) else '—')

def f77(E, r, c):
    # Ô K49
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,0)>0,"Thực thần ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:
    #   $F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thực thần ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Thực thần ở
    #    ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Thực thần ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$
    #   5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.574): Thực thần và Văn Xương cùng một trụ → "&IF(J49="","chưa đủ điều kiện ở cả hai phương pháp – không áp
    #    dụng.","ĐỦ ("&J49&") ⇒ "&G49&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Thực thần ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Thực thần ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thực thần ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thực thần ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.574): Thực thần và Văn Xương cùng một trụ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f78(E, r, c):
    # Ô E50, F50
    # = IF(AND((TRUNG_GIAN!$E$7=3),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),3.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f79(E, r, c):
    # Ô K50
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUN
    #   TIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.574): Nhật chủ Kim, có thương quan (Thủy) → "&IF
    #   (J50="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J50&") ⇒ "&G50&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.574): Nhật chủ Kim, có thương quan (Thủy) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f80(E, r, c):
    # Ô E51, F51
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$2
    #   0,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0))) else '—')

def f81(E, r, c):
    # Ô K51
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,1)>0,"Chính quan ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thất sát ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Chính ấn ở
    #    tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thiên ấn ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5
    #   :$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.574): Trụ tháng có cả quan và ấn → "&IF(J51="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J51&") ⇒ "&G51&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.574): Trụ tháng có cả quan và ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f82(E, r, c):
    # Ô E52, F52
    # = IF(OR(AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=
    #   4,TRUNG_GIAN!$E$6=10)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),10.0)))) else '—')

def f83(E, r, c):
    # Ô K52
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.301, 575): Trụ ngày là Nhâm Thìn, Canh Tuất, Canh Thìn hoặc Mậu Tuất (khôi canh) → "&IF(J52="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J52&") ⇒ "&G52&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.301, 575): Trụ ngày là Nhâm Thìn, Canh Tuất, Canh Thìn hoặc Mậu Tuất (khôi canh) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f84(E, r, c):
    # Ô E53, F53
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f85(E, r, c):
    # Ô K53
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Có cả thiên ấn
    #    (kiêu) và thực thần → "&IF(J53="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J53&") ⇒ "&G53&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.577): Có cả thiên ấn (kiêu) và thực thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f86(E, r, c):
    # Ô E54, F54
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$L$5:$L$20,0))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0),0.0)))) else '—')

def f87(E, r, c):
    # Ô K54
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thầ
    #   n: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$L$5:$L$20,0)>0,"Thực thần không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$
    #   20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Có thực thần nhưng mọi chữ thực thần đều bị khắc/hợp mất → "&IF(J54="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J54&") ⇒ "&G54&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.577): Có thực thần nhưng mọi chữ thực thần đều bị khắc/hợp mất → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f88(E, r, c):
    # Ô E55, F55
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1,SAO!$M$5:$M$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,13,20,13), 0.0)),0.0)) else '—')

def f89(E, r, c):
    # Ô K55
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J
    #   $20,1,SAO!$M$5:$M$20,0)>0,"Chính quan vượng tại lệnh không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:
    #   $J$20,1,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1,SAO!$M$5:$M$20,0)>0,"T
    #   hất sát vượng tại lệnh không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1,SAO!$M$5:$M$20,0)&" chữ;
    #    ",""))&" → Quy tắc (P1 tr.89): Quan tinh (chính quan/thất sát) có chữ vượng tại lệnh, không bị xung → "&IF(J55="","chưa đủ điều kiện ở cả hai phương 
    #   pháp – không áp dụng.","ĐỦ ("&J55&") ⇒ "&G55&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan vượng tại lệnh không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,13,20,13), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát vượng tại lệnh không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')))),' → Quy tắc (P1 tr.89): Quan tinh (chính quan/thất sát) có chữ vượng tại lệnh, không bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f90(E, r, c):
    # Ô E56, F56
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)
    #   )>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),0.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f91(E, r, c):
    # Ô K56
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính q
    #   uan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính qua
    #   n",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thất sát vượng tại
    #    lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.89): Có quan tinh như
    #   ng không chữ nào vượng tại lệnh → "&IF(J56="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J56&") ⇒ "&G56&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.89): Có quan tinh nhưng không chữ nào vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f92(E, r, c):
    # Ô E57, F57
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '—')

def f93(E, r, c):
    # Ô K57
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,0)>0,"Chính quan ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ",""))&" → Quy tắc (P1 t
    #   r.89): Chính quan ở trụ năm → "&IF(J57="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J57&") ⇒ "&G57&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.89): Chính quan ở trụ năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f94(E, r, c):
    # Ô E58, F58
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '—')

def f95(E, r, c):
    # Ô K58
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$2
    #   0,1)>0,"Thiên ấn ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.93
    #   ): Thiên ấn ở trụ tháng → "&IF(J58="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J58&") ⇒ "&G58&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''))),' → Quy tắc (P1 tr.93): Thiên ấn ở trụ tháng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f96(E, r, c):
    # Ô E59, F59
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0))) else '—')

def f97(E, r, c):
    # Ô K59
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can")>0,"Thực thần thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thực thần tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ",""))&" → Quy tắc (P1 tr.94): Thực thần có cả ở thiên can và địa chi → "&IF(J59="","c
    #   hưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J59&") ⇒ "&G59&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thực thần tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')))),' → Quy tắc (P1 tr.94): Thực thần có cả ở thiên can và địa chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f98(E, r, c):
    # Ô E60, F60
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '—')

def f99(E, r, c):
    # Ô K60
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can")>0,"Thiên ấn thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc
    #    (P1 tr.87): Thiên ấn lộ ở thiên can → "&IF(J60="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J60&") ⇒ "&G60&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''))),' → Quy tắc (P1 tr.87): Thiên ấn lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f100(E, r, c):
    # Ô E61, F61, E123, F123
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '—')

def f101(E, r, c):
    # Ô K61
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc
    #    (P1 tr.87): Thất sát lộ ở thiên can → "&IF(J61="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J61&") ⇒ "&G61&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''))),' → Quy tắc (P1 tr.87): Thất sát lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f102(E, r, c):
    # Ô E62, F62
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '—')

def f103(E, r, c):
    # Ô K62
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can")>0,"Chính quan thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Q
    #   uy tắc (P1 tr.87): Chính quan lộ ở thiên can → "&IF(J62="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J62&") ⇒ "&G62&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''))),' → Quy tắc (P1 tr.87): Chính quan lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f104(E, r, c):
    # Ô E63, F63
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$D$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,98,4),1.0))) else '—')

def f105(E, r, c):
    # Ô K63
    # = "Dữ kiện: "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 t
    #   r.575): Có thiên ấn và tài lộ → "&IF(J63="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J63&") ⇒ "&G63&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài lộ rõ: '),E.v(9,98,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.575): Có thiên ấn và tài lộ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f106(E, r, c):
    # Ô E64
    # = IF(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.ge(E.v(9,99,6),1.0))) else '—')

def f107(E, r, c):
    # Ô F64
    # = IF(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.ge(E.v(9,99,6),1.0))) else '—')

def f108(E, r, c):
    # Ô K64
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""
    #   ," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P2 tr.582): Nhật chủ vượng, trong tứ trụ có quan sát khắc chế → "&IF(J64="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J64&") ⇒ "&G64&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P2 tr.582): Nhật chủ vượng, trong tứ trụ có quan sát khắc chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f109(E, r, c):
    # Ô E65
    # = IF(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.eq(E.v(9,99,6),0.0))) else '—')

def f110(E, r, c):
    # Ô F65
    # = IF(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.eq(E.v(9,99,6),0.0))) else '—')

def f111(E, r, c):
    # Ô K65
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""
    #   ," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P2 tr.582): Nhật chủ vượng, không có quan sát khắc chế → "&IF(J65="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J65&") ⇒ "&G65&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P2 tr.582): Nhật chủ vượng, không có quan sát khắc chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f112(E, r, c):
    # Ô E66
    # = IF(AND((DUNG_HY_KY!$C$7=0),OR((PP1_DINH_TINH!$F$100>=1),(PP1_DINH_TINH!$F$96>=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.ge(E.v(9,100,6),1.0), X.ge(E.v(9,96,6),1.0)))) else '—')

def f113(E, r, c):
    # Ô F66
    # = IF(AND((DUNG_HY_KY!$D$7=0),OR((PP1_DINH_TINH!$F$100>=1),(PP1_DINH_TINH!$F$96>=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.ge(E.v(9,100,6),1.0), X.ge(E.v(9,96,6),1.0)))) else '—')

def f114(E, r, c):
    # Ô K66
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.582): Nhật chủ như
    #   ợc, có sinh phù (ấn hoặc tỉ kiếp) → "&IF(J66="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J66&") ⇒ "&G66&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.582): Nhật chủ nhược, có sinh phù (ấn hoặc tỉ kiếp) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f115(E, r, c):
    # Ô E67
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$100=0)),(PP1_DINH_TINH!$F$96=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.eq(E.v(9,100,6),0.0)), X.eq(E.v(9,96,6),0.0))) else '—')

def f116(E, r, c):
    # Ô F67
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$100=0)),(PP1_DINH_TINH!$F$96=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.eq(E.v(9,100,6),0.0)), X.eq(E.v(9,96,6),0.0))) else '—')

def f117(E, r, c):
    # Ô K67
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.582): Nhật chủ như
    #   ợc, không có ấn và tỉ kiếp → "&IF(J67="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J67&") ⇒ "&G67&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.582): Nhật chủ nhược, không có ấn và tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f118(E, r, c):
    # Ô E68
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f119(E, r, c):
    # Ô F68
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f120(E, r, c):
    # Ô K68
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$
    #   20,"Chính quan",SAO!$Y$5:$Y$20,0)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng 
    #   thần thuộc nhóm có Chính quan (chính cách), Chính quan có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J68="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J68&") ⇒ "&G68&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính quan (chính cách), Chính quan có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f121(E, r, c):
    # Ô E69
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:
    #   $Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f122(E, r, c):
    # Ô F69
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:
    #   $Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f123(E, r, c):
    # Ô K69
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y
    #   $20,0)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính qu
    #   an (chính cách), Chính quan quá nhiều (≥ P_NHIEU) → "&IF(J69="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J69&") ⇒ "&G69&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính quan (chính cách), Chính quan quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f124(E, r, c):
    # Ô E70
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f125(E, r, c):
    # Ô F70
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f126(E, r, c):
    # Ô K70
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Th
    #   ất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc n
    #   hóm có Thất sát (chính cách), Thất sát có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J70="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J70&") ⇒ "&G70&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thất sát (chính cách), Thất sát có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f127(E, r, c):
    # Ô E71
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f128(E, r, c):
    # Ô F71
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f129(E, r, c):
    # Ô K71
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$2
    #   0,0)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thất sát (chín
    #   h cách), Thất sát quá nhiều (≥ P_NHIEU) → "&IF(J71="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J71&") ⇒ "&G71&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thất sát (chính cách), Thất sát quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f130(E, r, c):
    # Ô E72
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),1.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f131(E, r, c):
    # Ô F72
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),1.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f132(E, r, c):
    # Ô K72
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,
    #   "Thực thần",SAO!$Y$5:$Y$20,0)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần t
    #   huộc nhóm có Thực thần (chính cách), Thực thần có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J72="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J72&") ⇒ "&G72&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thực thần (chính cách), Thực thần có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f133(E, r, c):
    # Ô E73
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=1)),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$
    #   Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),1.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f134(E, r, c):
    # Ô F73
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=1)),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$
    #   Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),1.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f135(E, r, c):
    # Ô K73
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$
    #   20,0)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thực thần (
    #   chính cách), Thực thần quá nhiều (≥ P_NHIEU) → "&IF(J73="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J73&") ⇒ "&G73&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thực thần (chính cách), Thực thần quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f136(E, r, c):
    # Ô E74
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),1.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f137(E, r, c):
    # Ô F74
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),1.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f138(E, r, c):
    # Ô K74
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:
    #   $F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581):
    #    Dụng thần thuộc nhóm có Thương quan (chính cách), Thương quan có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J74="","chưa đủ điều kiện ở cả hai phương phá
    #   p – không áp dụng.","ĐỦ ("&J74&") ⇒ "&G74&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thương quan (chính cách), Thương quan có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f139(E, r, c):
    # Ô E75
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=1)),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5
    #   :$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),1.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f140(E, r, c):
    # Ô F75
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=1)),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5
    #   :$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),1.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f141(E, r, c):
    # Ô K75
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$
    #   Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thươn
    #   g quan (chính cách), Thương quan quá nhiều (≥ P_NHIEU) → "&IF(J75="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J75&") ⇒ "&G75&"
    #   .")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thương quan (chính cách), Thương quan quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f142(E, r, c):
    # Ô E76
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0)),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f143(E, r, c):
    # Ô F76
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0)),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f144(E, r, c):
    # Ô K76
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1)>0,"Tỉ kiên: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ ki
    #   ên",SAO!$Y$5:$Y$20,0)>0,"Tỉ kiên: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có
    #    Tỉ kiên (chính cách), Tỉ kiên có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J76="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J76&"
    #   ) ⇒ "&G76&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Tỉ kiên: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Tỉ kiên: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Tỉ kiên (chính cách), Tỉ kiên có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f145(E, r, c):
    # Ô E77
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0)),(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$
    #   20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f146(E, r, c):
    # Ô F77
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0)),(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$
    #   20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f147(E, r, c):
    # Ô K77
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20
    #   ,0)>0,"Tỉ kiên: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Tỉ kiên (chính cá
    #   ch), Tỉ kiên quá nhiều (≥ P_NHIEU) → "&IF(J77="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J77&") ⇒ "&G77&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Tỉ kiên: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Tỉ kiên (chính cách), Tỉ kiên quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f148(E, r, c):
    # Ô E78
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0)),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f149(E, r, c):
    # Ô F78
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0)),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f150(E, r, c):
    # Ô K78
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Ki
    #   ếp tài",SAO!$Y$5:$Y$20,0)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc n
    #   hóm có Kiếp tài (chính cách), Kiếp tài có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J78="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J78&") ⇒ "&G78&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Kiếp tài (chính cách), Kiếp tài có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f151(E, r, c):
    # Ô E79
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0)),(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f152(E, r, c):
    # Ô F79
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0)),(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f153(E, r, c):
    # Ô K79
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$2
    #   0,0)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Kiếp tài (chín
    #   h cách), Kiếp tài quá nhiều (≥ P_NHIEU) → "&IF(J79="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J79&") ⇒ "&G79&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Kiếp tài (chính cách), Kiếp tài quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f154(E, r, c):
    # Ô E80
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f155(E, r, c):
    # Ô F80
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f156(E, r, c):
    # Ô K80
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1)>0,"Chính tài: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,
    #   "Chính tài",SAO!$Y$5:$Y$20,0)>0,"Chính tài: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần t
    #   huộc nhóm có Chính tài (chính cách), Chính tài có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J80="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J80&") ⇒ "&G80&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính tài (chính cách), Chính tài có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f157(E, r, c):
    # Ô E81
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$
    #   Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f158(E, r, c):
    # Ô F81
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$
    #   Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f159(E, r, c):
    # Ô K81
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$
    #   20,0)>0,"Chính tài: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính tài (
    #   chính cách), Chính tài quá nhiều (≥ P_NHIEU) → "&IF(J81="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J81&") ⇒ "&G81&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính tài (chính cách), Chính tài quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f160(E, r, c):
    # Ô E82
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f161(E, r, c):
    # Ô F82
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f162(E, r, c):
    # Ô K82
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1)>0,"Thiên tài: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,
    #   "Thiên tài",SAO!$Y$5:$Y$20,0)>0,"Thiên tài: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần t
    #   huộc nhóm có Thiên tài (chính cách), Thiên tài có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J82="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J82&") ⇒ "&G82&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thiên tài (chính cách), Thiên tài có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f163(E, r, c):
    # Ô E83
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$
    #   Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f164(E, r, c):
    # Ô F83
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$
    #   Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f165(E, r, c):
    # Ô K83
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$
    #   20,0)>0,"Thiên tài: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thiên tài (
    #   chính cách), Thiên tài quá nhiều (≥ P_NHIEU) → "&IF(J83="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J83&") ⇒ "&G83&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thiên tài (chính cách), Thiên tài quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f166(E, r, c):
    # Ô E84
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f167(E, r, c):
    # Ô F84
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f168(E, r, c):
    # Ô K84
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1)>0,"Chính ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Ch
    #   ính ấn",SAO!$Y$5:$Y$20,0)>0,"Chính ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc n
    #   hóm có Chính ấn (chính cách), Chính ấn có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J84="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J84&") ⇒ "&G84&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính ấn (chính cách), Chính ấn có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f169(E, r, c):
    # Ô E85
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4)),(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f170(E, r, c):
    # Ô F85
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4)),(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f171(E, r, c):
    # Ô K85
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$2
    #   0,0)>0,"Chính ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính ấn (chín
    #   h cách), Chính ấn quá nhiều (≥ P_NHIEU) → "&IF(J85="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J85&") ⇒ "&G85&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Chính ấn (chính cách), Chính ấn quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f172(E, r, c):
    # Ô E86
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f173(E, r, c):
    # Ô F86
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)<P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.lt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f174(E, r, c):
    # Ô K86
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Th
    #   iên ấn",SAO!$Y$5:$Y$20,0)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc n
    #   hóm có Thiên ấn (chính cách), Thiên ấn có mặt, chưa quá nhiều (< P_NHIEU) → "&IF(J86="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J86&") ⇒ "&G86&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thiên ấn (chính cách), Thiên ấn có mặt, chưa quá nhiều (< P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f175(E, r, c):
    # Ô E87
    # = IF(AND(AND((DUNG_HY_KY!$C$72=""),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,3),''), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f176(E, r, c):
    # Ô F87
    # = IF(AND(AND((DUNG_HY_KY!$D$72=""),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y
    #   $20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,72,4),''), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f177(E, r, c):
    # Ô K87
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$2
    #   0,0)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thiên ấn (chín
    #   h cách), Thiên ấn quá nhiều (≥ P_NHIEU) → "&IF(J87="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J87&") ⇒ "&G87&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.581): Dụng thần thuộc nhóm có Thiên ấn (chính cách), Thiên ấn quá nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f178(E, r, c):
    # Ô E88
    # = IF(OR(ISNUMBER(SEARCH("Khúc trực",DUNG_HY_KY!$C$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Khúc trực', E.v(8,72,3))))) else '—')

def f179(E, r, c):
    # Ô F88
    # = IF(OR(ISNUMBER(SEARCH("Khúc trực",DUNG_HY_KY!$D$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Khúc trực', E.v(8,72,4))))) else '—')

def f180(E, r, c):
    # Ô K88
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách khúc trực → "&IF(J88="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J88&") ⇒ "&G88&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách khúc trực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f181(E, r, c):
    # Ô E89
    # = IF(OR(ISNUMBER(SEARCH("Gia tường",DUNG_HY_KY!$C$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Gia tường', E.v(8,72,3))))) else '—')

def f182(E, r, c):
    # Ô F89
    # = IF(OR(ISNUMBER(SEARCH("Gia tường",DUNG_HY_KY!$D$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Gia tường', E.v(8,72,4))))) else '—')

def f183(E, r, c):
    # Ô K89
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách gia tường → "&IF(J89="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J89&") ⇒ "&G89&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách gia tường → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f184(E, r, c):
    # Ô E90
    # = IF(OR(ISNUMBER(SEARCH("Tòng cách",DUNG_HY_KY!$C$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Tòng cách', E.v(8,72,3))))) else '—')

def f185(E, r, c):
    # Ô F90
    # = IF(OR(ISNUMBER(SEARCH("Tòng cách",DUNG_HY_KY!$D$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Tòng cách', E.v(8,72,4))))) else '—')

def f186(E, r, c):
    # Ô K90
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách tòng cách → "&IF(J90="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J90&") ⇒ "&G90&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách tòng cách → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f187(E, r, c):
    # Ô E91
    # = IF(OR(ISNUMBER(SEARCH("Nhuận hạ",DUNG_HY_KY!$C$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Nhuận hạ', E.v(8,72,3))))) else '—')

def f188(E, r, c):
    # Ô F91
    # = IF(OR(ISNUMBER(SEARCH("Nhuận hạ",DUNG_HY_KY!$D$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Nhuận hạ', E.v(8,72,4))))) else '—')

def f189(E, r, c):
    # Ô K91
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách nhuận hạ → "&IF(J91="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng
    #   .","ĐỦ ("&J91&") ⇒ "&G91&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách nhuận hạ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f190(E, r, c):
    # Ô E92
    # = IF(OR(ISNUMBER(SEARCH("Viêm thượng",DUNG_HY_KY!$C$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Viêm thượng', E.v(8,72,3))))) else '—')

def f191(E, r, c):
    # Ô F92
    # = IF(OR(ISNUMBER(SEARCH("Viêm thượng",DUNG_HY_KY!$D$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Viêm thượng', E.v(8,72,4))))) else '—')

def f192(E, r, c):
    # Ô K92
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách viêm thượng → "&IF(J92="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J92&") ⇒ "&G92&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách viêm thượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f193(E, r, c):
    # Ô E93
    # = IF(OR(ISNUMBER(SEARCH("Tòng tài",DUNG_HY_KY!$C$72)),ISNUMBER(SEARCH("Tòng sát",DUNG_HY_KY!$C$72)),ISNUMBER(SEARCH("Tòng nhi",DUNG_HY_KY!$C$72))),"ĐỦ",
    #   "—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Tòng tài', E.v(8,72,3))), X.ISNUMBER(lambda: X.SEARCH('Tòng sát', E.v(8,72,3))), X.ISNUMBER(lambda: X.SEARCH('Tòng nhi', E.v(8,72,3))))) else '—')

def f194(E, r, c):
    # Ô F93
    # = IF(OR(ISNUMBER(SEARCH("Tòng tài",DUNG_HY_KY!$D$72)),ISNUMBER(SEARCH("Tòng sát",DUNG_HY_KY!$D$72)),ISNUMBER(SEARCH("Tòng nhi",DUNG_HY_KY!$D$72))),"ĐỦ",
    #   "—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Tòng tài', E.v(8,72,4))), X.ISNUMBER(lambda: X.SEARCH('Tòng sát', E.v(8,72,4))), X.ISNUMBER(lambda: X.SEARCH('Tòng nhi', E.v(8,72,4))))) else '—')

def f195(E, r, c):
    # Ô K93
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách tòng tài / tòng sát / tòng nhi → "&IF(J93="","chưa đủ điều kiện ở cả hai phươn
    #   g pháp – không áp dụng.","ĐỦ ("&J93&") ⇒ "&G93&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách tòng tài / tòng sát / tòng nhi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f196(E, r, c):
    # Ô E94
    # = IF(OR(ISNUMBER(SEARCH("Tòng vượng",DUNG_HY_KY!$C$72)),ISNUMBER(SEARCH("Tòng cường",DUNG_HY_KY!$C$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Tòng vượng', E.v(8,72,3))), X.ISNUMBER(lambda: X.SEARCH('Tòng cường', E.v(8,72,3))))) else '—')

def f197(E, r, c):
    # Ô F94
    # = IF(OR(ISNUMBER(SEARCH("Tòng vượng",DUNG_HY_KY!$D$72)),ISNUMBER(SEARCH("Tòng cường",DUNG_HY_KY!$D$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Tòng vượng', E.v(8,72,4))), X.ISNUMBER(lambda: X.SEARCH('Tòng cường', E.v(8,72,4))))) else '—')

def f198(E, r, c):
    # Ô K94
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách tòng vượng / tòng cường → "&IF(J94="","chưa đủ điều kiện ở cả hai phương pháp 
    #   – không áp dụng.","ĐỦ ("&J94&") ⇒ "&G94&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách tòng vượng / tòng cường → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f199(E, r, c):
    # Ô E95
    # = IF(OR(ISNUMBER(SEARCH("Hóa",DUNG_HY_KY!$C$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Hóa', E.v(8,72,3))))) else '—')

def f200(E, r, c):
    # Ô F95
    # = IF(OR(ISNUMBER(SEARCH("Hóa",DUNG_HY_KY!$D$72))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Hóa', E.v(8,72,4))))) else '—')

def f201(E, r, c):
    # Ô K95
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.582): Lá số thuộc cách hóa → "&IF(J95="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J95&") ⇒ "&G95&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.582): Lá số thuộc cách hóa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f202(E, r, c):
    # Ô E96, F96
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f203(E, r, c):
    # Ô K96
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can",SAO!$L$5:$L$20,0)>0,"Chính quan thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Chính quan lộ ở thiên can (không tính nhật can), còn tác dụng (không bị 
    #   khắc/hợp mất) → "&IF(J96="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J96&") ⇒ "&G96&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Chính quan lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f204(E, r, c):
    # Ô E97, F97
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f205(E, r, c):
    # Ô K97
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$L$5:$L$20,0)>0,"Thất sát thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Thất sát lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp
    #    mất) → "&IF(J97="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J97&") ⇒ "&G97&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Thất sát lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f206(E, r, c):
    # Ô E98, F98
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f207(E, r, c):
    # Ô K98
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$L$5:$L$20,0)>0,"Chính ấn thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Chính ấn lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp
    #    mất) → "&IF(J98="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J98&") ⇒ "&G98&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính ấn thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Chính ấn lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f208(E, r, c):
    # Ô E99, F99
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f209(E, r, c):
    # Ô K99
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$L$5:$L$20,0)>0,"Thiên ấn thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Thiên ấn lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp
    #    mất) → "&IF(J99="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J99&") ⇒ "&G99&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Thiên ấn lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f210(E, r, c):
    # Ô E100, F100
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f211(E, r, c):
    # Ô K100
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Can",SAO!$L$5:$L$20,0)>0,"Tỉ kiên thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Tỉ kiên lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất
    #   ) → "&IF(J100="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J100&") ⇒ "&G100&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Tỉ kiên thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Tỉ kiên lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f212(E, r, c):
    # Ô E101, F101
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f213(E, r, c):
    # Ô K101
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$L$5:$L$20,0)>0,"Kiếp tài thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Kiếp tài lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp
    #    mất) → "&IF(J101="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J101&") ⇒ "&G101&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Kiếp tài lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f214(E, r, c):
    # Ô E102, F102
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f215(E, r, c):
    # Ô K102
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$L$5:$L$20,0)>0,"Thực thần thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   C$5:$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Thực thần lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc
    #   /hợp mất) → "&IF(J102="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J102&") ⇒ "&G102&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Thực thần lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f216(E, r, c):
    # Ô E103, F103
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f217(E, r, c):
    # Ô K103
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Can",SAO!$L$5:$L$20,0)>0,"Thương quan thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Thương quan lộ ở thiên can (không tính nhật can), còn tác dụng (không
    #    bị khắc/hợp mất) → "&IF(J103="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J103&") ⇒ "&G103&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Thương quan lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f218(E, r, c):
    # Ô E104, F104
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f219(E, r, c):
    # Ô K104
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$L$5:$L$20,0)>0,"Chính tài thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   C$5:$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Chính tài lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc
    #   /hợp mất) → "&IF(J104="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J104&") ⇒ "&G104&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Chính tài lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f220(E, r, c):
    # Ô E105, F105
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f221(E, r, c):
    # Ô K105
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$L$5:$L$20,0)>0,"Thiên tài thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   C$5:$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.87-88): Thiên tài lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc
    #   /hợp mất) → "&IF(J105="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J105&") ⇒ "&G105&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.87-88): Thiên tài lộ ở thiên can (không tính nhật can), còn tác dụng (không bị khắc/hợp mất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f222(E, r, c):
    # Ô K110
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:
    #   $Y$20,0)>0,"Chính ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.573): Chính ấn ≥ P_NHIEU chữ → "&IF(J110
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J110&") ⇒ "&G110&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.573): Chính ấn ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f223(E, r, c):
    # Ô E111, F111
    # = IF(AND(AND(AND((PP1_DINH_TINH!$F$98>=1),(PP1_DINH_TINH!$F$99>=1)),(PP1_DINH_TINH!$F$100>=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.ge(E.v(9,98,6),1.0), X.ge(E.v(9,99,6),1.0)), X.ge(E.v(9,100,6),1.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f224(E, r, c):
    # Ô K111
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (
    #   chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Th
    #   ực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.573): Có đủ tài, quan sát, ấn, thực thần → "&IF(J111="","chưa đủ điều kiệ
    #   n ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J111&") ⇒ "&G111&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.573): Có đủ tài, quan sát, ấn, thực thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f225(E, r, c):
    # Ô E112, F112
    # = IF((INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2))) else '—')

def f226(E, r, c):
    # Ô K112
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.574): Kim ≥ P_NHIEU chữ → "&IF(J112="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J112&") ⇒ "&G112&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.574): Kim ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f227(E, r, c):
    # Ô E113, F113
    # = IF((INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2))) else '—')

def f228(E, r, c):
    # Ô K113
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.574): Mộc ≥ P_NHIEU chữ → "&IF(J113="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J113&") ⇒ "&G113&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.574): Mộc ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f229(E, r, c):
    # Ô E114, F114
    # = IF((INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2))) else '—')

def f230(E, r, c):
    # Ô K114
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.574): Thổ ≥ P_NHIEU chữ → "&IF(J114="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J114&") ⇒ "&G114&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.574): Thổ ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f231(E, r, c):
    # Ô E115, F115
    # = IF((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2))) else '—')

def f232(E, r, c):
    # Ô K115
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.574): Hỏa ≥ P_NHIEU chữ → "&IF(J115="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J115&") ⇒ "&G115&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.574): Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f233(E, r, c):
    # Ô K116
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.574): Thủy ≥ P_NHIEU chữ → "&IF(J116="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J116&") ⇒ "&G116&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.574): Thủy ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f234(E, r, c):
    # Ô E117
    # = IF(AND((DUNG_HY_KY!$C$7=0),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f235(E, r, c):
    # Ô F117
    # = IF(AND((DUNG_HY_KY!$D$7=0),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f236(E, r, c):
    # Ô K117
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$
    #   J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO
    #   !$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.574): Thân nhược, tài có lực → "&IF(J117="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.",
    #   "ĐỦ ("&J117&") ⇒ "&G117&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.574): Thân nhược, tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f237(E, r, c):
    # Ô E118, F118
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f238(E, r, c):
    # Ô K118
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên tà
    #   i: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.574): Có thiên tài → "&IF(J118="","chưa đ
    #   ủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J118&") ⇒ "&G118&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.574): Có thiên tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f239(E, r, c):
    # Ô E119, F119
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '—')

def f240(E, r, c):
    # Ô K119
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,3)>0,"Thiên tài ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.5
    #   74): Thiên tài ở trụ giờ → "&IF(J119="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J119&") ⇒ "&G119&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else ''))),' → Quy tắc (P2 tr.574): Thiên tài ở trụ giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f241(E, r, c):
    # Ô E120, F120
    # = IF((PP1_DINH_TINH!$D$98>=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,98,4),1.0)) else '—')

def f242(E, r, c):
    # Ô K120
    # = "Dữ kiện: "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&" → Quy tắc (P2 tr.574): Tài lộ ở thiên can → "&IF(J120="","chưa đủ điều kiện ở cả hai phương phá
    #   p – không áp dụng.","ĐỦ ("&J120&") ⇒ "&G120&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài lộ rõ: '),E.v(9,98,4)),' → Quy tắc (P2 tr.574): Tài lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f243(E, r, c):
    # Ô E121, F121
    # = IF(AND((PP1_DINH_TINH!$F$98>=1),(PP1_DINH_TINH!$D$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,6),1.0), X.eq(E.v(9,98,4),0.0))) else '—')

def f244(E, r, c):
    # Ô K121
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&" → Quy tắ
    #   c (P2 tr.574): Tài chỉ tàng trong chi, không lộ → "&IF(J121="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J121&") ⇒ "&G121&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Tài lộ rõ: '),E.v(9,98,4)),' → Quy tắc (P2 tr.574): Tài chỉ tàng trong chi, không lộ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f245(E, r, c):
    # Ô E122, F122
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f246(E, r, c):
    # Ô K122
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.575): Có thất sát và có ấn → "&IF(J122="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ (
    #   "&J122&") ⇒ "&G122&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.575): Có thất sát và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f247(E, r, c):
    # Ô K123
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc
    #    (P2 tr.575): Thất sát lộ ở thiên can → "&IF(J123="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J123&") ⇒ "&G123&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''))),' → Quy tắc (P2 tr.575): Thất sát lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f248(E, r, c):
    # Ô E124
    # = IF(AND((DUNG_HY_KY!$C$7=0),OR(TRUNG_GIAN!$E$28=6,TRUNG_GIAN!$E$28=7,TRUNG_GIAN!$E$28=8,TRUNG_GIAN!$E$28=9)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.eq(E.v(23,28,5),6.0), X.eq(E.v(23,28,5),7.0), X.eq(E.v(23,28,5),8.0), X.eq(E.v(23,28,5),9.0)))) else '—')

def f249(E, r, c):
    # Ô F124
    # = IF(AND((DUNG_HY_KY!$D$7=0),OR(TRUNG_GIAN!$E$28=6,TRUNG_GIAN!$E$28=7,TRUNG_GIAN!$E$28=8,TRUNG_GIAN!$E$28=9)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.eq(E.v(23,28,5),6.0), X.eq(E.v(23,28,5),7.0), X.eq(E.v(23,28,5),8.0), X.eq(E.v(23,28,5),9.0)))) else '—')

def f250(E, r, c):
    # Ô K124
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&" → Quy tắc (P2 tr.575): Thân nhược và nhật can ở Bệnh, Tử, Mộ, Tuy
    #   ệt tại chi ngày → "&IF(J124="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J124&") ⇒ "&G124&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),' → Quy tắc (P2 tr.575): Thân nhược và nhật can ở Bệnh, Tử, Mộ, Tuyệt tại chi ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f251(E, r, c):
    # Ô E125, F125, E155, F155
    # = IF((COUNTIFS(SAO!$R$5:$R$20,1)>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '—')

def f252(E, r, c):
    # Ô K125
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""
    #   ))&" → Quy tắc (P2 tr.575): Có kình dương ở chi → "&IF(J125="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J125&") ⇒ "&G125&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.575): Có kình dương ở chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f253(E, r, c):
    # Ô E126, F126
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f254(E, r, c):
    # Ô K126
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thầ
    #   n: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.576): Có thực thần → "&IF(J126="","chưa đ
    #   ủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J126&") ⇒ "&G126&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.576): Có thực thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f255(E, r, c):
    # Ô E127, F127
    # = IF((PP1_DINH_TINH!$F$96>=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,96,6),1.0)) else '—')

def f256(E, r, c):
    # Ô K127
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&" → Quy tắc (P2 tr.576): Có tỉ kiếp (ngoài nhật can)
    #    → "&IF(J127="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J127&") ⇒ "&G127&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),' → Quy tắc (P2 tr.576): Có tỉ kiếp (ngoài nhật can) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f257(E, r, c):
    # Ô E128
    # = IF(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f258(E, r, c):
    # Ô F128
    # = IF(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f259(E, r, c):
    # Ô K128
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.576): Thân nhược, tài ≥ P_NHIEU chữ → "&IF(J128="","chưa đủ điều kiện ở c
    #   ả hai phương pháp – không áp dụng.","ĐỦ ("&J128&") ⇒ "&G128&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.576): Thân nhược, tài ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f260(E, r, c):
    # Ô E129
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))), X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0))) else '—')

def f261(E, r, c):
    # Ô F129
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))), X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0))) else '—')

def f262(E, r, c):
    # Ô K129
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.576): Thân
    #    nhược, Thủy ≥ P_NHIEU, Mộc ≤ 1 chữ → "&IF(J129="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J129&") ⇒ "&G129&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.576): Thân nhược, Thủy ≥ P_NHIEU, Mộc ≤ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f263(E, r, c):
    # Ô K130
    # = "Dữ kiện: "&"Số chữ Ấn lộ rõ: "&PP1_DINH_TINH!$D$100&" → Quy tắc (P2 tr.577): Ấn lộ ở thiên can → "&IF(J130="","chưa đủ điều kiện ở cả hai phương pháp
    #    – không áp dụng.","ĐỦ ("&J130&") ⇒ "&G130&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn lộ rõ: '),E.v(9,100,4)),' → Quy tắc (P2 tr.577): Ấn lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f264(E, r, c):
    # Ô E131, F131
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),1.0)) else '—')

def f265(E, r, c):
    # Ô K131
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$
    #   F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Thất sát chỉ 1 chữ → "&IF(J131="","chưa đủ điều kiện ở cả hai phương phá
    #   p – không áp dụng.","ĐỦ ("&J131&") ⇒ "&G131&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.577): Thất sát chỉ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f266(E, r, c):
    # Ô E132, F132
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),2.0)) else '—')

def f267(E, r, c):
    # Ô K132
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$
    #   F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Thất sát ≥ 2 chữ → "&IF(J132="","chưa đủ điều kiện ở cả hai phương pháp 
    #   – không áp dụng.","ĐỦ ("&J132&") ⇒ "&G132&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.577): Thất sát ≥ 2 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f268(E, r, c):
    # Ô E133, F133
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f269(E, r, c):
    # Ô K133
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5
    #   :$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Có thất sát và kình dương ở chi → "&IF(J133="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J133&") ⇒ "&G133&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.577): Có thất sát và kình dương ở chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f270(E, r, c):
    # Ô E138, F138
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1,SAO!$O$5:$O$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),0.0)) else '—')

def f271(E, r, c):
    # Ô K138
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$2
    #   0,1)>0,"Chính ấn thiên/nguyệt đức: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)&" chữ; ","")&IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)>0,"Thiên ấn thiên/nguyệt đức: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.573): Có ấn đóng cùng trụ với thiên/nguyệt đức → "&IF(J138="","chưa đủ 
    #   điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J138&") ⇒ "&G138&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn thiên/nguyệt đức: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên/nguyệt đức: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.573): Có ấn đóng cùng trụ với thiên/nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f272(E, r, c):
    # Ô E139, F139
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(TRUNG_GIAN!$E$5=1)),OR(AND(TRUNG_GIAN!$D$5=6,PP2_DINH_LUONG!$D$43=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_
    #   GIAN!$F$5=6,PP2_DINH_LUONG!$E$43=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,5,5),1.0)), X.OR(X.AND(X.eq(E.v(23,5,4),6.0), X.eq(E.v(10,43,4),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0), X.eq(E.v(10,43,5),1.0))))) else '—')

def f273(E, r, c):
    # Ô K139
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HA
    #   NH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_
    #   GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&"
    #   )"&" → Quy tắc (P2 tr.573): Nam, nhật can Ất, có Canh kề bên thành ngũ hợp → "&IF(J139="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J139&") ⇒ "&G139&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.573): Nam, nhật can Ất, có Canh kề bên thành ngũ hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f274(E, r, c):
    # Ô E140
    # = IF(AND(AND((TRUNG_GIAN!$E$5=0),OR(AND(TRUE(),TRUNG_GIAN!$C$5=5),AND(TRUE(),TRUNG_GIAN!$D$5=5),AND(TRUE(),TRUNG_GIAN!$E$5=5),AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),TRUNG_GIAN!$F$5=5))),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,5,5),0.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),5.0)))), X.eq(E.v(8,7,3),1.0))) else '—')

def f275(E, r, c):
    # Ô F140
    # = IF(AND(AND((TRUNG_GIAN!$E$5=0),OR(AND(TRUE(),TRUNG_GIAN!$C$5=5),AND(TRUE(),TRUNG_GIAN!$D$5=5),AND(TRUE(),TRUNG_GIAN!$E$5=5),AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),TRUNG_GIAN!$F$5=5))),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,5,5),0.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),5.0)))), X.eq(E.v(8,7,4),1.0))) else '—')

def f276(E, r, c):
    # Ô K140
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUN
    #   G_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1
    #   )&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HA
    #   NH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.573): Nhật can Giáp, có Kỷ, thân vượng → "&IF(J140="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J140&") ⇒ "&G140&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.573): Nhật can Giáp, có Kỷ, thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f277(E, r, c):
    # Ô E141, F141
    # = IF(AND(AND(AND((TRUNG_GIAN!$E$5=1),OR(AND(TRUE(),TRUNG_GIAN!$C$5=6),AND(TRUE(),TRUNG_GIAN!$D$5=6),AND(TRUE(),TRUNG_GIAN!$E$5=6),AND(NOT(NHAP_LIEU!$C$9
    #   ="Không rõ giờ"),TRUNG_GIAN!$F$5=6))),OR(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$D$6=9)),NOT(OR(AND(TRUE(),TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$D$6)=6),AND(TR
    #   UE(),TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$F$6)=6),AND(TRUE(),T
    #   RUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6),AND(TRUE(),NOT(NHA
    #   P_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(23,5,5),1.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0)))), X.OR(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,4),9.0))), X.NOT(X.OR(X.AND(X.TRUE(), X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0)))))) else '—')

def f278(E, r, c):
    # Ô K141
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.573): Nhật can Ất, có Canh,
    #    sinh tháng Thân/Dậu (Canh vượng), không có chi xung → "&IF(J141="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J141&") ⇒ "&G141&
    #   ".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.573): Nhật can Ất, có Canh, sinh tháng Thân/Dậu (Canh vượng), không có chi xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f279(E, r, c):
    # Ô E142, F142
    # = IF(AND(AND((TRUNG_GIAN!$E$5=0),OR(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$D$6=4)),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D
    #   $6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!
    #   $D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=3)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,5,5),0.0), X.OR(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,4),4.0))), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0)))))) else '—')

def f280(E, r, c):
    # Ô K142
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.573): Nhật can Giáp, sinh m
    #   ùa xuân (tháng Dần, Mão, Thìn), có Dần hoặc Mão → "&IF(J142="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J142&") ⇒ "&G142&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.573): Nhật can Giáp, sinh mùa xuân (tháng Dần, Mão, Thìn), có Dần hoặc Mão → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f281(E, r, c):
    # Ô E143
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$C$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,3),0.0))) else '—')

def f282(E, r, c):
    # Ô F143
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$D$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,4),0.0))) else '—')

def f283(E, r, c):
    # Ô K143
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.573): Nữ, thâ
    #   n nhược → "&IF(J143="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J143&") ⇒ "&G143&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.573): Nữ, thân nhược → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f284(E, r, c):
    # Ô E144
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,3),1.0))) else '—')

def f285(E, r, c):
    # Ô F144
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,4),1.0))) else '—')

def f286(E, r, c):
    # Ô K144
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.575): Nữ, thâ
    #   n vượng → "&IF(J144="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J144&") ⇒ "&G144&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.575): Nữ, thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f287(E, r, c):
    # Ô E145, F145
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Th
    #   ực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f288(E, r, c):
    # Ô K145
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$J$5:$J$20,1)>0,"Thực thần thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$
    #   20,1)>0,"Thực thần bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$
    #   J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.573): Thực thần (can lộ hoặc bản khí) vượng tại lệnh → "&IF(J145="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J145&") ⇒ "&G145&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.573): Thực thần (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f289(E, r, c):
    # Ô E146, F146
    # = IF((THAN_SAT!$H$104="Có"),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(12,104,8),'Có')) else '—')

def f290(E, r, c):
    # Ô K146
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.574): Có khôi canh → "&IF(J146="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J1
    #   46&") ⇒ "&G146&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.574): Có khôi canh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f291(E, r, c):
    # Ô E147, F147
    # = IF(AND((TRUNG_GIAN!$E$5=2),OR(AND(TRUE(),TRUNG_GIAN!$C$5=7),AND(TRUE(),TRUNG_GIAN!$D$5=7),AND(TRUE(),TRUNG_GIAN!$E$5=7),AND(NOT(NHAP_LIEU!$C$9="Không 
    #   rõ giờ"),TRUNG_GIAN!$F$5=7))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),2.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0))))) else '—')

def f292(E, r, c):
    # Ô K147
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.574): Nhật can Bính, có Tân
    #    → "&IF(J147="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J147&") ⇒ "&G147&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.574): Nhật can Bính, có Tân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f293(E, r, c):
    # Ô E148, F148
    # = IF(OR(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$93)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,93,6)),0.0)))) else '—')

def f294(E, r, c):
    # Ô K148
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.574): Tướng tinh ở chi giờ → "&IF(J148="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","
    #   ĐỦ ("&J148&") ⇒ "&G148&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.574): Tướng tinh ở chi giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f295(E, r, c):
    # Ô E149, F149
    # = IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),2.0)) else '—')

def f296(E, r, c):
    # Ô K149
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&" → Quy tắc (P2 tr.574): Tứ trụ trung hòa (PP1) → "&IF(J149="","chư
    #   a đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J149&") ⇒ "&G149&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),' → Quy tắc (P2 tr.574): Tứ trụ trung hòa (PP1) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f297(E, r, c):
    # Ô E150, F150
    # = IF(AND((TRUNG_GIAN!$E$7=3),OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),3.0), X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)))) else '—')

def f298(E, r, c):
    # Ô K150
    # = "Dữ kiện: "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&
    #   "Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr
    #   .574): Nhật chủ Kim, sinh mùa đông (Hợi, Tí, Sửu) → "&IF(J150="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J150&") ⇒ "&G150&"."
    #   )
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.574): Nhật chủ Kim, sinh mùa đông (Hợi, Tí, Sửu) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f299(E, r, c):
    # Ô E151, F151
    # = IF(AND((TRUNG_GIAN!$E$5=3),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),3.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f300(E, r, c):
    # Ô K151
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUN
    #   TIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.574): Nhật can Đinh, có thương quan → "&IF(J151=
    #   "","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J151&") ⇒ "&G151&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.574): Nhật can Đinh, có thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f301(E, r, c):
    # Ô E152, F152
    # = IF(OR(AND(MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_GIAN!$D$7,MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$E$7,MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$F$7,NOT(NHAP_LIEU!$C$
    #   9="Không rõ giờ")),AND(MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$C$7,MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$D$7,MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG_GIAN!$E$7,NOT(
    #   NHAP_LIEU!$C$9="Không rõ giờ"))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,7,6)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))), X.AND(X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,3)), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))))) else '—')

def f302(E, r, c):
    # Ô K152
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.574): Can năm → tháng → ngà
    #   y → giờ sinh nhau liên tiếp (hoặc ngược lại) → "&IF(J152="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J152&") ⇒ "&G152&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.574): Can năm → tháng → ngày → giờ sinh nhau liên tiếp (hoặc ngược lại) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f303(E, r, c):
    # Ô E153, F153
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=8),AND(TRUE(),TRUNG_GIAN!$D$5=8),AND(TRUE(),TRUNG_GIAN!$E$5=8),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=8)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=9),AND(TRUE(),TRUNG_GIAN!$D$5=9),AND(TRUE(),TRUNG_GIAN!$E$5=9),AN
    #   D(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=9)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),8.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),9.0)))))) else '—')

def f304(E, r, c):
    # Ô K153
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUN
    #   G_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1
    #   )&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy 
    #   tắc (P2 tr.575): Thủy ≥ P_NHIEU chữ và Nhâm/Quý lộ ở can → "&IF(J153="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J153&") ⇒ "&G
    #   153&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.575): Thủy ≥ P_NHIEU chữ và Nhâm/Quý lộ ở can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f305(E, r, c):
    # Ô E154, F154
    # = IF(((IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$E$6,12)=1),1,0)+IF(AND(
    #   TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$E$6,12)
    #   =1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),MOD(TRUNG_GIAN!$E$6+TRUNG_GIAN!$F$6,12)=1),1,0))+(IF(AND(TRUE(),PP2_DINH_LUONG!$C$43=1),1,0)+IF(AND(TRUE(),PP2_DINH_LUONG!$D$43=1),1,0)+IF(AND(NO
    #   T(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1),1,0)))>=2,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0))) else 0.0)),X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,3),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else 0.0))),2.0)) else '—')

def f306(E, r, c):
    # Ô K154
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.575): Từ 2 cặp hợp trở lên 
    #   (ngũ hợp can kề nhau + lục hợp chi) → "&IF(J154="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J154&") ⇒ "&G154&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.575): Từ 2 cặp hợp trở lên (ngũ hợp can kề nhau + lục hợp chi) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f307(E, r, c):
    # Ô K155
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""
    #   ))&" → Quy tắc (P2 tr.575-576): Có kình dương ở chi → "&IF(J155="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J155&") ⇒ "&G155&"
    #   .")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.575-576): Có kình dương ở chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f308(E, r, c):
    # Ô E156, F156
    # = IF(AND((TRUNG_GIAN!$E$7=2),OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!$D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(NOT(NHAP_LIEU!$C$9="Không 
    #   rõ giờ"),TRUNG_GIAN!$F$6=3))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),2.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0))))) else '—')

def f309(E, r, c):
    # Ô K156
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.575): Nhật chủ Thổ, có chi 
    #   Mão → "&IF(J156="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J156&") ⇒ "&G156&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.575): Nhật chủ Thổ, có chi Mão → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f310(E, r, c):
    # Ô E157, F157
    # = IF(AND((TRUNG_GIAN!$E$7=3),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),3.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f311(E, r, c):
    # Ô K157
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(
    #   HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.575): Nhật chủ Kim, Hỏa ≥ P_NHIEU chữ → "&IF(J157="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J157&") ⇒ "&G157&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.575): Nhật chủ Kim, Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f312(E, r, c):
    # Ô E158, F158
    # = IF(OR(AND(TRUE(),OR(AND(TRUNG_GIAN!$C$7=3,TRUNG_GIAN!$D$7=0),AND(TRUNG_GIAN!$C$7=0,TRUNG_GIAN!$D$7=3))),AND(TRUE(),OR(AND(TRUNG_GIAN!$D$7=3,TRUNG_GIAN
    #   !$E$7=0),AND(TRUNG_GIAN!$D$7=0,TRUNG_GIAN!$E$7=3))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$7=3,TRUNG_GIAN!$F$7=0),AND(TRUNG_GIAN!
    #   $E$7=0,TRUNG_GIAN!$F$7=3)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,7,3),3.0), X.eq(E.v(23,7,4),0.0)), X.AND(X.eq(E.v(23,7,3),0.0), X.eq(E.v(23,7,4),3.0)))), X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,7,4),3.0), X.eq(E.v(23,7,5),0.0)), X.AND(X.eq(E.v(23,7,4),0.0), X.eq(E.v(23,7,5),3.0)))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,7,5),3.0), X.eq(E.v(23,7,6),0.0)), X.AND(X.eq(E.v(23,7,5),0.0), X.eq(E.v(23,7,6),3.0)))))) else '—')

def f313(E, r, c):
    # Ô K158
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.575, 577): Can Kim và can M
    #   ộc đứng kề nhau (xen kẽ, khắc nhau) → "&IF(J158="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J158&") ⇒ "&G158&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.575, 577): Can Kim và can Mộc đứng kề nhau (xen kẽ, khắc nhau) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f314(E, r, c):
    # Ô E159, F159
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0))) else '—')

def f315(E, r, c):
    # Ô K159
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Chính ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$
    #   L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thiên ấn bị khắc/hợp mất: "&COUNTIFS
    #   (SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.575): Nữ, có ấn bị khắc → "&IF(J159=""
    #   ,"chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J159&") ⇒ "&G159&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.575): Nữ, có ấn bị khắc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f316(E, r, c):
    # Ô E160, F160
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f317(E, r, c):
    # Ô K160
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.575): Có
    #    thương quan và có chính quan → "&IF(J160="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J160&") ⇒ "&G160&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.575): Có thương quan và có chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f318(E, r, c):
    # Ô E161, F161
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1))>0)),OR((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>=P_NHIEU),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>=P_NHIEU))),"ĐỦ
    #   ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.OR(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))))) else '—')

def f319(E, r, c):
    # Ô K161
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.575): Có th
    #   ất sát và thiên ấn, một trong hai ≥ P_NHIEU chữ → "&IF(J161="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J161&") ⇒ "&G161&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.575): Có thất sát và thiên ấn, một trong hai ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f320(E, r, c):
    # Ô E162, F162
    # = IF(OR(AND(TRUE(),LEN(THAN_SAT!$C$92)>0,THAN_SAT!$C$13=1),AND(TRUE(),LEN(THAN_SAT!$D$92)>0,THAN_SAT!$D$13=1),AND(TRUE(),LEN(THAN_SAT!$E$92)>0,THAN_SAT!
    #   $E$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$92)>0,THAN_SAT!$F$13=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,3)),0.0), X.eq(E.v(12,13,3),1.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,4)),0.0), X.eq(E.v(12,13,4),1.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0), X.eq(E.v(12,13,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,92,6)),0.0), X.eq(E.v(12,13,6),1.0)))) else '—')

def f321(E, r, c):
    # Ô K162
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.575): Trạch mã ở chi rơi vào tuần không → "&IF(J162="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J162&") ⇒ "&G162&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.575): Trạch mã ở chi rơi vào tuần không → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f322(E, r, c):
    # Ô E163, F163
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1))>0)),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f323(E, r, c):
    # Ô K163
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương
    #   : "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.575): Có thiên ấn, kiếp tài và kình dương → "&IF(J163="","chưa đủ điều kiện ở cả hai p
    #   hương pháp – không áp dụng.","ĐỦ ("&J163&") ⇒ "&G163&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.575): Có thiên ấn, kiếp tài và kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f324(E, r, c):
    # Ô E164
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNT
    #   IFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f325(E, r, c):
    # Ô F164
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNT
    #   IFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f326(E, r, c):
    # Ô K164
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$
    #   J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO
    #   !$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.575): Thân nhược, tài ≥ P_NHIEU chữ và tài có lực → "&IF(J164="","chưa đủ điều kiện ở cả hai phương ph
    #   áp – không áp dụng.","ĐỦ ("&J164&") ⇒ "&G164&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.575): Thân nhược, tài ≥ P_NHIEU chữ và tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f327(E, r, c):
    # Ô E165, F165
    # = IF(OR(AND(TRUE(),LEN(THAN_SAT!$C$89)>0,THAN_SAT!$C$13=1),AND(TRUE(),LEN(THAN_SAT!$D$89)>0,THAN_SAT!$D$13=1),AND(TRUE(),LEN(THAN_SAT!$E$89)>0,THAN_SAT!
    #   $E$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$89)>0,THAN_SAT!$F$13=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,89,3)),0.0), X.eq(E.v(12,13,3),1.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,89,4)),0.0), X.eq(E.v(12,13,4),1.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,89,5)),0.0), X.eq(E.v(12,13,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,89,6)),0.0), X.eq(E.v(12,13,6),1.0)))) else '—')

def f328(E, r, c):
    # Ô K165
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.576): Hoa cái ở chi rơi vào tuần không → "&IF(J165="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J165&") ⇒ "&G165&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.576): Hoa cái ở chi rơi vào tuần không → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f329(E, r, c):
    # Ô E166
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f330(E, r, c):
    # Ô F166
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f331(E, r, c):
    # Ô K166
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.576): Thân
    #    nhược, Thủy và Hỏa đều ≥ P_NHIEU chữ → "&IF(J166="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J166&") ⇒ "&G166&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.576): Thân nhược, Thủy và Hỏa đều ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f332(E, r, c):
    # Ô E167, F167
    # = IF(MOD(TRUNG_GIAN!$C$7+2,5)=TRUNG_GIAN!$E$7,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(23,7,5))) else '—')

def f333(E, r, c):
    # Ô K167
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.5
    #   76): Hành can năm khắc hành can ngày → "&IF(J167="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J167&") ⇒ "&G167&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.576): Hành can năm khắc hành can ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f334(E, r, c):
    # Ô E168, F168
    # = IF(MOD(TRUNG_GIAN!$E$7+2,5)=TRUNG_GIAN!$C$7,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,7,3))) else '—')

def f335(E, r, c):
    # Ô K168
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.5
    #   76): Hành can ngày khắc hành can năm → "&IF(J168="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J168&") ⇒ "&G168&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.576): Hành can ngày khắc hành can năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f336(E, r, c):
    # Ô E169, F169
    # = IF(AND(AND(AND((NHAP_LIEU!$C$22="Nam"),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=2),AND(TRUE(),TRUNG_GIAN!$D$5=2),A
    #   ND(TRUE(),TRUNG_GIAN!$E$5=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=3),AND(TRUE(),TRUNG_GIAN!$D$5=3)
    #   ,AND(TRUE(),TRUNG_GIAN!$E$5=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=3)))),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2))), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0))))), X.eq(X.INDEX(E.rg(8,24,3,28,3), 5.0),0.0))) else '—')

def f337(E, r, c):
    # Ô K169
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_
    #   TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.576): Nam, Hỏa ≥ P_NHIEU chữ, có Bính/Đinh lộ, không có Thủy chế → "&IF(J169="","chưa đủ điều kiện ở cả 
    #   hai phương pháp – không áp dụng.","ĐỦ ("&J169&") ⇒ "&G169&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.576): Nam, Hỏa ≥ P_NHIEU chữ, có Bính/Đinh lộ, không có Thủy chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f338(E, r, c):
    # Ô E170, F170
    # = IF(OR((PP2_DINH_LUONG!$F$91/PP2_DINH_LUONG!$C$100)>=0.5,(PP2_DINH_LUONG!$F$92/PP2_DINH_LUONG!$C$100)>=0.5,(PP2_DINH_LUONG!$F$93/PP2_DINH_LUONG!$C$100)
    #   >=0.5,(PP2_DINH_LUONG!$F$94/PP2_DINH_LUONG!$C$100)>=0.5,(PP2_DINH_LUONG!$F$95/PP2_DINH_LUONG!$C$100)>=0.5),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ge(X.div(E.v(10,91,6),E.v(10,100,3)),0.5), X.ge(X.div(E.v(10,92,6),E.v(10,100,3)),0.5), X.ge(X.div(E.v(10,93,6),E.v(10,100,3)),0.5), X.ge(X.div(E.v(10,94,6),E.v(10,100,3)),0.5), X.ge(X.div(E.v(10,95,6),E.v(10,100,3)),0.5))) else '—')

def f339(E, r, c):
    # Ô K170
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.576): Một hành chiếm ≥ 50% tổng độ PP2 → "&IF(J170="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J170&") ⇒ "&G170&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.576): Một hành chiếm ≥ 50% tổng độ PP2 → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f340(E, r, c):
    # Ô E171, F171
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1,SAO!$M$5:$M$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)) else '—')

def f341(E, r, c):
    # Ô K171
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$2
    #   0,1)>0,"Chính ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$
    #   F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.576): Ấn ở chi bị xung → "&IF(J171="","chưa đủ điều kiện ở cả hai phương pháp – không áp 
    #   dụng.","ĐỦ ("&J171&") ⇒ "&G171&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.576): Ấn ở chi bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f342(E, r, c):
    # Ô E172, F172
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$2
    #   0,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))) else '—')

def f343(E, r, c):
    # Ô K172
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M
    #   $20,1)>0,"Chính quan bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thất sát bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Chính ấn b
    #   ị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5
    #   :$M$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Cả quan sát và ấn đều ở chi bị xung → "&IF(J172="","chưa đủ điều kiện ở cả hai phương pháp – không áp
    #    dụng.","ĐỦ ("&J172&") ⇒ "&G172&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.577): Cả quan sát và ấn đều ở chi bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f344(E, r, c):
    # Ô E173, F173
    # = IF(AND(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí"))>0)),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí'),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f345(E, r, c):
    # Ô K173
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí")>0,"Thất sát bản khí: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí")&" chữ; ","")&IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",
    #   ""))&" → Quy tắc (P2 tr.576): Thất sát lộ ở can hoặc là bản khí chi, và có kình dương → "&IF(J173="","chưa đủ điều kiện ở cả hai phương pháp – không á
    #   p dụng.","ĐỦ ("&J173&") ⇒ "&G173&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thất sát bản khí: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí'),0.0)) else '')),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.576): Thất sát lộ ở can hoặc là bản khí chi, và có kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f346(E, r, c):
    # Ô E174, F174
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí'),0.0))) else '—')

def f347(E, r, c):
    # Ô K174
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí")>0,"Thất sát bản khí: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí")&" chữ; ",""))&" → Quy tắc (P2 tr.576): Thất sát lộ ở can hoặc là bản khí chi → "&IF(J174="","chư
    #   a đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J174&") ⇒ "&G174&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thất sát bản khí: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí'),0.0)) else '')))),' → Quy tắc (P2 tr.576): Thất sát lộ ở can hoặc là bản khí chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f348(E, r, c):
    # Ô E175, F175
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f349(E, r, c):
    # Ô K175
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính ấn:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.576): Có chính ấn và
    #    thực thần → "&IF(J175="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J175&") ⇒ "&G175&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.576): Có chính ấn và thực thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f350(E, r, c):
    # Ô E176, F176
    # = IF(AND(MOD(TRUNG_GIAN!$E$5,2)=0,OR(AND(TRUNG_GIAN!$D$12="Chính tài",PP2_DINH_LUONG!$D$43=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="C
    #   hính tài",PP2_DINH_LUONG!$E$43=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0), X.OR(X.AND(X.eq(E.v(23,12,4),'Chính tài'), X.eq(E.v(10,43,4),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Chính tài'), X.eq(E.v(10,43,5),1.0))))) else '—')

def f351(E, r, c):
    # Ô K176
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.576): Nhật can dương, can tháng/giờ là chính tài ngũ hợp với nhật can → "&IF(J176="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J176&") ⇒ "&G176&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.576): Nhật can dương, can tháng/giờ là chính tài ngũ hợp với nhật can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f352(E, r, c):
    # Ô E177, F177
    # = IF(OR(AND(TRUE(),SAO!$R$5=1,OR(AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$D$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$D$6+1
    #   )+1)="Thương quan"),AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$E$6+1)+1)="Thươn
    #   g quan"),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$F$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!
    #   $F$6+1)+1)="Thương quan"))),AND(TRUE(),SAO!$R$9=1,OR(AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$C$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$
    #   D$29,TRUNG_GIAN!$C$6+1)+1)="Thương quan"),AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_
    #   GIAN!$E$6+1)+1)="Thương quan"),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$
    #   D$18:$D$29,TRUNG_GIAN!$F$6+1)+1)="Thương quan"))),AND(TRUE(),SAO!$R$13=1,OR(AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$C$6)=6,INDEX(M_TT,TRUNG_GIAN!$E
    #   $5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$C$6+1)+1)="Thương quan"),AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$D$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(
    #   CSDL!$D$18:$D$29,TRUNG_GIAN!$D$6+1)+1)="Thương quan"),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6,INDEX(M_TT,TRUNG_G
    #   IAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$F$6+1)+1)="Thương quan"))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$R$17=1,OR(AND(TRUE(),ABS(TRUNG_GI
    #   AN!$F$6-TRUNG_GIAN!$C$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$C$6+1)+1)="Thương quan"),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRU
    #   NG_GIAN!$D$6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$D$6+1)+1)="Thương quan"),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$E$
    #   6)=6,INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$E$6+1)+1)="Thương quan")))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,18),1.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,4),1.0)),1.0)),'Thương quan')), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,5),1.0)),1.0)),'Thương quan')), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,6),1.0)),1.0)),'Thương quan')))), X.AND(X.TRUE(), X.eq(E.v(13,9,18),1.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,3))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,3),1.0)),1.0)),'Thương quan')), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,5),1.0)),1.0)),'Thương quan')), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,6),1.0)),1.0)),'Thương quan')))), X.AND(X.TRUE(), X.eq(E.v(13,13,18),1.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,3))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,3),1.0)),1.0)),'Thương quan')), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,4))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,4),1.0)),1.0)),'Thương quan')), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,6),1.0)),1.0)),'Thương quan')))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,18),1.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,3))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,3),1.0)),1.0)),'Thương quan')), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,4))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,4),1.0)),1.0)),'Thương quan')), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,5))),6.0), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,5),1.0)),1.0)),'Thương quan')))))) else '—')

def f353(E, r, c):
    # Ô K177
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.576): Chi kình dương bị chi
    #    có bản khí thương quan xung → "&IF(J177="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J177&") ⇒ "&G177&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.576): Chi kình dương bị chi có bản khí thương quan xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f354(E, r, c):
    # Ô E178, F178
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)))) else '—')

def f355(E, r, c):
    # Ô K178
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.576): Hỏa ≥ P_NHIEU và Thổ ≥ P_NHIEU → "&IF(J178="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J178&") ⇒ "&G178&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.576): Hỏa ≥ P_NHIEU và Thổ ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f356(E, r, c):
    # Ô E179, F179
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,98,6),1.0))) else '—')

def f357(E, r, c):
    # Ô K179
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Có kiếp tài và có tài → "&IF(J179="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("
    #   &J179&") ⇒ "&G179&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.577): Có kiếp tài và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f358(E, r, c):
    # Ô E180, F180
    # = IF((IF(AND(TRUE(),OR(MOD(TRUNG_GIAN!$C$7+2,5)=TRUNG_GIAN!$D$7,MOD(TRUNG_GIAN!$D$7+2,5)=TRUNG_GIAN!$C$7)),1,0)+IF(AND(TRUE(),OR(MOD(TRUNG_GIAN!$D$7+2,5
    #   )=TRUNG_GIAN!$E$7,MOD(TRUNG_GIAN!$E$7+2,5)=TRUNG_GIAN!$D$7)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(MOD(TRUNG_GIAN!$E$7+2,5)=TRUNG_GIAN!$F$
    #   7,MOD(TRUNG_GIAN!$F$7+2,5)=TRUNG_GIAN!$E$7)),1,0))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(23,7,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,7,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,7,6)), X.eq(X.MOD(X.add(E.v(23,7,6),2.0), 5.0),E.v(23,7,5))))) else 0.0)),3.0)) else '—')

def f359(E, r, c):
    # Ô K180
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.577): Cả 3 cặp can kề nhau 
    #   (năm–tháng, tháng–ngày, ngày–giờ) đều khắc nhau → "&IF(J180="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J180&") ⇒ "&G180&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.577): Cả 3 cặp can kề nhau (năm–tháng, tháng–ngày, ngày–giờ) đều khắc nhau → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f360(E, r, c):
    # Ô E181, F181
    # = IF(AND((IF(AND(TRUE(),PP2_DINH_LUONG!$C$43=1),1,0)+IF(AND(TRUE(),PP2_DINH_LUONG!$D$43=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG
    #   !$E$43=1),1,0))>=1,OR(AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$D$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$D$6=0),AND(TRUNG_GIAN!$C$6=2,TRU
    #   NG_GIAN!$D$6=5),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=5),AND(TRUNG_G
    #   IAN!$C$6=2,TRUNG_GIAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$D$
    #   6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$D$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=7),AND(TRUNG_GIAN!$C$6=
    #   7,TRUNG_GIAN!$D$6=1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$D$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),TR
    #   UE(),OR(AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$C$6
    #   =5,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$E$6=8),AND(T
    #   RUNG_GIAN!$C$6=8,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_G
    #   IAN!$E$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN
    #   !$C$6=TRUNG_GIAN!$E$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR
    #   (AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$C$6=5,TRUN
    #   G_GIAN!$F$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GI
    #   AN!$C$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$F$
    #   6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$C$6=T
    #   RUNG_GIAN!$F$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$
    #   E$6=3),AND(TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$D$6=
    #   5,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=2),AND(TR
    #   UNG_GIAN!$D$6=1,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GI
    #   AN!$E$6=10),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$E$6,OR(TRUNG_GIAN!$D$6=4
    #   ,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$F$6=3),
    #   AND(TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$D$6=5,TRUNG
    #   _GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIA
    #   N!$D$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6
    #   =10),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$D$6=4,TRUNG_
    #   GIAN!$D$6=6,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=0,TRUNG_GIAN!$F$6=3),AND(TRU
    #   NG_GIAN!$E$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$
    #   F$6=8),AND(TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$E$6=
    #   1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,TRUNG_GIAN!$F$6=10),AN
    #   D(TRUNG_GIAN!$E$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$E$6=4,TRUNG_GIAN!$E
    #   $6=6,TRUNG_GIAN!$E$6=9,TRUNG_GIAN!$E$6=11)))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,3),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else 0.0)),1.0), X.OR(X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,4),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,4),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,4)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,5),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,6,5),9.0), X.eq(E.v(23,6,5),11.0)))))))) else '—')

def f361(E, r, c):
    # Ô K181
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.577): Có ngũ hợp can và có 
    #   chi hình → "&IF(J181="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J181&") ⇒ "&G181&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.577): Có ngũ hợp can và có chi hình → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f362(E, r, c):
    # Ô E182, F182
    # = IF(OR(AND(TRUE(),LEN(THAN_SAT!$C$92)>0,OR(AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(NO
    #   T(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$F$6)=6))),AND(TRUE(),LEN(THAN_SAT!$D$92)>0,OR(AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GI
    #   AN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6))),AND(TR
    #   UE(),LEN(THAN_SAT!$E$92)>0,OR(AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$D$6)=6),AND(NOT(NHAP_LIEU!
    #   $C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$92)>0,OR(AND(TRUE(),ABS(TRUNG_GI
    #   AN!$F$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$E$6)=6)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,3)),0.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)))), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,4)),0.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)))), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,4))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0)))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,92,6)),0.0), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,5))),6.0)))))) else '—')

def f363(E, r, c):
    # Ô K182
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.577): Trạch mã ở chi bị xun
    #   g → "&IF(J182="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J182&") ⇒ "&G182&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.577): Trạch mã ở chi bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f364(E, r, c):
    # Ô E183, F183
    # = IF(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=0),AND(TRUE(),TRUNG_GIAN!$D$6=0),AND(TRUE(),TRUNG_GIAN!$E$6=0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!
    #   $F$6=0)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!$D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIA
    #   N!$F$6=3))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),0.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0))))) else '—')

def f365(E, r, c):
    # Ô K183
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.577): Có cả Tí và Mão (tươn
    #   g hình) → "&IF(J183="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J183&") ⇒ "&G183&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.577): Có cả Tí và Mão (tương hình) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f366(E, r, c):
    # Ô E184, F184
    # = IF((IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$E$6,12)=1),1,0)+IF(AND(T
    #   RUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$E$6,12)=
    #   1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ gi
    #   ờ"),MOD(TRUNG_GIAN!$E$6+TRUNG_GIAN!$F$6,12)=1),1,0))>=2,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0))) else 0.0)),2.0)) else '—')

def f367(E, r, c):
    # Ô K184
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.577): Lục hợp chi ≥ 2 cặp →
    #    "&IF(J184="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J184&") ⇒ "&G184&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.577): Lục hợp chi ≥ 2 cặp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f368(E, r, c):
    # Ô E185, F185
    # = IF(OR(TRUNG_GIAN!$E$28=7,TRUNG_GIAN!$E$28=9),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,28,5),7.0), X.eq(E.v(23,28,5),9.0))) else '—')

def f369(E, r, c):
    # Ô K185
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.577): Nhật can tại chi ngày ở Tử hoặc Tuyệt → "&IF(J185="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J185&") ⇒ "&G185&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.577): Nhật can tại chi ngày ở Tử hoặc Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f370(E, r, c):
    # Ô E186, F186
    # = IF(AND((THAN_SAT!$H$104="Có"),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(12,104,8),'Có'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f371(E, r, c):
    # Ô K186
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Có khôi canh và có thất sát → "&IF(J18
    #   6="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J186&") ⇒ "&G186&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.577): Có khôi canh và có thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f372(E, r, c):
    # Ô E187, F187
    # = IF(AND((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>=P_NHIEU),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f373(E, r, c):
    # Ô K187
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:
    #   $Y$20,0)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0
    #   ,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.577): Thiên ấn ≥ P_NHIEU và thương quan ≥ P_
    #   NHIEU → "&IF(J187="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J187&") ⇒ "&G187&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else ''),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.577): Thiên ấn ≥ P_NHIEU và thương quan ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f374(E, r, c):
    # Ô C193
    # = (PP2_DINH_LUONG!$F$91/PP2_DINH_LUONG!$C$100)
    return X.div(E.v(10,91,6),E.v(10,100,3))

def f375(E, r, c):
    # Ô D193
    # = IF(C193>=0.4,"Mạnh quá",IF(C193>=0.25,"Thịnh",IF(C193<=0.1,IF(OR(INDEX(M_TS,1,TRUNG_GIAN!$D$6+1)=7,INDEX(M_TS,1,TRUNG_GIAN!$D$6+1)=9),"Tử tuyệt","Suy 
    #   / không đủ"),"Bình thường")))
    return ('Mạnh quá' if X.truthy(X.ge(E.v(20,r,c-1),0.4)) else ('Thịnh' if X.truthy(X.ge(E.v(20,r,c-1),0.25)) else (('Tử tuyệt' if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), 1.0, X.add(E.v(23,6,4),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), 1.0, X.add(E.v(23,6,4),1.0)),9.0))) else 'Suy / không đủ') if X.truthy(X.le(E.v(20,r,c-1),0.1)) else 'Bình thường')))

def f376(E, r, c):
    # Ô E193
    # = "Mộc chủ về nhân, tính thẳng, tình ôn hòa. "&IF(D193="Mạnh quá","Người mộc thịnh thì tầm vóc cao, tay chân dài, phong cách đẹp, khóe miệng tươi, sắc m
    #   ặt trắng xanh, có lòng bác ái thương người, thanh cao khảng khái, chất phác, không giả dối.",IF(D193="Thịnh","Người mộc thịnh thì tầm vóc cao, tay châ
    #   n dài, phong cách đẹp, khóe miệng tươi, sắc mặt trắng xanh, có lòng bác ái thương người, thanh cao khảng khái, chất phác, không giả dối.",IF(D193="Tử 
    #   tuyệt","Người mà mộc khí tử tuyệt thì lông mày không ngay ngắn, cổ dài, yết hầu lồi, da thịt khô, hay lừa dối, bỉ lận.",IF(D193="Suy / không đủ","Ngườ
    #   i mộc suy thì vóc người gầy, tóc thưa, tính cách hẹp hòi, đố kị, bất nhân.","(Hành này ở mức bình thường – không nổi trội, không thiếu.)"))))
    return X.cat('Mộc chủ về nhân, tính thẳng, tình ôn hòa. ',('Người mộc thịnh thì tầm vóc cao, tay chân dài, phong cách đẹp, khóe miệng tươi, sắc mặt trắng xanh, có lòng bác ái thương người, thanh cao khảng khái, chất phác, không giả dối.' if X.truthy(X.eq(E.v(20,r,c-1),'Mạnh quá')) else ('Người mộc thịnh thì tầm vóc cao, tay chân dài, phong cách đẹp, khóe miệng tươi, sắc mặt trắng xanh, có lòng bác ái thương người, thanh cao khảng khái, chất phác, không giả dối.' if X.truthy(X.eq(E.v(20,r,c-1),'Thịnh')) else ('Người mà mộc khí tử tuyệt thì lông mày không ngay ngắn, cổ dài, yết hầu lồi, da thịt khô, hay lừa dối, bỉ lận.' if X.truthy(X.eq(E.v(20,r,c-1),'Tử tuyệt')) else ('Người mộc suy thì vóc người gầy, tóc thưa, tính cách hẹp hòi, đố kị, bất nhân.' if X.truthy(X.eq(E.v(20,r,c-1),'Suy / không đủ')) else '(Hành này ở mức bình thường – không nổi trội, không thiếu.)')))))

def f377(E, r, c):
    # Ô C194
    # = (PP2_DINH_LUONG!$F$92/PP2_DINH_LUONG!$C$100)
    return X.div(E.v(10,92,6),E.v(10,100,3))

def f378(E, r, c):
    # Ô D194, D195, D196, D197
    # = IF(C194>=0.4,"Mạnh quá",IF(C194>=0.25,"Thịnh",IF(C194<=0.1,IF(FALSE(),"Tử tuyệt","Suy / không đủ"),"Bình thường")))
    return ('Mạnh quá' if X.truthy(X.ge(E.v(20,r,c-1),0.4)) else ('Thịnh' if X.truthy(X.ge(E.v(20,r,c-1),0.25)) else (('Tử tuyệt' if X.truthy(X.FALSE()) else 'Suy / không đủ') if X.truthy(X.le(E.v(20,r,c-1),0.1)) else 'Bình thường')))

def f379(E, r, c):
    # Ô E194
    # = "Hỏa chủ về lễ, tính nóng nhưng tình cảm lễ độ. "&IF(D194="Mạnh quá","Người mà hỏa thịnh thì đầu nhỏ chân dài, vóc người trên nhọn dưới nở, mày rậm, t
    #   ai nhỏ, tinh thần hoạt bát, cung kính lễ độ với mọi người, tính nóng gấp.",IF(D194="Thịnh","Người mà hỏa thịnh thì đầu nhỏ chân dài, vóc người trên nh
    #   ọn dưới nở, mày rậm, tai nhỏ, tinh thần hoạt bát, cung kính lễ độ với mọi người, tính nóng gấp.",IF(D194="Tử tuyệt","Người mà hỏa suy thì dáng người g
    #   óc nhọn gầy vàng, nói năng ề à, dối trá, cay độc, làm việc có đầu, không có đuôi.",IF(D194="Suy / không đủ","Người mà hỏa suy thì dáng người góc nhọn 
    #   gầy vàng, nói năng ề à, dối trá, cay độc, làm việc có đầu, không có đuôi.","(Hành này ở mức bình thường – không nổi trội, không thiếu.)"))))
    return X.cat('Hỏa chủ về lễ, tính nóng nhưng tình cảm lễ độ. ',('Người mà hỏa thịnh thì đầu nhỏ chân dài, vóc người trên nhọn dưới nở, mày rậm, tai nhỏ, tinh thần hoạt bát, cung kính lễ độ với mọi người, tính nóng gấp.' if X.truthy(X.eq(E.v(20,r,c-1),'Mạnh quá')) else ('Người mà hỏa thịnh thì đầu nhỏ chân dài, vóc người trên nhọn dưới nở, mày rậm, tai nhỏ, tinh thần hoạt bát, cung kính lễ độ với mọi người, tính nóng gấp.' if X.truthy(X.eq(E.v(20,r,c-1),'Thịnh')) else ('Người mà hỏa suy thì dáng người góc nhọn gầy vàng, nói năng ề à, dối trá, cay độc, làm việc có đầu, không có đuôi.' if X.truthy(X.eq(E.v(20,r,c-1),'Tử tuyệt')) else ('Người mà hỏa suy thì dáng người góc nhọn gầy vàng, nói năng ề à, dối trá, cay độc, làm việc có đầu, không có đuôi.' if X.truthy(X.eq(E.v(20,r,c-1),'Suy / không đủ')) else '(Hành này ở mức bình thường – không nổi trội, không thiếu.)')))))

def f380(E, r, c):
    # Ô C195
    # = (PP2_DINH_LUONG!$F$93/PP2_DINH_LUONG!$C$100)
    return X.div(E.v(10,93,6),E.v(10,100,3))

def f381(E, r, c):
    # Ô E195
    # = "Thổ chủ về tín, tính tình đôn hậu. "&IF(D195="Mạnh quá","Người mà thổ khí mạnh quá thì đầu óc cứng nhắc, hiểu biết chậm, tính tình hướng nội, thích t
    #   rầm tĩnh.",IF(D195="Thịnh","Người thổ thịnh thì thắt lưng tròn, mũi nở, lông mày thanh tú, mắt đẹp, tiếng nói vang và lưu loát, trung hiếu, chân thành
    #   , độ lượng, giữ lời hứa, hành động chắc chắn, kết quả.",IF(D195="Tử tuyệt","Người mà thổ khí không đủ thì sắc mặt ủ rũ, mặt mỏng, mũi ngắn, lòng dạ ác
    #    độc, bất tín, vô tình.",IF(D195="Suy / không đủ","Người mà thổ khí không đủ thì sắc mặt ủ rũ, mặt mỏng, mũi ngắn, lòng dạ ác độc, bất tín, vô tình.",
    #   "(Hành này ở mức bình thường – không nổi trội, không thiếu.)"))))
    return X.cat('Thổ chủ về tín, tính tình đôn hậu. ',('Người mà thổ khí mạnh quá thì đầu óc cứng nhắc, hiểu biết chậm, tính tình hướng nội, thích trầm tĩnh.' if X.truthy(X.eq(E.v(20,r,c-1),'Mạnh quá')) else ('Người thổ thịnh thì thắt lưng tròn, mũi nở, lông mày thanh tú, mắt đẹp, tiếng nói vang và lưu loát, trung hiếu, chân thành, độ lượng, giữ lời hứa, hành động chắc chắn, kết quả.' if X.truthy(X.eq(E.v(20,r,c-1),'Thịnh')) else ('Người mà thổ khí không đủ thì sắc mặt ủ rũ, mặt mỏng, mũi ngắn, lòng dạ ác độc, bất tín, vô tình.' if X.truthy(X.eq(E.v(20,r,c-1),'Tử tuyệt')) else ('Người mà thổ khí không đủ thì sắc mặt ủ rũ, mặt mỏng, mũi ngắn, lòng dạ ác độc, bất tín, vô tình.' if X.truthy(X.eq(E.v(20,r,c-1),'Suy / không đủ')) else '(Hành này ở mức bình thường – không nổi trội, không thiếu.)')))))

def f382(E, r, c):
    # Ô C196
    # = (PP2_DINH_LUONG!$F$94/PP2_DINH_LUONG!$C$100)
    return X.div(E.v(10,94,6),E.v(10,100,3))

def f383(E, r, c):
    # Ô E196
    # = "Kim chủ về nghĩa, tính tình cương trực, mãnh liệt. "&IF(D196="Mạnh quá","Người mà kim mạnh quá thì hữu dũng vô mưu, tham muốn, bất nhân.",IF(D196="Th
    #   ịnh","Người mà kim thịnh là người không béo không gầy, mặt trắng trẻo, lông mày cao, mắt sâu, sức khỏe tốt, tinh thần minh mẫn, tác phong quyết đoán, 
    #   trọng nghĩa khinh tài, tự trọng.",IF(D196="Tử tuyệt","Người mà kim không đủ thì vóc người gầy nhỏ, tư cách vô tình, có khi nham hiểm, ham dâm, háo sát
    #   , bỉ lận, tham lam.",IF(D196="Suy / không đủ","Người mà kim không đủ thì vóc người gầy nhỏ, tư cách vô tình, có khi nham hiểm, ham dâm, háo sát, bỉ lậ
    #   n, tham lam.","(Hành này ở mức bình thường – không nổi trội, không thiếu.)"))))
    return X.cat('Kim chủ về nghĩa, tính tình cương trực, mãnh liệt. ',('Người mà kim mạnh quá thì hữu dũng vô mưu, tham muốn, bất nhân.' if X.truthy(X.eq(E.v(20,r,c-1),'Mạnh quá')) else ('Người mà kim thịnh là người không béo không gầy, mặt trắng trẻo, lông mày cao, mắt sâu, sức khỏe tốt, tinh thần minh mẫn, tác phong quyết đoán, trọng nghĩa khinh tài, tự trọng.' if X.truthy(X.eq(E.v(20,r,c-1),'Thịnh')) else ('Người mà kim không đủ thì vóc người gầy nhỏ, tư cách vô tình, có khi nham hiểm, ham dâm, háo sát, bỉ lận, tham lam.' if X.truthy(X.eq(E.v(20,r,c-1),'Tử tuyệt')) else ('Người mà kim không đủ thì vóc người gầy nhỏ, tư cách vô tình, có khi nham hiểm, ham dâm, háo sát, bỉ lận, tham lam.' if X.truthy(X.eq(E.v(20,r,c-1),'Suy / không đủ')) else '(Hành này ở mức bình thường – không nổi trội, không thiếu.)')))))

def f384(E, r, c):
    # Ô C197
    # = (PP2_DINH_LUONG!$F$95/PP2_DINH_LUONG!$C$100)
    return X.div(E.v(10,95,6),E.v(10,100,3))

def f385(E, r, c):
    # Ô E197
    # = "Thủy chủ về trí, thông minh hiền lành. "&IF(D197="Mạnh quá","Người mà thủy mạnh quá thì hay cãi cọ, tính tình linh tinh.",IF(D197="Thịnh","Người mà t
    #   hủy vượng sắc mặt hơi đen, nói năng nhẹ nhàng, rành rọt, hay lo cho người khác, túc trí đa mưu, học nhanh hơn người.",IF(D197="Tử tuyệt","Người mà thủ
    #   y không đủ thì vóc người thấp bé, tính tình bất thường, nhát gan, vô mưu, hành động không có thứ tự.",IF(D197="Suy / không đủ","Người mà thủy không đủ
    #    thì vóc người thấp bé, tính tình bất thường, nhát gan, vô mưu, hành động không có thứ tự.","(Hành này ở mức bình thường – không nổi trội, không thiếu
    #   .)"))))
    return X.cat('Thủy chủ về trí, thông minh hiền lành. ',('Người mà thủy mạnh quá thì hay cãi cọ, tính tình linh tinh.' if X.truthy(X.eq(E.v(20,r,c-1),'Mạnh quá')) else ('Người mà thủy vượng sắc mặt hơi đen, nói năng nhẹ nhàng, rành rọt, hay lo cho người khác, túc trí đa mưu, học nhanh hơn người.' if X.truthy(X.eq(E.v(20,r,c-1),'Thịnh')) else ('Người mà thủy không đủ thì vóc người thấp bé, tính tình bất thường, nhát gan, vô mưu, hành động không có thứ tự.' if X.truthy(X.eq(E.v(20,r,c-1),'Tử tuyệt')) else ('Người mà thủy không đủ thì vóc người thấp bé, tính tình bất thường, nhát gan, vô mưu, hành động không có thứ tự.' if X.truthy(X.eq(E.v(20,r,c-1),'Suy / không đủ')) else '(Hành này ở mức bình thường – không nổi trội, không thiếu.)')))))

def f386(E, r, c):
    # Ô C198
    # = "Mặt trội nổi (hành thịnh): "&IF(TRIM(IF(OR(D193="Thịnh",D193="Mạnh quá"),"Mộc ("&D193&") ","")&IF(OR(D194="Thịnh",D194="Mạnh quá"),"Hỏa ("&D194&") ",
    #   "")&IF(OR(D195="Thịnh",D195="Mạnh quá"),"Thổ ("&D195&") ","")&IF(OR(D196="Thịnh",D196="Mạnh quá"),"Kim ("&D196&") ","")&IF(OR(D197="Thịnh",D197="Mạnh 
    #   quá"),"Thủy ("&D197&") ",""))="","không có hành nào thịnh hẳn",TRIM(IF(OR(D193="Thịnh",D193="Mạnh quá"),"Mộc ("&D193&") ","")&IF(OR(D194="Thịnh",D194=
    #   "Mạnh quá"),"Hỏa ("&D194&") ","")&IF(OR(D195="Thịnh",D195="Mạnh quá"),"Thổ ("&D195&") ","")&IF(OR(D196="Thịnh",D196="Mạnh quá"),"Kim ("&D196&") ","")&
    #   IF(OR(D197="Thịnh",D197="Mạnh quá"),"Thủy ("&D197&") ","")))&"; mặt chìm ẩn / thiếu (hành suy): "&IF(TRIM(IF(OR(D193="Suy / không đủ",D193="Tử tuyệt")
    #   ,"Mộc ("&D193&") ","")&IF(OR(D194="Suy / không đủ",D194="Tử tuyệt"),"Hỏa ("&D194&") ","")&IF(OR(D195="Suy / không đủ",D195="Tử tuyệt"),"Thổ ("&D195&")
    #    ","")&IF(OR(D196="Suy / không đủ",D196="Tử tuyệt"),"Kim ("&D196&") ","")&IF(OR(D197="Suy / không đủ",D197="Tử tuyệt"),"Thủy ("&D197&") ",""))="","khô
    #   ng có",TRIM(IF(OR(D193="Suy / không đủ",D193="Tử tuyệt"),"Mộc ("&D193&") ","")&IF(OR(D194="Suy / không đủ",D194="Tử tuyệt"),"Hỏa ("&D194&") ","")&IF(O
    #   R(D195="Suy / không đủ",D195="Tử tuyệt"),"Thổ ("&D195&") ","")&IF(OR(D196="Suy / không đủ",D196="Tử tuyệt"),"Kim ("&D196&") ","")&IF(OR(D197="Suy / kh
    #   ông đủ",D197="Tử tuyệt"),"Thủy ("&D197&") ","")))&". Đặc điểm cụ thể xem từng dòng trên (P1 tr.25-26). Hình dáng chỉ là gợi ý của mệnh lý, sách dặn ph
    #   ải kết hợp diện tướng, cốt tướng."
    return X.cat(X.cat(X.cat(X.cat('Mặt trội nổi (hành thịnh): ',('không có hành nào thịnh hẳn' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Mộc (',E.v(20,r-5,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-5,c+1),'Thịnh'), X.eq(E.v(20,r-5,c+1),'Mạnh quá'))) else ''),(X.cat(X.cat('Hỏa (',E.v(20,r-4,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-4,c+1),'Thịnh'), X.eq(E.v(20,r-4,c+1),'Mạnh quá'))) else '')),(X.cat(X.cat('Thổ (',E.v(20,r-3,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-3,c+1),'Thịnh'), X.eq(E.v(20,r-3,c+1),'Mạnh quá'))) else '')),(X.cat(X.cat('Kim (',E.v(20,r-2,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-2,c+1),'Thịnh'), X.eq(E.v(20,r-2,c+1),'Mạnh quá'))) else '')),(X.cat(X.cat('Thủy (',E.v(20,r-1,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-1,c+1),'Thịnh'), X.eq(E.v(20,r-1,c+1),'Mạnh quá'))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Mộc (',E.v(20,r-5,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-5,c+1),'Thịnh'), X.eq(E.v(20,r-5,c+1),'Mạnh quá'))) else ''),(X.cat(X.cat('Hỏa (',E.v(20,r-4,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-4,c+1),'Thịnh'), X.eq(E.v(20,r-4,c+1),'Mạnh quá'))) else '')),(X.cat(X.cat('Thổ (',E.v(20,r-3,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-3,c+1),'Thịnh'), X.eq(E.v(20,r-3,c+1),'Mạnh quá'))) else '')),(X.cat(X.cat('Kim (',E.v(20,r-2,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-2,c+1),'Thịnh'), X.eq(E.v(20,r-2,c+1),'Mạnh quá'))) else '')),(X.cat(X.cat('Thủy (',E.v(20,r-1,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-1,c+1),'Thịnh'), X.eq(E.v(20,r-1,c+1),'Mạnh quá'))) else ''))))),'; mặt chìm ẩn / thiếu (hành suy): '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Mộc (',E.v(20,r-5,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-5,c+1),'Suy / không đủ'), X.eq(E.v(20,r-5,c+1),'Tử tuyệt'))) else ''),(X.cat(X.cat('Hỏa (',E.v(20,r-4,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-4,c+1),'Suy / không đủ'), X.eq(E.v(20,r-4,c+1),'Tử tuyệt'))) else '')),(X.cat(X.cat('Thổ (',E.v(20,r-3,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-3,c+1),'Suy / không đủ'), X.eq(E.v(20,r-3,c+1),'Tử tuyệt'))) else '')),(X.cat(X.cat('Kim (',E.v(20,r-2,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-2,c+1),'Suy / không đủ'), X.eq(E.v(20,r-2,c+1),'Tử tuyệt'))) else '')),(X.cat(X.cat('Thủy (',E.v(20,r-1,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-1,c+1),'Suy / không đủ'), X.eq(E.v(20,r-1,c+1),'Tử tuyệt'))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Mộc (',E.v(20,r-5,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-5,c+1),'Suy / không đủ'), X.eq(E.v(20,r-5,c+1),'Tử tuyệt'))) else ''),(X.cat(X.cat('Hỏa (',E.v(20,r-4,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-4,c+1),'Suy / không đủ'), X.eq(E.v(20,r-4,c+1),'Tử tuyệt'))) else '')),(X.cat(X.cat('Thổ (',E.v(20,r-3,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-3,c+1),'Suy / không đủ'), X.eq(E.v(20,r-3,c+1),'Tử tuyệt'))) else '')),(X.cat(X.cat('Kim (',E.v(20,r-2,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-2,c+1),'Suy / không đủ'), X.eq(E.v(20,r-2,c+1),'Tử tuyệt'))) else '')),(X.cat(X.cat('Thủy (',E.v(20,r-1,c+1)),') ') if X.truthy(X.OR(X.eq(E.v(20,r-1,c+1),'Suy / không đủ'), X.eq(E.v(20,r-1,c+1),'Tử tuyệt'))) else ''))))),'. Đặc điểm cụ thể xem từng dòng trên (P1 tr.25-26). Hình dáng chỉ là gợi ý của mệnh lý, sách dặn phải kết hợp diện tướng, cốt tướng.')

def f387(E, r, c):
    # Ô E203, F203
    # = IF(OR(TRUNG_GIAN!$E$5=0,TRUNG_GIAN!$E$5=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,5,5),0.0), X.eq(E.v(23,5,5),1.0))) else '—')

def f388(E, r, c):
    # Ô K203
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.503): Can ngày Giáp/Ất → "&IF(J203="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J203&") ⇒ "&G203&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.503): Can ngày Giáp/Ất → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f389(E, r, c):
    # Ô E204, F204
    # = IF(OR(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$5=3),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0))) else '—')

def f390(E, r, c):
    # Ô K204
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.503): Can ngày Bính/Đinh → "&IF(J204="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J204&") ⇒ "&G204&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.503): Can ngày Bính/Đinh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f391(E, r, c):
    # Ô E205, F205
    # = IF(OR(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$5=5),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,5,5),5.0))) else '—')

def f392(E, r, c):
    # Ô K205
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.503): Can ngày Mậu/Kỷ → "&IF(J205="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J205&") ⇒ "&G205&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.503): Can ngày Mậu/Kỷ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f393(E, r, c):
    # Ô E206, F206
    # = IF(OR(TRUNG_GIAN!$E$5=6),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,5,5),6.0))) else '—')

def f394(E, r, c):
    # Ô K206
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.503): Can ngày Canh → "&IF(J206="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J206&") ⇒ "&G206&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.503): Can ngày Canh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f395(E, r, c):
    # Ô E207, F207
    # = IF(OR(TRUNG_GIAN!$E$5=7),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,5,5),7.0))) else '—')

def f396(E, r, c):
    # Ô K207
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.503): Can ngày Tân → "&IF(J207="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J207&") ⇒ "&G207&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.503): Can ngày Tân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f397(E, r, c):
    # Ô E208, F208
    # = IF(OR(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$5=9),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,5,5),9.0))) else '—')

def f398(E, r, c):
    # Ô K208
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.503): Can ngày Nhâm/Quý → "&IF(J208="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J208&") ⇒ "&G208&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.503): Can ngày Nhâm/Quý → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f399(E, r, c):
    # Ô E209, F209
    # = IF(OR((IF(AND(TRUE(),OR(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$C$6=1)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$D$6=1)),1,0)+IF(AND(TRUE(),OR(TRUNG_
    #   GIAN!$E$6=2,TRUNG_GIAN!$E$6=1)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=2,TRUNG_GIAN!$F$6=1)),1,0))>=2,(IF(AND(TRUE(),OR(TRU
    #   NG_GIAN!$C$6=10,TRUNG_GIAN!$C$6=11)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$D$6=11)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$E$6=10,TRUNG_GIAN
    #   !$E$6=11)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=10,TRUNG_GIAN!$F$6=11)),1,0))>=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,3),1.0)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,4),1.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,5),1.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),2.0), X.eq(E.v(23,6,6),1.0)))) else 0.0)),2.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,3),11.0)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,4),11.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,5),11.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),10.0), X.eq(E.v(23,6,6),11.0)))) else 0.0)),2.0))) else '—')

def f400(E, r, c):
    # Ô K209
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.503): Có ≥ 2 chi Dần/Sửu ho
    #   ặc ≥ 2 chi Tuất/Hợi → "&IF(J209="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J209&") ⇒ "&G209&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.503): Có ≥ 2 chi Dần/Sửu hoặc ≥ 2 chi Tuất/Hợi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f401(E, r, c):
    # Ô E210, F210
    # = IF(OR((IF(AND(TRUE(),OR(TRUNG_GIAN!$C$6=3)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$D$6=3)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$E$6=3)),1,0)+IF(AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=3)),1,0))>=3,(IF(AND(TRUE(),OR(TRUNG_GIAN!$C$6=4)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$D$6=4)),1,0)+IF(AND(T
    #   RUE(),OR(TRUNG_GIAN!$E$6=4)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=4)),1,0))>=3),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),3.0)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),3.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),3.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),3.0)))) else 0.0)),3.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),4.0)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),4.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),4.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),4.0)))) else 0.0)),3.0))) else '—')

def f402(E, r, c):
    # Ô K210
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.503): Có ≥ 3 chi Mão hoặc ≥
    #    3 chi Thìn → "&IF(J210="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J210&") ⇒ "&G210&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.503): Có ≥ 3 chi Mão hoặc ≥ 3 chi Thìn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f403(E, r, c):
    # Ô E211, F211
    # = IF((IF(AND(TRUE(),TRUNG_GIAN!$C$7=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$9=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$7=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$9=4)
    #   ,1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$7=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$9=4),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$7=4),1,0)+IF
    #   (AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$9=4),1,0))>=4,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,3),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,4),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,5),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,5),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,9,6),4.0))) else 0.0)),4.0)) else '—')

def f404(E, r, c):
    # Ô K211
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.504): Có ≥ 4 chữ (can + chi
    #   ) hành Thủy → "&IF(J211="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J211&") ⇒ "&G211&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Có ≥ 4 chữ (can + chi) hành Thủy → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f405(E, r, c):
    # Ô E212, F212
    # = IF((IF(AND(TRUE(),TRUNG_GIAN!$C$7=3),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$9=3),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$7=3),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$9=3)
    #   ,1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$7=3),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$9=3),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$7=3),1,0)+IF
    #   (AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$9=3),1,0))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),3.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,3),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,4),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,5),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,5),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,9,6),3.0))) else 0.0)),3.0)) else '—')

def f406(E, r, c):
    # Ô K212
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi
    #   ) hành Kim → "&IF(J212="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J212&") ⇒ "&G212&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi) hành Kim → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f407(E, r, c):
    # Ô E213, F213
    # = IF((IF(AND(TRUE(),TRUNG_GIAN!$C$7=0),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$9=0),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$7=0),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$9=0)
    #   ,1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$7=0),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$9=0),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$7=0),1,0)+IF
    #   (AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$9=0),1,0))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),0.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,3),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,4),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,5),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,5),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,9,6),0.0))) else 0.0)),3.0)) else '—')

def f408(E, r, c):
    # Ô K213
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi
    #   ) hành Mộc → "&IF(J213="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J213&") ⇒ "&G213&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi) hành Mộc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f409(E, r, c):
    # Ô E214, F214
    # = IF((IF(AND(TRUE(),TRUNG_GIAN!$C$7=2),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$9=2),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$7=2),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$9=2)
    #   ,1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$7=2),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$9=2),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$7=2),1,0)+IF
    #   (AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$9=2),1,0))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),2.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,3),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,4),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,5),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,5),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,9,6),2.0))) else 0.0)),3.0)) else '—')

def f410(E, r, c):
    # Ô K214
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi
    #   ) hành Thổ → "&IF(J214="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J214&") ⇒ "&G214&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi) hành Thổ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f411(E, r, c):
    # Ô E215, F215
    # = IF((IF(AND(TRUE(),TRUNG_GIAN!$C$7=1),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$9=1),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$7=1),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$9=1)
    #   ,1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$7=1),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$9=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$7=1),1,0)+IF
    #   (AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$9=1),1,0))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,3),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,5),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,9,5),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,9,6),1.0))) else 0.0)),3.0)) else '—')

def f412(E, r, c):
    # Ô K215
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi
    #   ) hành Hỏa → "&IF(J215="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J215&") ⇒ "&G215&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Có ≥ 3 chữ (can + chi) hành Hỏa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f413(E, r, c):
    # Ô E216, F216
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '—')

def f414(E, r, c):
    # Ô K216
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thất sát bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản k
    #   hí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.504): Bản khí chi ngày là thất sát → "&IF(J216="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J216&") ⇒ "&G216&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P2 tr.504): Bản khí chi ngày là thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f415(E, r, c):
    # Ô E217, F217
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=0,TRUNG_GIAN!$F$6=6,TRUNG_GIAN!$F$6=3,TRUNG_GIAN!$F$6=9)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),0.0), X.eq(E.v(23,6,6),6.0), X.eq(E.v(23,6,6),3.0), X.eq(E.v(23,6,6),9.0)))) else '—')

def f416(E, r, c):
    # Ô K217
    # = "Dữ kiện: "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy
    #    tắc (P2 tr.504): Giờ Tí, Ngọ, Mão, Dậu → "&IF(J217="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J217&") ⇒ "&G217&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Giờ Tí, Ngọ, Mão, Dậu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f417(E, r, c):
    # Ô E218, F218
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=4,TRUNG_GIAN!$F$6=10,TRUNG_GIAN!$F$6=1,TRUNG_GIAN!$F$6=7)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),4.0), X.eq(E.v(23,6,6),10.0), X.eq(E.v(23,6,6),1.0), X.eq(E.v(23,6,6),7.0)))) else '—')

def f418(E, r, c):
    # Ô K218
    # = "Dữ kiện: "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy
    #    tắc (P2 tr.504): Giờ Thìn, Tuất, Sửu, Mùi → "&IF(J218="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J218&") ⇒ "&G218&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Giờ Thìn, Tuất, Sửu, Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f419(E, r, c):
    # Ô E219, F219
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=2,TRUNG_GIAN!$F$6=8,TRUNG_GIAN!$F$6=5,TRUNG_GIAN!$F$6=11)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),2.0), X.eq(E.v(23,6,6),8.0), X.eq(E.v(23,6,6),5.0), X.eq(E.v(23,6,6),11.0)))) else '—')

def f420(E, r, c):
    # Ô K219
    # = "Dữ kiện: "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy
    #    tắc (P2 tr.504): Giờ Dần, Thân, Tị, Hợi → "&IF(J219="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J219&") ⇒ "&G219&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.504): Giờ Dần, Thân, Tị, Hợi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f421(E, r, c):
    # Ô E220, F220
    # = IF(OR(AND(TRUNG_GIAN!$E$7=0,OR(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$D$6=4)),AND(OR(TRUNG_GIAN!$E$7=1,TRUNG_GIAN!$E$7=2),OR(TRUNG_GIAN!$D$6=
    #   5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7)),AND(TRUNG_GIAN!$E$7=3,OR(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=10)),AND(TRUNG_GIAN!$E$7=4,OR(TRU
    #   NG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,7,5),0.0), X.OR(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,4),4.0))), X.AND(X.OR(X.eq(E.v(23,7,5),1.0), X.eq(E.v(23,7,5),2.0)), X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))), X.AND(X.eq(E.v(23,7,5),3.0), X.OR(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),10.0))), X.AND(X.eq(E.v(23,7,5),4.0), X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))))) else '—')

def f422(E, r, c):
    # Ô K220
    # = "Dữ kiện: "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&
    #   "Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr
    #   .504): Nhật chủ sinh đúng mùa vượng (Mộc–xuân, Hỏa/Thổ–hạ, Kim–thu, Thủy–đông) → "&IF(J220="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.
    #   ","ĐỦ ("&J220&") ⇒ "&G220&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.504): Nhật chủ sinh đúng mùa vượng (Mộc–xuân, Hỏa/Thổ–hạ, Kim–thu, Thủy–đông) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f423(E, r, c):
    # Ô C222
    # = TRIM(IF(L203="ĐỦ","• "&G203&" ["&H203&"] ","")&IF(L204="ĐỦ","• "&G204&" ["&H204&"] ","")&IF(L205="ĐỦ","• "&G205&" ["&H205&"] ","")&IF(L206="ĐỦ","• "&G
    #   206&" ["&H206&"] ","")&IF(L207="ĐỦ","• "&G207&" ["&H207&"] ","")&IF(L208="ĐỦ","• "&G208&" ["&H208&"] ","")&IF(L209="ĐỦ","• "&G209&" ["&H209&"] ","")&I
    #   F(L210="ĐỦ","• "&G210&" ["&H210&"] ","")&IF(L211="ĐỦ","• "&G211&" ["&H211&"] ","")&IF(L212="ĐỦ","• "&G212&" ["&H212&"] ","")&IF(L213="ĐỦ","• "&G213&" 
    #   ["&H213&"] ","")&IF(L214="ĐỦ","• "&G214&" ["&H214&"] ","")&IF(L215="ĐỦ","• "&G215&" ["&H215&"] ","")&IF(L216="ĐỦ","• "&G216&" ["&H216&"] ","")&IF(L217
    #   ="ĐỦ","• "&G217&" ["&H217&"] ","")&IF(L218="ĐỦ","• "&G218&" ["&H218&"] ","")&IF(L219="ĐỦ","• "&G219&" ["&H219&"] ","")&IF(L220="ĐỦ","• "&G220&" ["&H22
    #   0&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-19,c+4)),' ['),E.v(20,r-19,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-19,c+9),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-18,c+4)),' ['),E.v(20,r-18,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-18,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-17,c+4)),' ['),E.v(20,r-17,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-17,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-16,c+4)),' ['),E.v(20,r-16,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-16,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-15,c+4)),' ['),E.v(20,r-15,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-15,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-14,c+4)),' ['),E.v(20,r-14,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-14,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-13,c+4)),' ['),E.v(20,r-13,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-13,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-12,c+4)),' ['),E.v(20,r-12,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-12,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-11,c+4)),' ['),E.v(20,r-11,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-11,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-10,c+4)),' ['),E.v(20,r-10,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-10,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-9,c+4)),' ['),E.v(20,r-9,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-9,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-8,c+4)),' ['),E.v(20,r-8,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-8,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-7,c+4)),' ['),E.v(20,r-7,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-7,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-6,c+4)),' ['),E.v(20,r-6,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-6,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-5,c+4)),' ['),E.v(20,r-5,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-5,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-4,c+4)),' ['),E.v(20,r-4,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-4,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-3,c+4)),' ['),E.v(20,r-3,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-3,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,r-2,c+4)),' ['),E.v(20,r-2,c+5)),'] ') if X.truthy(X.eq(E.v(20,r-2,c+9),'ĐỦ')) else '')))

def f424(E, r, c):
    # Ô E227, F227
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=0,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=4),AND(TRUNG_
    #   GIAN!$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=10))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),0.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),10.0))))) else '—')

def f425(E, r, c):
    # Ô K227
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.573; UH tr.102, 243-244): Ngày Giáp Dần, Mậu Thìn, Bính Thìn, Canh Thìn, Nhâm Tuất (nhật đức – UH tr.102) → "&IF(J227="","chưa đủ điều ki
    #   ện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J227&") ⇒ "&G227&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.573; UH tr.102, 243-244): Ngày Giáp Dần, Mậu Thìn, Bính Thìn, Canh Thìn, Nhâm Tuất (nhật đức – UH tr.102) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f426(E, r, c):
    # Ô E228, F228
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else '—')

def f427(E, r, c):
    # Ô K228
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Thực thần bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"B
    #   ản khí",SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.573; UH tr.168, 176): Bản khí chi tháng là thực thần (""nguyệt lệnh gặp thực thần"" – UH tr
    #   .168, 176) → "&IF(J228="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J228&") ⇒ "&G228&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.573; UH tr.168, 176): Bản khí chi tháng là thực thần ("nguyệt lệnh gặp thực thần" – UH tr.168, 176) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f428(E, r, c):
    # Ô E229, F229
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG
    #   _GIAN!$E$5=4,TRUNG_GIAN!$E$6=10))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),10.0))))) else '—')

def f429(E, r, c):
    # Ô K229
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (UH tr.103, 244-245): Ngày Nhâm Thìn, Canh Tuất, Canh Thìn, Mậu Tuất (khôi cương) → "&IF(J229="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J229&") ⇒ "&G229&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (UH tr.103, 244-245): Ngày Nhâm Thìn, Canh Tuất, Canh Thìn, Mậu Tuất (khôi cương) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f430(E, r, c):
    # Ô E230, F230
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!$E$6=9),AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!$E$6=11),AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!$E$6=5),AND(TRUNG
    #   _GIAN!$E$5=9,TRUNG_GIAN!$E$6=3))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,5),9.0)), X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,5),11.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,5),3.0))))) else '—')

def f431(E, r, c):
    # Ô K230
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (UH tr.101-102, 242): Ngày Đinh Dậu, Đinh Hợi, Quý Tị, Quý Mão (nhật quý) → "&IF(J230="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng
    #   .","ĐỦ ("&J230&") ⇒ "&G230&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (UH tr.101-102, 242): Ngày Đinh Dậu, Đinh Hợi, Quý Tị, Quý Mão (nhật quý) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

def f432(E, r, c):
    # Ô E231, F231
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=6),AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=6),AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=0))),"ĐỦ","—"
    #   )
    return ('ĐỦ' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),6.0)), X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),6.0)), X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),0.0))))) else '—')

def f433(E, r, c):
    # Ô K231
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (UH tr.102-103): Ngày Mậu Ngọ, Bính Ngọ, Nhâm Tý (nhật nhận) → "&IF(J231="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J23
    #   1&") ⇒ "&G231&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (UH tr.102-103): Ngày Mậu Ngọ, Bính Ngọ, Nhâm Tý (nhật nhận) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(20,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(20,r,c-1)),') ⇒ '),E.v(20,r,c-4)),'.')))

CELLS = {
    'C5': (f0, 5, 3),
    'H5': (f1, 5, 8),
    'I5': (f2, 5, 9),
    'J5': (f3, 5, 10),
    'K5': (f4, 5, 11),
    'L5': (f5, 5, 12),
    'H6': (f6, 6, 8),
    'I6': (f7, 6, 9),
    'J6': (f8, 6, 10),
    'K6': (f9, 6, 11),
    'L6': (f10, 6, 12),
    'C7': (f11, 7, 3),
    'D7': (f12, 7, 4),
    'C8': (f13, 8, 3),
    'C9': (f14, 9, 3),
    'C12': (f15, 12, 3),
    'C13': (f16, 13, 3),
    'C16': (f17, 16, 3),
    'M16': (f18, 16, 13),
    'C17': (f19, 17, 3),
    'C18': (f20, 18, 3),
    'C19': (f21, 19, 3),
    'C22': (f22, 22, 3),
    'C23': (f23, 23, 3),
    'C24': (f24, 24, 3),
    'C25': (f25, 25, 3),
    'AB25': (f26, 25, 28),
    'AC25': (f27, 25, 29),
    'AD25': (f28, 25, 30),
    'AE25': (f29, 25, 31),
    'AF25': (f30, 25, 32),
    'AG25': (f31, 25, 33),
    'C29': (f32, 29, 3),
    'E29': (f33, 29, 5),
    'F29': (f34, 29, 6),
    'G29': (f35, 29, 7),
    'H29': (f36, 29, 8),
    'I29': (f37, 29, 9),
    'J29': (f38, 29, 10),
    'K29': (f39, 29, 11),
    'C30': (f40, 30, 3),
    'E30': (f41, 30, 5),
    'F30': (f34, 30, 6),
    'G30': (f42, 30, 7),
    'H30': (f36, 30, 8),
    'I30': (f43, 30, 9),
    'J30': (f38, 30, 10),
    'K30': (f39, 30, 11),
    'C31': (f44, 31, 3),
    'E31': (f45, 31, 5),
    'F31': (f34, 31, 6),
    'G31': (f46, 31, 7),
    'H31': (f36, 31, 8),
    'I31': (f47, 31, 9),
    'J31': (f38, 31, 10),
    'K31': (f39, 31, 11),
    'C32': (f48, 32, 3),
    'E32': (f49, 32, 5),
    'F32': (f34, 32, 6),
    'G32': (f50, 32, 7),
    'H32': (f36, 32, 8),
    'I32': (f51, 32, 9),
    'J32': (f38, 32, 10),
    'K32': (f39, 32, 11),
    'C33': (f52, 33, 3),
    'E33': (f53, 33, 5),
    'F33': (f34, 33, 6),
    'G33': (f54, 33, 7),
    'H33': (f36, 33, 8),
    'I33': (f55, 33, 9),
    'J33': (f38, 33, 10),
    'K33': (f39, 33, 11),
    'E34': (f56, 34, 5),
    'F34': (f56, 34, 6),
    'G34': (f56, 34, 7),
    'H34': (f56, 34, 8),
    'I34': (f56, 34, 9),
    'J34': (f56, 34, 10),
    'C35': (f57, 35, 3),
    'C36': (f58, 36, 3),
    'C37': (f59, 37, 3),
    'E42': (f60, 42, 5),
    'F42': (f60, 42, 6),
    'J42': (f61, 42, 10),
    'K42': (f62, 42, 11),
    'L42': (f63, 42, 12),
    'E43': (f64, 43, 5),
    'F43': (f64, 43, 6),
    'J43': (f61, 43, 10),
    'K43': (f65, 43, 11),
    'L43': (f63, 43, 12),
    'E44': (f66, 44, 5),
    'F44': (f66, 44, 6),
    'J44': (f61, 44, 10),
    'K44': (f67, 44, 11),
    'L44': (f63, 44, 12),
    'E45': (f68, 45, 5),
    'F45': (f68, 45, 6),
    'J45': (f61, 45, 10),
    'K45': (f69, 45, 11),
    'L45': (f63, 45, 12),
    'E46': (f70, 46, 5),
    'F46': (f70, 46, 6),
    'J46': (f61, 46, 10),
    'K46': (f71, 46, 11),
    'L46': (f63, 46, 12),
    'E47': (f72, 47, 5),
    'F47': (f72, 47, 6),
    'J47': (f61, 47, 10),
    'K47': (f73, 47, 11),
    'L47': (f63, 47, 12),
    'E48': (f74, 48, 5),
    'F48': (f74, 48, 6),
    'J48': (f61, 48, 10),
    'K48': (f75, 48, 11),
    'L48': (f63, 48, 12),
    'E49': (f76, 49, 5),
    'F49': (f76, 49, 6),
    'J49': (f61, 49, 10),
    'K49': (f77, 49, 11),
    'L49': (f63, 49, 12),
    'E50': (f78, 50, 5),
    'F50': (f78, 50, 6),
    'J50': (f61, 50, 10),
    'K50': (f79, 50, 11),
    'L50': (f63, 50, 12),
    'E51': (f80, 51, 5),
    'F51': (f80, 51, 6),
    'J51': (f61, 51, 10),
    'K51': (f81, 51, 11),
    'L51': (f63, 51, 12),
    'E52': (f82, 52, 5),
    'F52': (f82, 52, 6),
    'J52': (f61, 52, 10),
    'K52': (f83, 52, 11),
    'L52': (f63, 52, 12),
    'E53': (f84, 53, 5),
    'F53': (f84, 53, 6),
    'J53': (f61, 53, 10),
    'K53': (f85, 53, 11),
    'L53': (f63, 53, 12),
    'E54': (f86, 54, 5),
    'F54': (f86, 54, 6),
    'J54': (f61, 54, 10),
    'K54': (f87, 54, 11),
    'L54': (f63, 54, 12),
    'E55': (f88, 55, 5),
    'F55': (f88, 55, 6),
    'J55': (f61, 55, 10),
    'K55': (f89, 55, 11),
    'L55': (f63, 55, 12),
    'E56': (f90, 56, 5),
    'F56': (f90, 56, 6),
    'J56': (f61, 56, 10),
    'K56': (f91, 56, 11),
    'L56': (f63, 56, 12),
    'E57': (f92, 57, 5),
    'F57': (f92, 57, 6),
    'J57': (f61, 57, 10),
    'K57': (f93, 57, 11),
    'L57': (f63, 57, 12),
    'E58': (f94, 58, 5),
    'F58': (f94, 58, 6),
    'J58': (f61, 58, 10),
    'K58': (f95, 58, 11),
    'L58': (f63, 58, 12),
    'E59': (f96, 59, 5),
    'F59': (f96, 59, 6),
    'J59': (f61, 59, 10),
    'K59': (f97, 59, 11),
    'L59': (f63, 59, 12),
    'E60': (f98, 60, 5),
    'F60': (f98, 60, 6),
    'J60': (f61, 60, 10),
    'K60': (f99, 60, 11),
    'L60': (f63, 60, 12),
    'E61': (f100, 61, 5),
    'F61': (f100, 61, 6),
    'J61': (f61, 61, 10),
    'K61': (f101, 61, 11),
    'L61': (f63, 61, 12),
    'E62': (f102, 62, 5),
    'F62': (f102, 62, 6),
    'J62': (f61, 62, 10),
    'K62': (f103, 62, 11),
    'L62': (f63, 62, 12),
    'E63': (f104, 63, 5),
    'F63': (f104, 63, 6),
    'J63': (f61, 63, 10),
    'K63': (f105, 63, 11),
    'L63': (f63, 63, 12),
    'E64': (f106, 64, 5),
    'F64': (f107, 64, 6),
    'J64': (f61, 64, 10),
    'K64': (f108, 64, 11),
    'L64': (f63, 64, 12),
    'E65': (f109, 65, 5),
    'F65': (f110, 65, 6),
    'J65': (f61, 65, 10),
    'K65': (f111, 65, 11),
    'L65': (f63, 65, 12),
    'E66': (f112, 66, 5),
    'F66': (f113, 66, 6),
    'J66': (f61, 66, 10),
    'K66': (f114, 66, 11),
    'L66': (f63, 66, 12),
    'E67': (f115, 67, 5),
    'F67': (f116, 67, 6),
    'J67': (f61, 67, 10),
    'K67': (f117, 67, 11),
    'L67': (f63, 67, 12),
    'E68': (f118, 68, 5),
    'F68': (f119, 68, 6),
    'J68': (f61, 68, 10),
    'K68': (f120, 68, 11),
    'L68': (f63, 68, 12),
    'E69': (f121, 69, 5),
    'F69': (f122, 69, 6),
    'J69': (f61, 69, 10),
    'K69': (f123, 69, 11),
    'L69': (f63, 69, 12),
    'E70': (f124, 70, 5),
    'F70': (f125, 70, 6),
    'J70': (f61, 70, 10),
    'K70': (f126, 70, 11),
    'L70': (f63, 70, 12),
    'E71': (f127, 71, 5),
    'F71': (f128, 71, 6),
    'J71': (f61, 71, 10),
    'K71': (f129, 71, 11),
    'L71': (f63, 71, 12),
    'E72': (f130, 72, 5),
    'F72': (f131, 72, 6),
    'J72': (f61, 72, 10),
    'K72': (f132, 72, 11),
    'L72': (f63, 72, 12),
    'E73': (f133, 73, 5),
    'F73': (f134, 73, 6),
    'J73': (f61, 73, 10),
    'K73': (f135, 73, 11),
    'L73': (f63, 73, 12),
    'E74': (f136, 74, 5),
    'F74': (f137, 74, 6),
    'J74': (f61, 74, 10),
    'K74': (f138, 74, 11),
    'L74': (f63, 74, 12),
    'E75': (f139, 75, 5),
    'F75': (f140, 75, 6),
    'J75': (f61, 75, 10),
    'K75': (f141, 75, 11),
    'L75': (f63, 75, 12),
    'E76': (f142, 76, 5),
    'F76': (f143, 76, 6),
    'J76': (f61, 76, 10),
    'K76': (f144, 76, 11),
    'L76': (f63, 76, 12),
    'E77': (f145, 77, 5),
    'F77': (f146, 77, 6),
    'J77': (f61, 77, 10),
    'K77': (f147, 77, 11),
    'L77': (f63, 77, 12),
    'E78': (f148, 78, 5),
    'F78': (f149, 78, 6),
    'J78': (f61, 78, 10),
    'K78': (f150, 78, 11),
    'L78': (f63, 78, 12),
    'E79': (f151, 79, 5),
    'F79': (f152, 79, 6),
    'J79': (f61, 79, 10),
    'K79': (f153, 79, 11),
    'L79': (f63, 79, 12),
    'E80': (f154, 80, 5),
    'F80': (f155, 80, 6),
    'J80': (f61, 80, 10),
    'K80': (f156, 80, 11),
    'L80': (f63, 80, 12),
    'E81': (f157, 81, 5),
    'F81': (f158, 81, 6),
    'J81': (f61, 81, 10),
    'K81': (f159, 81, 11),
    'L81': (f63, 81, 12),
    'E82': (f160, 82, 5),
    'F82': (f161, 82, 6),
    'J82': (f61, 82, 10),
    'K82': (f162, 82, 11),
    'L82': (f63, 82, 12),
    'E83': (f163, 83, 5),
    'F83': (f164, 83, 6),
    'J83': (f61, 83, 10),
    'K83': (f165, 83, 11),
    'L83': (f63, 83, 12),
    'E84': (f166, 84, 5),
    'F84': (f167, 84, 6),
    'J84': (f61, 84, 10),
    'K84': (f168, 84, 11),
    'L84': (f63, 84, 12),
    'E85': (f169, 85, 5),
    'F85': (f170, 85, 6),
    'J85': (f61, 85, 10),
    'K85': (f171, 85, 11),
    'L85': (f63, 85, 12),
    'E86': (f172, 86, 5),
    'F86': (f173, 86, 6),
    'J86': (f61, 86, 10),
    'K86': (f174, 86, 11),
    'L86': (f63, 86, 12),
    'E87': (f175, 87, 5),
    'F87': (f176, 87, 6),
    'J87': (f61, 87, 10),
    'K87': (f177, 87, 11),
    'L87': (f63, 87, 12),
    'E88': (f178, 88, 5),
    'F88': (f179, 88, 6),
    'J88': (f61, 88, 10),
    'K88': (f180, 88, 11),
    'L88': (f63, 88, 12),
    'E89': (f181, 89, 5),
    'F89': (f182, 89, 6),
    'J89': (f61, 89, 10),
    'K89': (f183, 89, 11),
    'L89': (f63, 89, 12),
    'E90': (f184, 90, 5),
    'F90': (f185, 90, 6),
    'J90': (f61, 90, 10),
    'K90': (f186, 90, 11),
    'L90': (f63, 90, 12),
    'E91': (f187, 91, 5),
    'F91': (f188, 91, 6),
    'J91': (f61, 91, 10),
    'K91': (f189, 91, 11),
    'L91': (f63, 91, 12),
    'E92': (f190, 92, 5),
    'F92': (f191, 92, 6),
    'J92': (f61, 92, 10),
    'K92': (f192, 92, 11),
    'L92': (f63, 92, 12),
    'E93': (f193, 93, 5),
    'F93': (f194, 93, 6),
    'J93': (f61, 93, 10),
    'K93': (f195, 93, 11),
    'L93': (f63, 93, 12),
    'E94': (f196, 94, 5),
    'F94': (f197, 94, 6),
    'J94': (f61, 94, 10),
    'K94': (f198, 94, 11),
    'L94': (f63, 94, 12),
    'E95': (f199, 95, 5),
    'F95': (f200, 95, 6),
    'J95': (f61, 95, 10),
    'K95': (f201, 95, 11),
    'L95': (f63, 95, 12),
    'E96': (f202, 96, 5),
    'F96': (f202, 96, 6),
    'J96': (f61, 96, 10),
    'K96': (f203, 96, 11),
    'L96': (f63, 96, 12),
    'E97': (f204, 97, 5),
    'F97': (f204, 97, 6),
    'J97': (f61, 97, 10),
    'K97': (f205, 97, 11),
    'L97': (f63, 97, 12),
    'E98': (f206, 98, 5),
    'F98': (f206, 98, 6),
    'J98': (f61, 98, 10),
    'K98': (f207, 98, 11),
    'L98': (f63, 98, 12),
    'E99': (f208, 99, 5),
    'F99': (f208, 99, 6),
    'J99': (f61, 99, 10),
    'K99': (f209, 99, 11),
    'L99': (f63, 99, 12),
    'E100': (f210, 100, 5),
    'F100': (f210, 100, 6),
    'J100': (f61, 100, 10),
    'K100': (f211, 100, 11),
    'L100': (f63, 100, 12),
    'E101': (f212, 101, 5),
    'F101': (f212, 101, 6),
    'J101': (f61, 101, 10),
    'K101': (f213, 101, 11),
    'L101': (f63, 101, 12),
    'E102': (f214, 102, 5),
    'F102': (f214, 102, 6),
    'J102': (f61, 102, 10),
    'K102': (f215, 102, 11),
    'L102': (f63, 102, 12),
    'E103': (f216, 103, 5),
    'F103': (f216, 103, 6),
    'J103': (f61, 103, 10),
    'K103': (f217, 103, 11),
    'L103': (f63, 103, 12),
    'E104': (f218, 104, 5),
    'F104': (f218, 104, 6),
    'J104': (f61, 104, 10),
    'K104': (f219, 104, 11),
    'L104': (f63, 104, 12),
    'E105': (f220, 105, 5),
    'F105': (f220, 105, 6),
    'J105': (f61, 105, 10),
    'K105': (f221, 105, 11),
    'L105': (f63, 105, 12),
    'E110': (f72, 110, 5),
    'F110': (f72, 110, 6),
    'J110': (f61, 110, 10),
    'K110': (f222, 110, 11),
    'L110': (f63, 110, 12),
    'E111': (f223, 111, 5),
    'F111': (f223, 111, 6),
    'J111': (f61, 111, 10),
    'K111': (f224, 111, 11),
    'L111': (f63, 111, 12),
    'E112': (f225, 112, 5),
    'F112': (f225, 112, 6),
    'J112': (f61, 112, 10),
    'K112': (f226, 112, 11),
    'L112': (f63, 112, 12),
    'E113': (f227, 113, 5),
    'F113': (f227, 113, 6),
    'J113': (f61, 113, 10),
    'K113': (f228, 113, 11),
    'L113': (f63, 113, 12),
    'E114': (f229, 114, 5),
    'F114': (f229, 114, 6),
    'J114': (f61, 114, 10),
    'K114': (f230, 114, 11),
    'L114': (f63, 114, 12),
    'E115': (f231, 115, 5),
    'F115': (f231, 115, 6),
    'J115': (f61, 115, 10),
    'K115': (f232, 115, 11),
    'L115': (f63, 115, 12),
    'E116': (f74, 116, 5),
    'F116': (f74, 116, 6),
    'J116': (f61, 116, 10),
    'K116': (f233, 116, 11),
    'L116': (f63, 116, 12),
    'E117': (f234, 117, 5),
    'F117': (f235, 117, 6),
    'J117': (f61, 117, 10),
    'K117': (f236, 117, 11),
    'L117': (f63, 117, 12),
    'E118': (f237, 118, 5),
    'F118': (f237, 118, 6),
    'J118': (f61, 118, 10),
    'K118': (f238, 118, 11),
    'L118': (f63, 118, 12),
    'E119': (f239, 119, 5),
    'F119': (f239, 119, 6),
    'J119': (f61, 119, 10),
    'K119': (f240, 119, 11),
    'L119': (f63, 119, 12),
    'E120': (f241, 120, 5),
    'F120': (f241, 120, 6),
    'J120': (f61, 120, 10),
    'K120': (f242, 120, 11),
    'L120': (f63, 120, 12),
    'E121': (f243, 121, 5),
    'F121': (f243, 121, 6),
    'J121': (f61, 121, 10),
    'K121': (f244, 121, 11),
    'L121': (f63, 121, 12),
    'E122': (f245, 122, 5),
    'F122': (f245, 122, 6),
    'J122': (f61, 122, 10),
    'K122': (f246, 122, 11),
    'L122': (f63, 122, 12),
    'E123': (f100, 123, 5),
    'F123': (f100, 123, 6),
    'J123': (f61, 123, 10),
    'K123': (f247, 123, 11),
    'L123': (f63, 123, 12),
    'E124': (f248, 124, 5),
    'F124': (f249, 124, 6),
    'J124': (f61, 124, 10),
    'K124': (f250, 124, 11),
    'L124': (f63, 124, 12),
    'E125': (f251, 125, 5),
    'F125': (f251, 125, 6),
    'J125': (f61, 125, 10),
    'K125': (f252, 125, 11),
    'L125': (f63, 125, 12),
    'E126': (f253, 126, 5),
    'F126': (f253, 126, 6),
    'J126': (f61, 126, 10),
    'K126': (f254, 126, 11),
    'L126': (f63, 126, 12),
    'E127': (f255, 127, 5),
    'F127': (f255, 127, 6),
    'J127': (f61, 127, 10),
    'K127': (f256, 127, 11),
    'L127': (f63, 127, 12),
    'E128': (f257, 128, 5),
    'F128': (f258, 128, 6),
    'J128': (f61, 128, 10),
    'K128': (f259, 128, 11),
    'L128': (f63, 128, 12),
    'E129': (f260, 129, 5),
    'F129': (f261, 129, 6),
    'J129': (f61, 129, 10),
    'K129': (f262, 129, 11),
    'L129': (f63, 129, 12),
    'E130': (f68, 130, 5),
    'F130': (f68, 130, 6),
    'J130': (f61, 130, 10),
    'K130': (f263, 130, 11),
    'L130': (f63, 130, 12),
    'E131': (f264, 131, 5),
    'F131': (f264, 131, 6),
    'J131': (f61, 131, 10),
    'K131': (f265, 131, 11),
    'L131': (f63, 131, 12),
    'E132': (f266, 132, 5),
    'F132': (f266, 132, 6),
    'J132': (f61, 132, 10),
    'K132': (f267, 132, 11),
    'L132': (f63, 132, 12),
    'E133': (f268, 133, 5),
    'F133': (f268, 133, 6),
    'J133': (f61, 133, 10),
    'K133': (f269, 133, 11),
    'L133': (f63, 133, 12),
    'E138': (f270, 138, 5),
    'F138': (f270, 138, 6),
    'J138': (f61, 138, 10),
    'K138': (f271, 138, 11),
    'L138': (f63, 138, 12),
    'E139': (f272, 139, 5),
    'F139': (f272, 139, 6),
    'J139': (f61, 139, 10),
    'K139': (f273, 139, 11),
    'L139': (f63, 139, 12),
    'E140': (f274, 140, 5),
    'F140': (f275, 140, 6),
    'J140': (f61, 140, 10),
    'K140': (f276, 140, 11),
    'L140': (f63, 140, 12),
    'E141': (f277, 141, 5),
    'F141': (f277, 141, 6),
    'J141': (f61, 141, 10),
    'K141': (f278, 141, 11),
    'L141': (f63, 141, 12),
    'E142': (f279, 142, 5),
    'F142': (f279, 142, 6),
    'J142': (f61, 142, 10),
    'K142': (f280, 142, 11),
    'L142': (f63, 142, 12),
    'E143': (f281, 143, 5),
    'F143': (f282, 143, 6),
    'J143': (f61, 143, 10),
    'K143': (f283, 143, 11),
    'L143': (f63, 143, 12),
    'E144': (f284, 144, 5),
    'F144': (f285, 144, 6),
    'J144': (f61, 144, 10),
    'K144': (f286, 144, 11),
    'L144': (f63, 144, 12),
    'E145': (f287, 145, 5),
    'F145': (f287, 145, 6),
    'J145': (f61, 145, 10),
    'K145': (f288, 145, 11),
    'L145': (f63, 145, 12),
    'E146': (f289, 146, 5),
    'F146': (f289, 146, 6),
    'J146': (f61, 146, 10),
    'K146': (f290, 146, 11),
    'L146': (f63, 146, 12),
    'E147': (f291, 147, 5),
    'F147': (f291, 147, 6),
    'J147': (f61, 147, 10),
    'K147': (f292, 147, 11),
    'L147': (f63, 147, 12),
    'E148': (f293, 148, 5),
    'F148': (f293, 148, 6),
    'J148': (f61, 148, 10),
    'K148': (f294, 148, 11),
    'L148': (f63, 148, 12),
    'E149': (f295, 149, 5),
    'F149': (f295, 149, 6),
    'J149': (f61, 149, 10),
    'K149': (f296, 149, 11),
    'L149': (f63, 149, 12),
    'E150': (f297, 150, 5),
    'F150': (f297, 150, 6),
    'J150': (f61, 150, 10),
    'K150': (f298, 150, 11),
    'L150': (f63, 150, 12),
    'E151': (f299, 151, 5),
    'F151': (f299, 151, 6),
    'J151': (f61, 151, 10),
    'K151': (f300, 151, 11),
    'L151': (f63, 151, 12),
    'E152': (f301, 152, 5),
    'F152': (f301, 152, 6),
    'J152': (f61, 152, 10),
    'K152': (f302, 152, 11),
    'L152': (f63, 152, 12),
    'E153': (f303, 153, 5),
    'F153': (f303, 153, 6),
    'J153': (f61, 153, 10),
    'K153': (f304, 153, 11),
    'L153': (f63, 153, 12),
    'E154': (f305, 154, 5),
    'F154': (f305, 154, 6),
    'J154': (f61, 154, 10),
    'K154': (f306, 154, 11),
    'L154': (f63, 154, 12),
    'E155': (f251, 155, 5),
    'F155': (f251, 155, 6),
    'J155': (f61, 155, 10),
    'K155': (f307, 155, 11),
    'L155': (f63, 155, 12),
    'E156': (f308, 156, 5),
    'F156': (f308, 156, 6),
    'J156': (f61, 156, 10),
    'K156': (f309, 156, 11),
    'L156': (f63, 156, 12),
    'E157': (f310, 157, 5),
    'F157': (f310, 157, 6),
    'J157': (f61, 157, 10),
    'K157': (f311, 157, 11),
    'L157': (f63, 157, 12),
    'E158': (f312, 158, 5),
    'F158': (f312, 158, 6),
    'J158': (f61, 158, 10),
    'K158': (f313, 158, 11),
    'L158': (f63, 158, 12),
    'E159': (f314, 159, 5),
    'F159': (f314, 159, 6),
    'J159': (f61, 159, 10),
    'K159': (f315, 159, 11),
    'L159': (f63, 159, 12),
    'E160': (f316, 160, 5),
    'F160': (f316, 160, 6),
    'J160': (f61, 160, 10),
    'K160': (f317, 160, 11),
    'L160': (f63, 160, 12),
    'E161': (f318, 161, 5),
    'F161': (f318, 161, 6),
    'J161': (f61, 161, 10),
    'K161': (f319, 161, 11),
    'L161': (f63, 161, 12),
    'E162': (f320, 162, 5),
    'F162': (f320, 162, 6),
    'J162': (f61, 162, 10),
    'K162': (f321, 162, 11),
    'L162': (f63, 162, 12),
    'E163': (f322, 163, 5),
    'F163': (f322, 163, 6),
    'J163': (f61, 163, 10),
    'K163': (f323, 163, 11),
    'L163': (f63, 163, 12),
    'E164': (f324, 164, 5),
    'F164': (f325, 164, 6),
    'J164': (f61, 164, 10),
    'K164': (f326, 164, 11),
    'L164': (f63, 164, 12),
    'E165': (f327, 165, 5),
    'F165': (f327, 165, 6),
    'J165': (f61, 165, 10),
    'K165': (f328, 165, 11),
    'L165': (f63, 165, 12),
    'E166': (f329, 166, 5),
    'F166': (f330, 166, 6),
    'J166': (f61, 166, 10),
    'K166': (f331, 166, 11),
    'L166': (f63, 166, 12),
    'E167': (f332, 167, 5),
    'F167': (f332, 167, 6),
    'J167': (f61, 167, 10),
    'K167': (f333, 167, 11),
    'L167': (f63, 167, 12),
    'E168': (f334, 168, 5),
    'F168': (f334, 168, 6),
    'J168': (f61, 168, 10),
    'K168': (f335, 168, 11),
    'L168': (f63, 168, 12),
    'E169': (f336, 169, 5),
    'F169': (f336, 169, 6),
    'J169': (f61, 169, 10),
    'K169': (f337, 169, 11),
    'L169': (f63, 169, 12),
    'E170': (f338, 170, 5),
    'F170': (f338, 170, 6),
    'J170': (f61, 170, 10),
    'K170': (f339, 170, 11),
    'L170': (f63, 170, 12),
    'E171': (f340, 171, 5),
    'F171': (f340, 171, 6),
    'J171': (f61, 171, 10),
    'K171': (f341, 171, 11),
    'L171': (f63, 171, 12),
    'E172': (f342, 172, 5),
    'F172': (f342, 172, 6),
    'J172': (f61, 172, 10),
    'K172': (f343, 172, 11),
    'L172': (f63, 172, 12),
    'E173': (f344, 173, 5),
    'F173': (f344, 173, 6),
    'J173': (f61, 173, 10),
    'K173': (f345, 173, 11),
    'L173': (f63, 173, 12),
    'E174': (f346, 174, 5),
    'F174': (f346, 174, 6),
    'J174': (f61, 174, 10),
    'K174': (f347, 174, 11),
    'L174': (f63, 174, 12),
    'E175': (f348, 175, 5),
    'F175': (f348, 175, 6),
    'J175': (f61, 175, 10),
    'K175': (f349, 175, 11),
    'L175': (f63, 175, 12),
    'E176': (f350, 176, 5),
    'F176': (f350, 176, 6),
    'J176': (f61, 176, 10),
    'K176': (f351, 176, 11),
    'L176': (f63, 176, 12),
    'E177': (f352, 177, 5),
    'F177': (f352, 177, 6),
    'J177': (f61, 177, 10),
    'K177': (f353, 177, 11),
    'L177': (f63, 177, 12),
    'E178': (f354, 178, 5),
    'F178': (f354, 178, 6),
    'J178': (f61, 178, 10),
    'K178': (f355, 178, 11),
    'L178': (f63, 178, 12),
    'E179': (f356, 179, 5),
    'F179': (f356, 179, 6),
    'J179': (f61, 179, 10),
    'K179': (f357, 179, 11),
    'L179': (f63, 179, 12),
    'E180': (f358, 180, 5),
    'F180': (f358, 180, 6),
    'J180': (f61, 180, 10),
    'K180': (f359, 180, 11),
    'L180': (f63, 180, 12),
    'E181': (f360, 181, 5),
    'F181': (f360, 181, 6),
    'J181': (f61, 181, 10),
    'K181': (f361, 181, 11),
    'L181': (f63, 181, 12),
    'E182': (f362, 182, 5),
    'F182': (f362, 182, 6),
    'J182': (f61, 182, 10),
    'K182': (f363, 182, 11),
    'L182': (f63, 182, 12),
    'E183': (f364, 183, 5),
    'F183': (f364, 183, 6),
    'J183': (f61, 183, 10),
    'K183': (f365, 183, 11),
    'L183': (f63, 183, 12),
    'E184': (f366, 184, 5),
    'F184': (f366, 184, 6),
    'J184': (f61, 184, 10),
    'K184': (f367, 184, 11),
    'L184': (f63, 184, 12),
    'E185': (f368, 185, 5),
    'F185': (f368, 185, 6),
    'J185': (f61, 185, 10),
    'K185': (f369, 185, 11),
    'L185': (f63, 185, 12),
    'E186': (f370, 186, 5),
    'F186': (f370, 186, 6),
    'J186': (f61, 186, 10),
    'K186': (f371, 186, 11),
    'L186': (f63, 186, 12),
    'E187': (f372, 187, 5),
    'F187': (f372, 187, 6),
    'J187': (f61, 187, 10),
    'K187': (f373, 187, 11),
    'L187': (f63, 187, 12),
    'C193': (f374, 193, 3),
    'D193': (f375, 193, 4),
    'E193': (f376, 193, 5),
    'C194': (f377, 194, 3),
    'D194': (f378, 194, 4),
    'E194': (f379, 194, 5),
    'C195': (f380, 195, 3),
    'D195': (f378, 195, 4),
    'E195': (f381, 195, 5),
    'C196': (f382, 196, 3),
    'D196': (f378, 196, 4),
    'E196': (f383, 196, 5),
    'C197': (f384, 197, 3),
    'D197': (f378, 197, 4),
    'E197': (f385, 197, 5),
    'C198': (f386, 198, 3),
    'E203': (f387, 203, 5),
    'F203': (f387, 203, 6),
    'J203': (f61, 203, 10),
    'K203': (f388, 203, 11),
    'L203': (f63, 203, 12),
    'E204': (f389, 204, 5),
    'F204': (f389, 204, 6),
    'J204': (f61, 204, 10),
    'K204': (f390, 204, 11),
    'L204': (f63, 204, 12),
    'E205': (f391, 205, 5),
    'F205': (f391, 205, 6),
    'J205': (f61, 205, 10),
    'K205': (f392, 205, 11),
    'L205': (f63, 205, 12),
    'E206': (f393, 206, 5),
    'F206': (f393, 206, 6),
    'J206': (f61, 206, 10),
    'K206': (f394, 206, 11),
    'L206': (f63, 206, 12),
    'E207': (f395, 207, 5),
    'F207': (f395, 207, 6),
    'J207': (f61, 207, 10),
    'K207': (f396, 207, 11),
    'L207': (f63, 207, 12),
    'E208': (f397, 208, 5),
    'F208': (f397, 208, 6),
    'J208': (f61, 208, 10),
    'K208': (f398, 208, 11),
    'L208': (f63, 208, 12),
    'E209': (f399, 209, 5),
    'F209': (f399, 209, 6),
    'J209': (f61, 209, 10),
    'K209': (f400, 209, 11),
    'L209': (f63, 209, 12),
    'E210': (f401, 210, 5),
    'F210': (f401, 210, 6),
    'J210': (f61, 210, 10),
    'K210': (f402, 210, 11),
    'L210': (f63, 210, 12),
    'E211': (f403, 211, 5),
    'F211': (f403, 211, 6),
    'J211': (f61, 211, 10),
    'K211': (f404, 211, 11),
    'L211': (f63, 211, 12),
    'E212': (f405, 212, 5),
    'F212': (f405, 212, 6),
    'J212': (f61, 212, 10),
    'K212': (f406, 212, 11),
    'L212': (f63, 212, 12),
    'E213': (f407, 213, 5),
    'F213': (f407, 213, 6),
    'J213': (f61, 213, 10),
    'K213': (f408, 213, 11),
    'L213': (f63, 213, 12),
    'E214': (f409, 214, 5),
    'F214': (f409, 214, 6),
    'J214': (f61, 214, 10),
    'K214': (f410, 214, 11),
    'L214': (f63, 214, 12),
    'E215': (f411, 215, 5),
    'F215': (f411, 215, 6),
    'J215': (f61, 215, 10),
    'K215': (f412, 215, 11),
    'L215': (f63, 215, 12),
    'E216': (f413, 216, 5),
    'F216': (f413, 216, 6),
    'J216': (f61, 216, 10),
    'K216': (f414, 216, 11),
    'L216': (f63, 216, 12),
    'E217': (f415, 217, 5),
    'F217': (f415, 217, 6),
    'J217': (f61, 217, 10),
    'K217': (f416, 217, 11),
    'L217': (f63, 217, 12),
    'E218': (f417, 218, 5),
    'F218': (f417, 218, 6),
    'J218': (f61, 218, 10),
    'K218': (f418, 218, 11),
    'L218': (f63, 218, 12),
    'E219': (f419, 219, 5),
    'F219': (f419, 219, 6),
    'J219': (f61, 219, 10),
    'K219': (f420, 219, 11),
    'L219': (f63, 219, 12),
    'E220': (f421, 220, 5),
    'F220': (f421, 220, 6),
    'J220': (f61, 220, 10),
    'K220': (f422, 220, 11),
    'L220': (f63, 220, 12),
    'C222': (f423, 222, 3),
    'E227': (f424, 227, 5),
    'F227': (f424, 227, 6),
    'J227': (f61, 227, 10),
    'K227': (f425, 227, 11),
    'L227': (f63, 227, 12),
    'E228': (f426, 228, 5),
    'F228': (f426, 228, 6),
    'J228': (f61, 228, 10),
    'K228': (f427, 228, 11),
    'L228': (f63, 228, 12),
    'E229': (f428, 229, 5),
    'F229': (f428, 229, 6),
    'J229': (f61, 229, 10),
    'K229': (f429, 229, 11),
    'L229': (f63, 229, 12),
    'E230': (f430, 230, 5),
    'F230': (f430, 230, 6),
    'J230': (f61, 230, 10),
    'K230': (f431, 230, 11),
    'L230': (f63, 230, 12),
    'E231': (f432, 231, 5),
    'F231': (f432, 231, 6),
    'J231': (f61, 231, 10),
    'K231': (f433, 231, 11),
    'L231': (f63, 231, 12),
}
