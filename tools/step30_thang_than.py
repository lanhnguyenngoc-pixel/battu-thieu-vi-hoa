# -*- coding: utf-8 -*-
"""Bản 6.14 (cũng thêm mục VI "Danh mục nghề theo ngũ hành" ở A80:H87, tools/step30_thang_than.py: danh_muc_nghe) – thang kết luận thân vượng/yếu MỚI của mục II bảng NGU_HANH_THIEU (người dùng yêu cầu 28/11/2026):
>=3 quá vượng; >2..<3 vượng (có thông căn, thông gốc và chỉ số >2 thì không cần hóa giải); 1..2 trung hòa; 0..<1 hơi yếu;
-1<..<0 yếu; -2..-1 quá yếu; <-2 cực yếu. Điểm 6 tiêu chí giữ nguyên; chỉ đổi nhãn (F27) và điều kiện "cần hóa giải"
(E27<0 -> E27<1) ở F29, C32, F32 và chú thích B30. Quy ước tại -1: thuộc "quá yếu" (như thang cũ: -2..-1 thuộc mức dưới)."""
S = 'NGU_HANH_THIEU'
F27 = ('IF(E27>=3,"Thân quá vượng",IF(E27>2,"Thân vượng",IF(E27>=1,"Thân trung hòa",IF(E27>=0,"Thân hơi yếu",'
       'IF(E27>-1,"Thân yếu",IF(E27>=-2,"Thân quá yếu","Thân cực yếu"))))))')
B30_OLD_A = 'Kết luận (thang 8 mức): tổng ≥ 3 quá vượng; 2 đến dưới 3 vượng; 1 đến dưới 2 hơi vượng; 0 đến dưới 1 trung hòa; −1 < tổng < 0 hơi yếu; −2 đến −1 yếu; −3 đến dưới −2 quá yếu; < −3 cực yếu (kiểm chứng 34 lá; mốc đúng 0 chưa có mẫu).'
B30_NEW_A = ('Kết luận (thang 7 mức, bản 6.14): tổng ≥ 3 quá vượng; trên 2 đến dưới 3 vượng (có thông căn, thông gốc và tổng > 2 thì không cần hóa giải); '
             '1 đến 2 trung hòa; 0 đến dưới 1 hơi yếu; −1 < tổng < 0 yếu; −2 đến −1 quá yếu; < −2 cực yếu.')
B30_OLD_B = 'Thân yếu (tổng < 0), hoặc không thông căn, hoặc không thông gốc → cần hóa giải: điều chỉnh ngũ hành theo mục I và bổ sung linh phù thông căn, thông gốc để thông khí cho tứ trụ.'
B30_NEW_B = ('Cần hóa giải khi: (1) thân vượng mà không thông căn hoặc không thông gốc (hoặc cả hai); (2) thân trung hòa mà không thông căn hoặc không thông gốc (hoặc cả hai); '
             '(3) thân yếu (tổng dưới 1). Cách hóa giải: điều chỉnh ngũ hành theo mục I và bổ sung Thông căn phù để thông khí cho tứ trụ.')
# đổi tên linh phù -> "Thông căn phù" (chủ dự án 28/11/2026) và quy tắc hóa giải trong chuỗi dùng chung
THAY = (('thân yếu, hoặc không thông căn, hoặc không thông gốc – thì cần hóa giải và bổ sung linh phù thông căn, thông gốc',
         'thân vượng hoặc trung hòa mà không thông căn hoặc không thông gốc, hoặc thân yếu – thì cần hóa giải và bổ sung Thông căn phù'),
        ('thân yếu, hoặc không thông căn, hoặc không thông gốc thì cần hóa giải và bổ sung linh phù thông căn, thông gốc',
         'thân vượng hoặc trung hòa mà không thông căn hoặc không thông gốc, hoặc thân yếu, thì cần hóa giải và bổ sung Thông căn phù'),
        ('linh phù thông căn, thông gốc', 'Thông căn phù'), ('không cần linh phù', 'không cần Thông căn phù'))


def thang_than(p, why):
    p.set_formula(S, 'F27', F27, why=why)
    for ref in ('F29', 'C32', 'F32'):
        f = p.formula(S, ref); assert f.count('E27<0') == 1, ref
        p.set_formula(S, ref, f.replace('E27<0', 'E27<1'), why=why)
    b = p.text_of(S, 'B30'); assert B30_OLD_A in b and B30_OLD_B in b
    p.set_text(S, 'B30', b.replace(B30_OLD_A, B30_NEW_A).replace(B30_OLD_B, B30_NEW_B), why=why)
    for ref in ('F29', 'F32'):
        f = p.formula(S, ref)
        for a, z in THAY: f = f.replace(a, z)
        assert 'linh phù' not in f, ref
        p.set_formula(S, ref, f, why=why)
    ss = p.data['xl/sharedStrings.xml'].decode('utf-8')
    for a, z in THAY: ss = ss.replace(a, z)
    p.data['xl/sharedStrings.xml'] = ss.encode('utf-8')


def danh_muc_nghe(p, why):
    """Thêm mục VI (dòng 80–87) vào NGU_HANH_THIEU: hằng chữ, không công thức. Dòng 80 tiêu đề, 81 ghi chú, 82 đầu bảng, 83–87 năm hành."""
    import sys, os, re
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
    from battu_core.danh_muc_nghe import DANH_MUC_NGHE, THU_TU, GHI_CHU
    import xpatch
    key = 'xl/worksheets/sheet%d.xml' % xpatch.SHEETS[S]
    x = p.data[key].decode('utf-8'); assert '<row r="80"' not in x
    esc = xpatch.esc
    def c(ref, s, t): return '<c r="%s" s="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, esc(t))
    def blank(cols, s): return ''.join('<c r="%s" s="%s"/>' % (k, s) for k in cols)
    rows = ['<row r="80" customFormat="false" ht="18" hidden="false" customHeight="true" outlineLevel="0" collapsed="false">' + c('A80', 49, 'VI. DANH MỤC NGHỀ THEO NGŨ HÀNH (tham khảo khi bổ sung hành thiếu; chủ dự án cung cấp 28/11/2026)') + blank('BCDEFGH', 49) + '</row>',
            '<row r="81" customFormat="false" ht="48" hidden="false" customHeight="true" outlineLevel="0" collapsed="false">' + c('B81', 34, GHI_CHU) + blank('CDEFGH', 34) + '</row>',
            '<row r="82" customFormat="false" ht="20" hidden="false" customHeight="true" outlineLevel="0" collapsed="false">' + c('A82', 29, 'Ngũ hành') + c('B82', 29, 'Nghề thuộc hành') + blank('CDEFGH', 29) + '</row>']
    for i, h in enumerate(THU_TU):
        r = 83 + i
        rows.append('<row r="%d" customFormat="false" ht="62" hidden="false" customHeight="true" outlineLevel="0" collapsed="false">' % r + c('A%d' % r, 29, h) + c('B%d' % r, 34, DANH_MUC_NGHE[h]) + blank('CDEFGH', 34) + '</row>')
    x = x.replace('</sheetData>', ''.join(rows) + '</sheetData>', 1)
    merges = ['A80:H80', 'B81:H81', 'B82:H82'] + ['B%d:H%d' % (83 + i, 83 + i) for i in range(5)]
    m = re.search(r'<mergeCells count="(\d+)">', x); n = int(m.group(1))
    x = x.replace(m.group(0), '<mergeCells count="%d">' % (n + len(merges)), 1).replace('</mergeCells>', ''.join('<mergeCell ref="%s"/>' % r for r in merges) + '</mergeCells>', 1)
    x = re.sub(r'<dimension ref="A1:Y78"/>', '<dimension ref="A1:Y87"/>', x, 1)
    p.data[key] = x.encode('utf-8'); p.log.append((S, 'A80:H87', None, 'mục VI danh mục nghề', why))
