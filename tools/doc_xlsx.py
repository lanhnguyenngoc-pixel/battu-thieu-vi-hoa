# -*- coding: utf-8 -*-
"""Đọc trực tiếp XML của workbook .xlsx (không qua openpyxl để không đổi kiểu dữ liệu).

Trả về:
  sheets: danh sách tên sheet theo thứ tự
  cells[sheet][ref] = ('f', công_thức, thuộc_tính)  – ô công thức (đã dịch shared formula)
                    = ('v', giá_trị)                  – ô hằng (số thực / chuỗi / bool)
  names: {tên: 'SHEET!$A$1:$B$2'}
  data_tables[sheet] = [(ref_vùng, thuộc_tính)]
Cách nhận ô công thức giống oracle_dump.py: ô có thẻ <f> bên trong <c ...>…</c>.
"""
import re, zipfile, html
from openpyxl.formula.translate import Translator

_CELL = re.compile(r'<c r="([A-Z]+\d+)"((?:\s[^>]*?)?)(?:/>|>(.*?)</c>)', re.S)


def _attrs(s):
    return dict(re.findall(r'(\w+)="([^"]*)"', s or ''))


def doc(path):
    z = zipfile.ZipFile(path)
    wbx = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2t = {}
    for m in re.finditer(r'<Relationship\s([^>]*)/?>', rels):
        a = _attrs(m.group(1)); rid2t[a.get('Id')] = a.get('Target')
    # chuỗi dùng chung
    ss = []
    if 'xl/sharedStrings.xml' in z.namelist():
        x = z.read('xl/sharedStrings.xml').decode('utf-8')
        for si in re.finditer(r'<si>(.*?)</si>', x, re.S):
            body = re.sub(r'<rPh\b.*?</rPh>', '', si.group(1), flags=re.S)
            ss.append(html.unescape(''.join(re.findall(r'<t(?:\s[^>]*)?>(.*?)</t>', body, re.S))))
    sheets, cells, dts = [], {}, {}
    for m in re.finditer(r'<sheet\s([^>]*)/>', wbx):
        a = _attrs(m.group(1)); name = html.unescape(a['name']); rid = a['r:id'] if 'r:id' in a else a.get('id')
        if rid is None:
            rid = re.search(r'r:id="([^"]+)"', m.group(1)).group(1)
        sheets.append(name)
        tgt = rid2t[rid].lstrip('/')
        if not tgt.startswith('xl/'): tgt = 'xl/' + tgt
        x = z.read(tgt).decode('utf-8')
        cs = {}; shared = {}; dt = []
        for cm in _CELL.finditer(x):
            ref, at, body = cm.group(1), _attrs(cm.group(2)), cm.group(3)
            if body is None: continue
            fm = re.search(r'<f((?:\s[^>]*?)?)(?:/>|>(.*?)</f>)', body, re.S)
            if fm:
                fa = _attrs(fm.group(1)); txt = html.unescape(fm.group(2) or '')
                if fa.get('t') == 'dataTable':
                    dt.append((fa.get('ref'), fa)); cs[ref] = ('f', None, fa); continue
                if fa.get('t') == 'shared':
                    si = fa.get('si')
                    if txt: shared[si] = (ref, txt)
                    else:
                        oref, otxt = shared[si]
                        txt = Translator('=' + otxt, origin=oref).translate_formula(ref)[1:]
                cs[ref] = ('f', txt, fa); continue
            t = at.get('t', 'n')
            vm = re.search(r'<v>(.*?)</v>', body, re.S)
            if t == 's':
                cs[ref] = ('v', ss[int(vm.group(1))])
            elif t == 'inlineStr':
                cs[ref] = ('v', html.unescape(''.join(re.findall(r'<t(?:\s[^>]*)?>(.*?)</t>', body, re.S))))
            elif t == 'b':
                cs[ref] = ('v', vm.group(1) == '1')
            elif t == 'str':
                cs[ref] = ('v', html.unescape(vm.group(1)) if vm else '')
            elif t == 'e':
                cs[ref] = ('v', ('#ERR', vm.group(1) if vm else ''))
            elif vm is not None and vm.group(1) != '':
                cs[ref] = ('v', float(vm.group(1)))
        cells[name] = cs; dts[name] = dt
    names = {}
    for m in re.finditer(r'<definedName\s([^>]*)>(.*?)</definedName>', wbx, re.S):
        a = _attrs(m.group(1))
        if a.get('hidden') == '1' or a['name'].startswith('_xlnm'): continue
        names[html.unescape(a['name'])] = html.unescape(m.group(2))
    return sheets, cells, names, dts


BUILTIN_NUMFMT = {0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
                  12: '# ?/?', 13: '# ??/??', 14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM',
                  19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm', 49: '@'}


def dinh_dang(path):
    """{sheet: {ô: mã định dạng số}} cho mọi ô có định dạng khác General."""
    z = zipfile.ZipFile(path)
    st = z.read('xl/styles.xml').decode('utf-8')
    custom = {int(a): html.unescape(b) for a, b in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', st)}
    xfs = re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', st, re.S).group(1)
    xf_fmt = []
    for m in re.finditer(r'<xf\s([^>]*?)(?:/>|>)', xfs):
        a = _attrs(m.group(1)); nid = int(a.get('numFmtId', 0))
        xf_fmt.append(custom.get(nid, BUILTIN_NUMFMT.get(nid, 'General')))
    wbx = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2t = {}
    for m in re.finditer(r'<Relationship\s([^>]*)/?>', rels):
        a = _attrs(m.group(1)); rid2t[a.get('Id')] = a.get('Target')
    out = {}
    for m in re.finditer(r'<sheet\s([^>]*)/>', wbx):
        a = _attrs(m.group(1)); name = html.unescape(a['name'])
        rid = re.search(r'r:id="([^"]+)"', m.group(1)).group(1)
        tgt = rid2t[rid].lstrip('/'); tgt = tgt if tgt.startswith('xl/') else 'xl/' + tgt
        x = z.read(tgt).decode('utf-8'); d = {}
        for cm in re.finditer(r'<c r="([A-Z]+\d+)"((?:\s[^>]*?)?)(?:/>|>)', x):
            s_ = _attrs(cm.group(2)).get('s')
            if s_ is None: continue
            f = xf_fmt[int(s_)]
            if f != 'General': d[cm.group(1)] = f
        out[name] = d
    return out


def bo_cuc(path):
    """{sheet: {'gop', 'rong', 'an_cot', 'an_dong'}} – ô gộp, độ rộng cột, cột/dòng ẩn của workbook."""
    z = zipfile.ZipFile(path)
    wbx = z.read('xl/workbook.xml').decode('utf-8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2t = {}
    for m in re.finditer(r'<Relationship\s([^>]*)/?>', rels):
        a = _attrs(m.group(1)); rid2t[a.get('Id')] = a.get('Target')
    out = {}
    for m in re.finditer(r'<sheet\s([^>]*)/>', wbx):
        a = _attrs(m.group(1)); name = html.unescape(a['name'])
        rid = re.search(r'r:id="([^"]+)"', m.group(1)).group(1)
        tgt = rid2t[rid].lstrip('/'); tgt = tgt if tgt.startswith('xl/') else 'xl/' + tgt
        x = z.read(tgt).decode('utf-8')
        gop = re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"', x)
        rong = {}; an_cot = []
        for cm in re.finditer(r'<col\s([^>]*)/>', x):
            ca = _attrs(cm.group(1))
            for k in range(int(ca['min']), min(int(ca['max']), 200) + 1):
                if 'width' in ca: rong[k] = float(ca['width'])
                if ca.get('hidden') in ('1', 'true'): an_cot.append(k)
        an_dong = [int(_attrs(rm.group(1))['r']) for rm in re.finditer(r'<row\s([^>]*)>', x)
                   if _attrs(rm.group(1)).get('hidden') in ('1', 'true')]
        out[name] = {'gop': gop, 'rong': rong, 'an_cot': an_cot, 'an_dong': an_dong}
    return out


def danh_sach_chon(path):
    """{sheet: {sqref: [lựa chọn…]}} – danh sách chọn (Data validation kiểu list, giá trị viết trực tiếp) của workbook.
    Ưu tiên <x12ac:list> (giữ đúng mục có dấu phẩy trong ngoặc kép), nếu không có thì đọc <formula1>."""
    import csv, io
    z = zipfile.ZipFile(path)
    wbx = z.read('xl/workbook.xml').decode('utf-8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2t = {}
    for m in re.finditer(r'<Relationship\s([^>]*)/?>', rels):
        a = _attrs(m.group(1)); rid2t[a.get('Id')] = a.get('Target')
    out = {}
    for m in re.finditer(r'<sheet\s([^>]*)/>', wbx):
        a = _attrs(m.group(1)); name = html.unescape(a['name'])
        rid = re.search(r'r:id="([^"]+)"', m.group(1)).group(1)
        tgt = rid2t[rid].lstrip('/'); tgt = tgt if tgt.startswith('xl/') else 'xl/' + tgt
        x = z.read(tgt).decode('utf-8')
        for dv in re.finditer(r'<dataValidation\s([^>]*)>(.*?)</dataValidation>', x, re.S):
            da = _attrs(dv.group(1))
            if da.get('type') != 'list': continue
            lst = re.search(r'<x12ac:list>(.*?)</x12ac:list>', dv.group(2), re.S)
            if lst: muc = next(csv.reader(io.StringIO(html.unescape(lst.group(1)))))
            else:
                f1 = re.search(r'<formula1>(.*?)</formula1>', dv.group(2), re.S)
                f1 = html.unescape(f1.group(1)) if f1 else ''
                if not (f1.startswith('"') and f1.endswith('"')): continue      # danh sách theo tên vùng (CHI_TEN…) – bỏ qua
                muc = f1[1:-1].split(',')
            out.setdefault(name, {})[da['sqref']] = [s.strip() for s in muc]
    return out
