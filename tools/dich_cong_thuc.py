# -*- coding: utf-8 -*-
"""BỘ DỊCH CÔNG THỨC (MT-03, PA-A đã duyệt 13/11/2026)
Đọc công thức NGUYÊN VĂN từ XML của workbook 6.3 → sinh module Python cho từng sheet
(battu_core/sheets/<module>.py). Mỗi ô công thức = một hàm (các ô cùng mẫu dùng chung hàm, tham số r, c).
KHÔNG gõ lại công thức nào bằng tay. Mã sinh ra không sửa tay: muốn đổi thì sửa workbook rồi dịch lại.

Dùng: python dich_cong_thuc.py <workbook.xlsx> <pham_vi.json> <thư_mục_ra battu_core/sheets>
"""
import sys, os, re, json, hashlib, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from doc_xlsx import doc
from openpyxl.formula.tokenizer import Tokenizer, Token
from openpyxl.utils.cell import range_boundaries, column_index_from_string, coordinate_from_string, get_column_letter

MODULE = {'TRANG_CHU': 'trang_chu', 'NHAP_LIEU': 'nhap_lieu', 'KET_LUAN': 'ket_luan', 'LUAN_GIAI': 'luan_giai',
          'THEO_DOI': 'theo_doi', 'KHONG_RO_GIO': 'khong_ro_gio', 'DOI_CHIEU': 'doi_chieu', 'DUNG_HY_KY': 'dung_hy_ky',
          'PP1_DINH_TINH': 'pp1', 'PP2_DINH_LUONG': 'pp2', 'THONG_CAN': 'thong_can', 'THAN_SAT': 'than_sat', 'SAO': 'sao',
          'DAI_VAN': 'dai_van', 'LUU_NIEN': 'luu_nien', 'LUC_THAN': 'luc_than', 'HON_NHAN': 'hon_nhan',
          'MENH_CUC': 'menh_cuc', 'TAI_LOC': 'tai_loc', 'TO_CHAT': 'to_chat', 'SU_NGHIEP': 'su_nghiep',
          'SUC_KHOE': 'suc_khoe', 'TRUNG_GIAN': 'trung_gian', 'CSDL': 'csdl', 'KIEM_THU': 'kiem_thu',
          'THAM_KHAO': 'tham_khao', 'NGU_HANH_THIEU': 'ngu_hanh_thieu', 'BAO_PHU': 'bao_phu',
          'DOC_NHANH': 'doc_nhanh', 'SU_KIEN': 'su_kien'}

# Vị trí tham số nhận VÙNG Ô (không lấy giá trị)
ALL = 'all'
REFARGS = {'INDEX': {0}, 'MATCH': {1}, 'COUNTIF': {0}, 'COUNTIFS': 'even', 'SUMPRODUCT': ALL, 'SUM': ALL,
           'MAX': ALL, 'MIN': ALL, 'AVERAGE': ALL, 'COUNT': ALL, 'COUNTA': ALL, 'LARGE': {0}, 'SMALL': {0},
           'RANK': {1}, 'ROW': {0}, 'COLUMN': {0}, 'AND': ALL, 'OR': ALL}
# Hàm cần tính lười / bắt lỗi: tham số bọc lambda
LAZY = {'IFERROR': ALL, 'IFNA': ALL, 'ISNUMBER': ALL, 'ISTEXT': ALL, 'ISNA': ALL, 'ISERROR': ALL, 'ISERR': ALL,
        'ISBLANK': ALL, 'ISLOGICAL': ALL, 'CHOOSE': 'tail'}
RENAME = {'NA': 'NA_'}
BINOP = {'+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '^': 'pow_', '&': 'cat',
         '=': 'eq', '<>': 'ne', '<': 'lt', '<=': 'le', '>': 'gt', '>=': 'ge'}
PREC = {'=': 1, '<>': 1, '<': 1, '<=': 1, '>': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5}


class Dich:
    def __init__(self, sheets, names):
        self.sheets = sheets; self.si = {s: i for i, s in enumerate(sheets)}; self.names = names

    # ---------- phân tích cú pháp (đệ quy theo độ ưu tiên kiểu Excel)
    def parse(self, formula):
        toks = [t for t in Tokenizer('=' + formula).items if t.type != Token.WSPACE]
        self.t = toks; self.i = 0
        node = self.expr(0)
        if self.i != len(self.t): raise SyntaxError('thừa token: %r' % self.t[self.i:])
        return node

    def peek(self): return self.t[self.i] if self.i < len(self.t) else None
    def nxt(self): tk = self.t[self.i]; self.i += 1; return tk

    def expr(self, minp):
        left = self.unary()
        while True:
            tk = self.peek()
            if tk is None or tk.type != Token.OP_IN or tk.value not in PREC: break
            p = PREC[tk.value]
            if p < minp: break
            self.nxt()
            right = self.expr(p + 1)
            left = ('op', tk.value, left, right)
        return left

    def unary(self):
        tk = self.peek()
        if tk is not None and tk.type == Token.OP_PRE:
            self.nxt(); a = self.unary()
            return ('neg', a) if tk.value == '-' else a
        a = self.primary()
        while self.peek() is not None and self.peek().type == Token.OP_POST:
            self.nxt(); a = ('pct', a)
        return a

    def primary(self):
        tk = self.nxt()
        if tk.type == Token.OPERAND:
            st = tk.subtype
            if st == Token.NUMBER: return ('num', float(tk.value))
            if st == Token.TEXT: return ('str', tk.value[1:-1].replace('""', '"'))
            if st == Token.LOGICAL: return ('bool', tk.value.upper() == 'TRUE')
            if st == Token.ERROR: return ('err', tk.value)
            return self.ref(tk.value)
        if tk.type == Token.FUNC and tk.subtype == Token.OPEN:
            name = tk.value[:-1].upper()
            if name.startswith('_XLFN.'): name = name[6:]
            args = []
            if self.peek().type == Token.FUNC and self.peek().subtype == Token.CLOSE:
                self.nxt(); return ('func', name, args)
            while True:
                tkp = self.peek()
                if tkp.type == Token.SEP or (tkp.type == Token.FUNC and tkp.subtype == Token.CLOSE):
                    args.append(('missing',))
                else:
                    args.append(self.expr(0))
                tk2 = self.nxt()
                if tk2.type == Token.FUNC and tk2.subtype == Token.CLOSE: break
                if tk2.type != Token.SEP: raise SyntaxError('mong dấu phẩy: %r' % tk2)
            return ('func', name, args)
        if tk.type == Token.PAREN and tk.subtype == Token.OPEN:
            e = self.expr(0); c = self.nxt()
            if c.type != Token.PAREN: raise SyntaxError('thiếu )')
            return e
        if tk.type == Token.FUNC and tk.subtype == Token.OPEN and tk.value == '{':
            pass
        if tk.type == 'ARRAY' or tk.value == '{':
            rows = [[]]
            while True:
                t2 = self.nxt()
                if t2.value == '}': break
                if t2.type == Token.SEP:
                    if t2.subtype == Token.ROW: rows.append([])
                    continue
                if t2.type == Token.OPERAND:
                    if t2.subtype == Token.NUMBER: rows[-1].append(float(t2.value))
                    elif t2.subtype == Token.TEXT: rows[-1].append(t2.value[1:-1].replace('""', '"'))
                    elif t2.subtype == Token.LOGICAL: rows[-1].append(t2.value.upper() == 'TRUE')
                    else: raise SyntaxError('mảng: %r' % t2)
                elif t2.type == Token.OP_PRE and t2.value == '-':
                    t3 = self.nxt(); rows[-1].append(-float(t3.value))
            return ('arr', rows)
        raise SyntaxError('token lạ: %r %r %r' % (tk.type, tk.subtype, tk.value))

    def ref(self, s):
        if s in self.names:
            return self.ref(self.names[s])
        if '!' in s:
            sh, r = s.rsplit('!', 1); sh = sh.strip("'").replace("''", "'")
        else:
            sh, r = None, s
        parts = r.split(':')
        def one(p):
            m = re.fullmatch(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', p)
            if not m: raise SyntaxError('địa chỉ lạ: %r' % s)
            return (int(m.group(4)), m.group(3) == '$', column_index_from_string(m.group(2)), m.group(1) == '$')
        a = one(parts[0]); b = one(parts[1]) if len(parts) > 1 else a
        return ('ref', sh, a, b)

    # ---------- sinh mã
    def gen(self, node, cur, r0, c0, ctx='v', arr=False):
        k = node[0]
        if arr:
            if k == 'op':
                return "X.A(%r,%s,%s)" % (BINOP[node[1]], self.gen(node[2], cur, r0, c0, 'r', True), self.gen(node[3], cur, r0, c0, 'r', True))
            if k == 'neg':
                return "X.A('neg',%s)" % self.gen(node[1], cur, r0, c0, 'r', True)
            if k == 'func' and node[1] in ('MOD', 'MID', 'SEARCH', 'MATCH'):
                return 'X.AF_%s(%s)' % (node[1], ', '.join(self.gen(a, cur, r0, c0, 'r', True) for a in node[2]))
            if k == 'func' and node[1] == 'ISNUMBER':
                return 'X.AF_ISNUMBER(lambda: %s)' % self.gen(node[2][0], cur, r0, c0, 'r', True)
            if k == 'ref' and node[2][::2] == node[3][::2]:
                ctx = 'v'
        if k == 'num': return repr(node[1])
        if k == 'str': return repr(node[1])
        if k == 'bool': return 'True' if node[1] else 'False'
        if k == 'err': return 'X.loi(%r)' % node[1]
        if k == 'missing': return 'X.MISSING'
        if k == 'arr': return 'X.Arr(%r)' % (node[1],)
        if k == 'ref':
            sh = node[1] or cur; si = self.si[sh]
            (ra, raa, ca, caa), (rb, rba, cb, cba) = node[2], node[3]
            def R(row, ab): return str(row) if ab else ('r' if row == r0 else 'r%+d' % (row - r0))
            def C(col, ab): return str(col) if ab else ('c' if col == c0 else 'c%+d' % (col - c0))
            single = (ra, ca) == (rb, cb)
            if single and ctx == 'v':
                return 'E.v(%d,%s,%s)' % (si, R(ra, raa), C(ca, caa))
            return 'E.rg(%d,%s,%s,%s,%s)' % (si, R(ra, raa), C(ca, caa), R(rb, rba), C(cb, cba))
        if k == 'neg': return 'X.neg(%s)' % self.gen(node[1], cur, r0, c0)
        if k == 'pct': return 'X.pct(%s)' % self.gen(node[1], cur, r0, c0)
        if k == 'op':
            return 'X.%s(%s,%s)' % (BINOP[node[1]], self.gen(node[2], cur, r0, c0, 'v'), self.gen(node[3], cur, r0, c0, 'v'))
        if k == 'func':
            name, args = node[1], node[2]
            if name == 'IF':
                cnd = self.gen(args[0], cur, r0, c0)
                a = self.gen(args[1], cur, r0, c0, 'v') if len(args) > 1 else 'False'
                b = self.gen(args[2], cur, r0, c0, 'v') if len(args) > 2 else 'False'
                if len(args) > 1 and args[1][0] == 'missing': a = 'X.MISSING'
                if len(args) > 2 and args[2][0] == 'missing': b = 'X.MISSING'
                return '(%s if X.truthy(%s) else %s)' % (a, cnd, b)
            if name in ('ROW', 'COLUMN') and not args:
                return 'float(r)' if name == 'ROW' else 'float(c)'
            ra = REFARGS.get(name, set()); lz = LAZY.get(name)
            out = []
            for i, a in enumerate(args):
                isref = ra == ALL or (ra == 'even' and i % 2 == 0) or (isinstance(ra, set) and i in ra)
                code = self.gen(a, cur, r0, c0, 'r' if isref else 'v', arr or name == 'SUMPRODUCT')
                if lz == ALL or (lz == 'tail' and i > 0):
                    code = 'lambda: ' + code
                out.append(code)
            fn = RENAME.get(name, name)
            if not hasattr(__import__('battu_core.xl', fromlist=['x']), fn):
                raise NotImplementedError('Hàm chưa có trong thư viện xl: ' + name)
            return 'X.%s(%s)' % (fn, ', '.join(out))
        raise ValueError(node)


def main():
    wbp, pv, outdir = sys.argv[1:4]
    sha = hashlib.sha256(open(wbp, 'rb').read()).hexdigest()
    sheets, cells, names, dts = doc(wbp)
    D = Dich(sheets, names)
    pham_vi = json.load(open(pv))
    os.makedirs(outdir, exist_ok=True)
    tong = {}
    for sh, spec in pham_vi['sheets'].items():
        rects = None
        if spec != 'all':
            rects = [range_boundaries(x) for x in spec]
        mod = MODULE[sh]; si = D.si[sh]
        pats = {}; cellmap = []; mau_o = {}
        for ref, v in cells[sh].items():
            if v[0] != 'f' or v[1] is None: continue
            col, row = coordinate_from_string(ref); c = column_index_from_string(col)
            if rects and not any(c1 <= c <= c2 and r1 <= row <= r2 for c1, r1, c2, r2 in rects): continue
            ast = D.parse(v[1])
            code = D.gen(ast, sh, row, c, 'v')
            if code not in pats:
                pats[code] = 'f%d' % len(pats); mau_o[code] = (ref, v[1], [])
            mau_o[code][2].append(ref)
            cellmap.append((ref, pats[code], row, c))
        L = ['# -*- coding: utf-8 -*-',
             '"""Module sheet %s – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.' % sh,
             'Nguồn: %s (SHA-256 %s).' % (os.path.basename(wbp), sha),
             'Phạm vi: %s. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""' % (
                 'toàn sheet' if spec == 'all' else ', '.join(spec)),
             'from battu_core import xl as X', '', 'SHEET = %r' % sh, 'SI = %d' % si, '']
        for code, fname in pats.items():
            ref, f, refs = mau_o[code]
            L.append('def %s(E, r, c):' % fname)
            L.append('    # Ô %s%s' % (', '.join(refs[:12]), ' … (%d ô)' % len(refs) if len(refs) > 12 else ''))
            for i in range(0, len(f), 150):
                L.append('    # %s %s' % ('=' if i == 0 else ' ', f[i:i + 150]))
            L.append('    return %s' % code)
            L.append('')
        L.append('CELLS = {')
        for ref, fname, row, c in cellmap:
            L.append('    %r: (%s, %d, %d),' % (ref, fname, row, c))
        L.append('}')
        open(os.path.join(outdir, mod + '.py'), 'w').write('\n'.join(L) + '\n')
        tong[sh] = (len(cellmap), len(pats))
        print('%-16s %-14s ô=%5d  mẫu=%4d' % (sh, mod + '.py', len(cellmap), len(pats)))
    # chỉ mục
    I = ['# -*- coding: utf-8 -*-', '"""Danh sách module sheet đã dịch – SINH TỰ ĐỘNG."""',
         'NGUON_SHA256 = %r' % sha, 'SHEETS = %r' % (sheets,), 'PHAM_VI = %r' % (pham_vi['sheets'],),
         'MODULES = %r' % ({s: MODULE[s] for s in pham_vi['sheets']},)]
    open(os.path.join(outdir, '__init__.py'), 'w').write('\n'.join(I) + '\n')


if __name__ == '__main__':
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
    main()
