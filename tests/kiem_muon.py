# Kiểm: ô nào của giai đoạn sau thực sự được đọc khi tính GĐ hiện tại; chạy lại chỉ mượn đúng các ô đó.
import sys, os, json, collections
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine, cell_modules
from battu_core.du_lieu import model, ref_of
from tests.so_dap_an import oracle_nhap, khop, ten_tep
kt = sys.argv[1]; M = model(); F = cell_modules()
cases = json.load(open(os.path.join(kt, 'cases_dap_an_o.json'), encoding='utf-8'))
dung = n = 0; da_doc = collections.Counter()
for c, nhap in oracle_nhap(cases):
    da = json.load(open(os.path.join(kt, 'dap_an_o', ten_tep(c['id'])), encoding='utf-8'))
    can = {}; truoc = -1
    while True:  # lặp tới khi tập ô mượn ổn định (nhánh IF phụ thuộc giá trị mượn)
        E0 = Engine({'NHAP_LIEU': nhap}, can)
        for si in F: E0.tinh_sheet(M.sheets[si])
        for si, r, cc in E0.chua_cai:
            sh = M.sheets[si]; can.setdefault(sh, {})[ref_of(r, cc)] = da[sh][ref_of(r, cc)]
        n_ = sum(len(v) for v in can.values())
        if n_ == truoc: break
        truoc = n_
    for sh, d_ in can.items():
        for ref in d_: da_doc[sh + '!' + ref] += 1
    E = Engine({'NHAP_LIEU': nhap}, can)
    for si in F:
        sh = M.sheets[si]
        for ref, v in E.tinh_sheet(sh).items():
            n += 1; dung += khop(v, da[sh][ref])
print('Ô giai đoạn sau đã đọc:', dict(da_doc)); print('Khớp khi chỉ mượn các ô đó: %d/%d' % (dung, n))
