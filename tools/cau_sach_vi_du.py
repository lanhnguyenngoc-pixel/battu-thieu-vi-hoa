"""Bước 3 – Bảng câu sách ↔ ví dụ ứng nghiệm; dò câu đó đã có trong công cụ chưa (so chữ với workbook)."""
import sys, re, json, collections, difflib, unicodedata, openpyxl
sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from dap_an_chuong import VD
G = json.load(open('/home/claude/battu/tai_lieu/giai_doan_A/dap_an_tong.json'))
def chuan(s): return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', unicodedata.normalize('NFC', s.lower()))).strip()
# câu trích trong ngoặc kép
cau = collections.defaultdict(list)
def nhat(txt, la, lv, nh, tr):
    for q in re.findall(r'["“]([^"”]{8,160})["”]', txt):
        cau[chuan(q)].append((la, lv, nh, tr, q))
for (ma, gioi, tr, tt, lv, nh, c, sv) in VD: nhat(c, ma, lv, nh, tr)
for k, v in G.items():
    for lv in ['TC','SN','TL','HN','CC','LT','SK']:
        if lv in v: nhat(v[lv][1] + ' ' + v[lv][3], k, lv, v[lv][0], v[lv][2])
# văn bản workbook
wb = openpyxl.load_workbook('/home/claude/HO_SO_PM/HO_SO_PM/03_NGUON_CHUAN/BAT_TU_THIEU_VI_HOA_6.4.xlsx', read_only=True)
kho = []
for ws in wb:
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        for j, x in enumerate(row):
            if isinstance(x, str) and len(x) > 12 and not x.startswith('='):
                for q in re.findall(r'["“]([^"”]{8,200})["”]', x) or [x[:200]]:
                    kho.append((chuan(q), ws.title, i))
            elif isinstance(x, str) and x.startswith('=') and '"' in x:
                for q in re.findall(r'"([^"]{12,200})"', x): kho.append((chuan(q), ws.title, i))
khoset = {}
for q, sh, r in kho: khoset.setdefault(q, (sh, r))
keys = list(khoset)
out = []
for q, ds in sorted(cau.items(), key=lambda x: -len(x[1])):
    best, ratio, vt = None, 0, ''
    for k in keys:
        if q in k or k in q and len(k) > 15: best, ratio = k, 1.0; break
    if ratio < 1:
        m = difflib.get_close_matches(q, keys, n=1, cutoff=0.72)
        if m: best = m[0]; ratio = difflib.SequenceMatcher(None, q, best).ratio()
    if best: vt = '%s dòng %d' % khoset[best]
    out.append(dict(cau=ds[0][4], so_vd=len(ds), linh_vuc=sorted({d[1] for d in ds}), nhan=collections.Counter(d[2] for d in ds), vi_du=['%s (tr.%s)' % (d[0], d[3]) for d in ds], trong_cong_cu=vt, do_giong=round(ratio, 2)))
json.dump(out, open('/home/claude/battu/tai_lieu/giai_doan_A/cau_sach_vi_du.json', 'w'), ensure_ascii=False, indent=1)
co = sum(1 for o in out if o['trong_cong_cu'])
print('Số câu sách được ví dụ minh chứng:', len(out), '| đã thấy trong công cụ:', co, '| chưa thấy:', len(out) - co)
for o in out[:12]: print(o['so_vd'], o['cau'][:70], '|', o['trong_cong_cu'] or 'CHƯA')
