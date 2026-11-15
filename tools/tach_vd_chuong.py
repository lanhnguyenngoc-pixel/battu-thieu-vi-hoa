import re, json, sys
D='/home/claude/ocr_tvh2'
lines=[]
for p in range(432,585):
    try: t=open(f'{D}/p{p}.txt').read()
    except: continue
    for l in t.splitlines():
        if l.strip() and not l.strip().endswith('.pdf') and not re.fullmatch(r'\s*\d{3}\s*',l): lines.append((p,l.rstrip()))
hdr=re.compile(r'^\s*V[íi] ?d[uụ]\s*([0-9ồõðốã8]+)\s*[,.]\s*(.*)')
blocks=[]; cur=None; chuong=''
for p,l in lines:
    m=re.match(r'^\s*Chương\s*([0-9 †!]+)',l)
    if m: chuong=m.group(1).replace(' ','')
    h=hdr.match(l)
    if h:
        if cur: blocks.append(cur)
        cur=dict(trang=p,chuong=chuong,dau=l.strip(),txt=[])
        continue
    if cur is not None:
        if m: blocks.append(cur); cur=None; continue
        cur['txt'].append(l)
if cur: blocks.append(cur)
json.dump(blocks,open('/home/claude/battu/tai_lieu/giai_doan_A/vd_chuong_tho.json','w'),ensure_ascii=False)
a,b=int(sys.argv[1]),int(sys.argv[2]); n=int(sys.argv[3]) if len(sys.argv)>3 else 1400
for i,bk in enumerate(blocks):
    if a<=bk['trang']<=b:
        s='\n'.join(bk['txt']); print(f"### #{i} tr.{bk['trang']} | {bk['dau']}\n{s[:n]}\n")
print('TONG',len(blocks))
