# -*- coding: utf-8 -*-
"""Xuất JSON (YÊU CẦU 5.4): dữ liệu nhập + TOÀN BỘ kết quả theo {sheet: {ô: giá trị}} – mọi ô công thức của 33 sheet
(gồm cả sheet chỉ tính ẩn và 288 ô Bảng dữ liệu KHONG_RO_GIO). Giá trị gốc: số, chữ, true/false; lỗi ghi dạng "#N/A"."""
import json
from battu_core import xl as X
from battu_core.du_lieu import ref_of
from battu_core import sheets as S
from export.noi_dung import PHIEN_BAN_QUY_TAC, TEN_PHAN_MEM, ngay_in


def du_lieu(kq):
    E = kq.E; M = E.M; ra = {}
    for si in sorted(E.F):
        o = set(E.F.get(si, {})) | {(r, c) for (s_, r, c) in E.dt if s_ == si}
        d = {}
        for r, c in sorted(o):
            v = E.raw(si, r, c)
            if isinstance(v, X.ErrVal): v = v.code
            elif v is None: v = ''
            d[ref_of(r, c)] = v
        ra[M.sheets[si]] = d
    return {'phan_mem': TEN_PHAN_MEM, 'phien_ban_quy_tac': PHIEN_BAN_QUY_TAC, 'nguon_sha256': S.NGUON_SHA256,
            'ngay_xuat': ngay_in(), 'nhap': kq.la_so.to_dict(), 'ket_qua': ra}


def xuat_json(kq, duong_dan):
    with open(duong_dan, 'w', encoding='utf-8') as f:
        json.dump(du_lieu(kq), f, ensure_ascii=False, indent=1)
    return duong_dan
