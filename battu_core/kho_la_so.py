# -*- coding: utf-8 -*-
"""Kho lá số (SQLite) – YÊU CẦU 5.3: lưu / sửa / xóa / tìm; nhóm; ghi chú; mở lại cho kết quả như lúc lưu.
Mỗi bản ghi lưu: dữ liệu nhập (LaSo.to_dict), phiên bản quy tắc (6.3 + SHA-256 workbook nguồn) và ảnh chụp
khối TÓM TẮT (KET_LUAN B3:C12) + tứ trụ lúc lưu. Mở lại: tính lại từ dữ liệu nhập; nếu phiên bản quy tắc khác
thì báo và vẫn cho xem ảnh chụp lúc lưu."""
import os, json, sqlite3, datetime
from battu_core.nhap_lieu import LaSo
from battu_core import sheets as S

PHIEN_BAN_QUY_TAC = '6.4'
O_TOM_TAT = ['B3'] + ['%s%d' % (c, r) for r in range(4, 13) for c in 'BC']


def thu_muc_mac_dinh():
    if os.name == 'nt':
        goc = os.environ.get('APPDATA') or os.path.expanduser('~')
    elif os.uname().sysname == 'Darwin':
        goc = os.path.expanduser('~/Library/Application Support')
    else:
        goc = os.environ.get('XDG_DATA_HOME') or os.path.expanduser('~/.local/share')
    d = os.path.join(goc, 'BatTuThieuViHoa'); os.makedirs(d, exist_ok=True)
    return os.path.join(d, 'la_so.sqlite')


def anh_chup(kq):
    """Ảnh chụp kết quả chính lúc lưu (chuỗi hiển thị như workbook)."""
    from battu_core.hien_thi import hien_thi_o
    E = kq.E
    d = {'KET_LUAN!' + r: hien_thi_o(E, 'KET_LUAN', r) for r in O_TOM_TAT}
    d['KET_LUAN!B2'] = hien_thi_o(E, 'KET_LUAN', 'B2')
    d['tu_tru'] = ' · '.join(hien_thi_o(E, 'KET_LUAN', '%s15' % c) for c in 'CDEF')
    return d


class Kho:
    def __init__(self, duong_dan=None):
        self.duong_dan = duong_dan or thu_muc_mac_dinh()
        self.db = sqlite3.connect(self.duong_dan)
        self.db.row_factory = sqlite3.Row
        self.db.execute('''CREATE TABLE IF NOT EXISTS la_so (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ten TEXT NOT NULL DEFAULT '', nhom TEXT NOT NULL DEFAULT '', ghi_chu TEXT NOT NULL DEFAULT '',
            nhap TEXT NOT NULL, phien_ban TEXT NOT NULL, nguon_sha TEXT NOT NULL,
            anh_chup TEXT NOT NULL, tu_tru TEXT NOT NULL DEFAULT '',
            tao_luc TEXT NOT NULL, sua_luc TEXT NOT NULL)''')
        self.db.commit()

    @staticmethod
    def _bay_gio(): return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    def luu(self, kq, ten=None, nhom='', ghi_chu='', id_=None):
        """Lưu mới (id_=None) hoặc ghi đè dữ liệu nhập + ảnh chụp của bản ghi id_. Trả id."""
        s = kq.la_so; ac = anh_chup(kq); ten = ten if ten is not None else (s.ho_ten or '')
        g = (ten, nhom, ghi_chu, json.dumps(s.to_dict(), ensure_ascii=False), PHIEN_BAN_QUY_TAC, S.NGUON_SHA256,
             json.dumps(ac, ensure_ascii=False), ac['tu_tru'])
        if id_ is None:
            t = self._bay_gio()
            cur = self.db.execute('INSERT INTO la_so (ten, nhom, ghi_chu, nhap, phien_ban, nguon_sha, anh_chup, tu_tru, tao_luc, sua_luc)'
                                  ' VALUES (?,?,?,?,?,?,?,?,?,?)', g + (t, t))
            id_ = cur.lastrowid
        else:
            self.db.execute('UPDATE la_so SET ten=?, nhom=?, ghi_chu=?, nhap=?, phien_ban=?, nguon_sha=?, anh_chup=?, tu_tru=?, sua_luc=?'
                            ' WHERE id=?', g + (self._bay_gio(), id_))
        self.db.commit(); return id_

    def sua_thong_tin(self, id_, ten, nhom, ghi_chu):
        self.db.execute('UPDATE la_so SET ten=?, nhom=?, ghi_chu=?, sua_luc=? WHERE id=?', (ten, nhom, ghi_chu, self._bay_gio(), id_))
        self.db.commit()

    def xoa(self, id_):
        self.db.execute('DELETE FROM la_so WHERE id=?', (id_,)); self.db.commit()

    def lay(self, id_):
        r = self.db.execute('SELECT * FROM la_so WHERE id=?', (id_,)).fetchone()
        return dict(r) if r else None

    def la_so(self, id_):
        r = self.lay(id_)
        return LaSo.from_dict(json.loads(r['nhap'])) if r else None

    def cung_phien_ban(self, id_):
        r = self.lay(id_)
        return bool(r) and r['phien_ban'] == PHIEN_BAN_QUY_TAC and r['nguon_sha'] == S.NGUON_SHA256

    def nhom(self):
        return [r[0] for r in self.db.execute("SELECT DISTINCT nhom FROM la_so WHERE nhom<>'' ORDER BY nhom")]

    def tim(self, chu='', nhom=None):
        """Tìm theo tên, ghi chú, nhóm, tứ trụ, ngày sinh (không phân biệt hoa thường)."""
        sql = 'SELECT * FROM la_so'; dk = []; ts = []
        if nhom: dk.append('nhom=?'); ts.append(nhom)
        rows = [dict(r) for r in self.db.execute(sql + (' WHERE ' + ' AND '.join(dk) if dk else '') + ' ORDER BY sua_luc DESC', ts)]
        chu = (chu or '').strip().lower()
        if not chu: return rows
        ra = []
        for r in rows:
            n = json.loads(r['nhap'])
            ngay = '%s/%s/%s' % (n.get('ngay'), n.get('thang'), n.get('nam'))
            if any(chu in (x or '').lower() for x in (r['ten'], r['ghi_chu'], r['nhom'], r['tu_tru'], ngay)): ra.append(r)
        return ra

    # ---- sao lưu / khôi phục (bản 1.4.0)
    def sao_luu(self, duong_dan):
        """Chép toàn bộ kho ra một tệp .sqlite (dùng cơ chế backup của SQLite – an toàn cả khi kho đang mở)."""
        if os.path.abspath(duong_dan) == os.path.abspath(self.duong_dan): raise ValueError('Tệp sao lưu trùng tệp kho.')
        if os.path.exists(duong_dan): os.remove(duong_dan)
        dich = sqlite3.connect(duong_dan)
        try: self.db.backup(dich)
        finally: dich.close()
        return self.db.execute('SELECT COUNT(*) FROM la_so').fetchone()[0]

    def khoi_phuc(self, duong_dan):
        """Nhập các lá số từ một tệp sao lưu vào kho hiện tại (gộp, không xóa gì). Lá trùng hoàn toàn
        (cùng tên, nhóm, dữ liệu nhập) thì bỏ qua. Trả (số lá đã thêm, số lá bỏ qua vì trùng)."""
        nguon = sqlite3.connect('file:%s?mode=ro' % duong_dan, uri=True); nguon.row_factory = sqlite3.Row
        try:
            cot = {r[1] for r in nguon.execute('PRAGMA table_info(la_so)')}
            can = {'ten', 'nhom', 'ghi_chu', 'nhap', 'phien_ban', 'nguon_sha', 'anh_chup', 'tu_tru', 'tao_luc', 'sua_luc'}
            if not can <= cot: raise ValueError('Tệp này không phải bản sao lưu kho lá số.')
            co = {(r['ten'], r['nhom'], r['nhap']) for r in self.db.execute('SELECT ten, nhom, nhap FROM la_so')}
            them = bo = 0
            for r in nguon.execute('SELECT * FROM la_so ORDER BY id'):
                json.loads(r['nhap'])                                   # dữ liệu hỏng → lỗi, không ghi gì
                if (r['ten'], r['nhom'], r['nhap']) in co: bo += 1; continue
                self.db.execute('INSERT INTO la_so (ten, nhom, ghi_chu, nhap, phien_ban, nguon_sha, anh_chup, tu_tru, tao_luc, sua_luc)'
                                ' VALUES (?,?,?,?,?,?,?,?,?,?)', tuple(r[k] for k in ('ten', 'nhom', 'ghi_chu', 'nhap', 'phien_ban',
                                                                                       'nguon_sha', 'anh_chup', 'tu_tru', 'tao_luc', 'sua_luc')))
                co.add((r['ten'], r['nhom'], r['nhap'])); them += 1
            self.db.commit()
        except Exception:
            self.db.rollback(); raise
        finally:
            nguon.close()
        return them, bo

    def dong(self): self.db.close()
