# -*- coding: utf-8 -*-
"""Xuất bản luận giải tổng hợp (battu_core.tong_hop.viet_ban) ra HTML (màn hình), PDF và Word .docx.
PDF dùng lớp phân trang sẵn có (export.xuat_pdf._ve: đầu trang, chân trang, "Trang x/y", phông Inter).
.docx dựng bằng thư viện chuẩn (zipfile + XML WordprocessingML) – không thêm phụ thuộc."""
import html, re, zipfile, datetime

CSS = """
body { font-family: 'Inter'; font-size: 10.5pt; color: #1e2140; }
h1 { font-size: 16pt; color: #3730a3; margin: 0 0 2pt 0; }
h2 { font-size: 12.5pt; color: #3730a3; margin: 12pt 0 4pt 0; border-bottom: 1px solid #c7c2f5; }
p { margin: 3pt 0 5pt 0; }
ul { margin: 0 0 2pt 0; }
li { margin: 0 0 1pt 0; }
table.bang { border-collapse: collapse; margin: 3pt 0 4pt 0; }
table.bang td, table.bang th { border: 0.5px solid #d9dbe9; padding: 2pt 3.5pt; vertical-align: top; }
table.bang th { background: #f3f0ff; color: #4338ca; text-align: left; }
.nho { color: #6b7090; font-size: 8.5pt; }
"""


def _in(t):
    """Chữ có **đậm** → HTML."""
    return re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', html.escape(t or ''))


def html_ban(B, css=True):
    h = ['<html><head><style>%s</style></head><body>' % CSS if css else '<html><body>']
    for k in B:
        if k[0] == 'h1': h.append('<h1>%s</h1>' % _in(k[1]))
        elif k[0] == 'h2': h.append('<h2>%s</h2>' % _in(k[1]))
        elif k[0] == 'p': h.append('<p>%s</p>' % _in(k[1]))
        elif k[0] == 'nho': h.append('<p class="nho">%s</p>' % _in(k[1]))
        elif k[0] == 'ul': h.append('<ul>%s</ul>' % ''.join('<li>%s</li>' % _in(x) for x in k[1]))
        elif k[0] == 'bang':
            h.append('<table class="bang" width="100%%" cellspacing="0"><tr>%s</tr>' % ''.join('<th>%s</th>' % _in(x) for x in k[1]))
            for r in k[2]: h.append('<tr>%s</tr>' % ''.join('<td>%s</td>' % _in(x) for x in r))
            h.append('</table>')
    h.append('</body></html>')
    return ''.join(h)


def xuat_pdf(B, duong_dan, tieu_de='Luận giải Bát tự', ten=''):
    """PDF A4. Cần QApplication (giao diện hoặc QT_QPA_PLATFORM=offscreen). → số trang."""
    from PySide6.QtGui import QTextDocument, QPdfWriter, QPageSize, QPageLayout, QFont
    from PySide6.QtCore import QMarginsF
    from export.xuat_pdf import _ve
    from export.noi_dung import TEN_PHAN_MEM, chu_ky
    doc = QTextDocument(); doc.setDefaultStyleSheet(CSS); doc.setDefaultFont(QFont('Inter', 10.5))
    doc.setHtml(html_ban(B, css=False))
    w = QPdfWriter(duong_dan); w.setResolution(150)
    w.setPageSize(QPageSize(QPageSize.A4)); w.setPageMargins(QMarginsF(15, 10, 15, 10), QPageLayout.Millimeter)
    w.setTitle(tieu_de); w.setCreator(TEN_PHAN_MEM + ' · ' + chu_ky())
    return _ve(doc, w, None, '%s%s' % (TEN_PHAN_MEM, (' · ' + ten) if ten else ''), 'Luận giải tổng hợp · ' + chu_ky())


# ------------------------------------------------------------------ Word .docx (WordprocessingML tối thiểu)
_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _x(t): return html.escape(t or '', quote=False)


def _runs(t, co=20, mau=None, dam=False):
    ra = []
    for i, phan in enumerate(re.split(r'\*\*', t or '')):
        if not phan: continue
        b = dam or (i % 2 == 1)
        rp = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>%s%s<w:sz w:val="%d"/></w:rPr>' % (
            '<w:b/>' if b else '', '<w:color w:val="%s"/>' % mau if mau else '', co)
        ra.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rp, _x(phan)))
    return ''.join(ra)


def _p(t, co=20, mau=None, dam=False, truoc=40, sau=60, le=None):
    ind = '<w:ind w:left="360" w:hanging="220"/>' if le else ''
    return '<w:p><w:pPr><w:spacing w:before="%d" w:after="%d"/>%s</w:pPr>%s%s</w:p>' % (
        truoc, sau, ind, _runs('• ', co) if le else '', _runs(t, co, mau, dam))


def _bang(dau, hang):
    n = len(dau); rong = 9600 // max(n, 1)
    bd = ''.join('<w:%s w:val="single" w:sz="4" w:color="D9DBE9"/>' % k for k in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    x = ['<w:tbl><w:tblPr><w:tblW w:w="9600" w:type="dxa"/><w:tblBorders>%s</w:tblBorders></w:tblPr><w:tblGrid>%s</w:tblGrid>' % (
        bd, ''.join('<w:gridCol w:w="%d"/>' % rong for _ in range(n)))]
    for j, r in enumerate([dau] + hang):
        x.append('<w:tr>')
        for o in r:
            nen = '<w:shd w:val="clear" w:color="auto" w:fill="F3F0FF"/>' if j == 0 else ''
            x.append('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>' % (rong, nen, _p(o, 17, '4338CA' if j == 0 else None, j == 0, 0, 0)))
        x.append('</w:tr>')
    x.append('</w:tbl>')
    return ''.join(x) + _p('', 8, truoc=0, sau=0)


def xuat_docx(B, duong_dan, tieu_de='Luận giải Bát tự'):
    than = []
    for k in B:
        if k[0] == 'h1': than.append(_p(k[1], 30, '3730A3', True, 0, 80))
        elif k[0] == 'h2': than.append(_p(k[1], 23, '3730A3', True, 180, 60))
        elif k[0] == 'p': than.append(_p(k[1]))
        elif k[0] == 'nho': than.append(_p(k[1], 16, '6B7090'))
        elif k[0] == 'ul': than += [_p(x, le=True, truoc=0, sau=30) for x in k[1]]
        elif k[0] == 'bang': than.append(_bang(k[1], k[2]))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="850" '
           'w:header="400" w:footer="400" w:gutter="0"/></w:sectPr></w:body></w:document>') % (_W, ''.join(than))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '</Relationships>')
    now = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:creator>Bát Tự Thiệu Vĩ Hoa</dc:creator>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created></cp:coreProperties>') % (_x(tieu_de), now)
    with zipfile.ZipFile(duong_dan, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', doc); z.writestr('docProps/core.xml', core)
    return duong_dan
