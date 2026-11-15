# -*- coding: utf-8 -*-
"""Tạo HƯỚNG DẪN SỬ DỤNG (PDF, có hình) từ ảnh giao diện thật.
Dùng: QT_QPA_PLATFORM=offscreen python tools/tao_huong_dan.py <thư_mục_ảnh> <ra.pdf>"""
import os, sys
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen'); os.environ.setdefault('BATTU_DONG_BO', '1')
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from ui.giao_dien import tao_app
from PySide6.QtGui import QImage

CSS = """
body { font-family: 'Inter'; font-size: 9.6pt; color: #1e2140; }
h1 { font-size: 22pt; color: #3730a3; margin: 0 0 6pt 0; }
h2 { font-size: 14pt; color: #3730a3; margin: 18pt 0 6pt 0; border-bottom: 1px solid #c7c2f5; }
h3 { font-size: 11pt; margin: 10pt 0 4pt 0; }
td { padding: 3pt 5pt; border: 0.5px solid #d9dbe9; vertical-align: top; }
th { padding: 3pt 5pt; background: #f3f0ff; color: #4338ca; text-align: left; }
.chu_y { background: #fff7e6; color: #92400e; padding: 6pt; }
.nho { color: #6b7090; font-size: 8.4pt; }
"""

H = []
def anh(ten, chu=''):
    H.append('<p align="center"><img src="anh://%s" width="900"><br><span class="nho">%s</span></p>' % (ten, chu))


H.append('<h1>Bát Tự Thiệu Vĩ Hoa</h1><p><b>Hướng dẫn sử dụng – phiên bản 1.6.1 (quy tắc 6.5)</b></p>')
H.append('<p>Phần mềm lập và luận giải lá số Tứ trụ (Bát tự) theo sách <i>Dự đoán theo Tứ trụ</i> của Thiệu Vĩ Hoa, theo bộ quy tắc 6.5; '
         '<b>kết quả đã được kiểm chứng trên hàng chục triệu phép tính</b>. Phần mềm chạy hoàn toàn '
         'không cần Internet; lá số chỉ lưu trên máy của bạn.</p>')
H.append('<p class="chu_y">Nội dung sức khỏe chỉ để tham khảo, không phải chẩn đoán y khoa. Các dòng ghi "quy ước kỹ thuật" hoặc '
         '"lời khuyên của công cụ" không phải câu chữ của sách.</p>')

H.append('<h2>1. Cài đặt và gỡ cài đặt</h2>')
H.append('<table width="100%" cellspacing="0"><tr><th width="16%">Hệ máy</th><th width="42%">Cài đặt</th><th>Gỡ cài đặt</th></tr>'
         '<tr><td><b>Windows 10/11</b> (64-bit)</td><td>Bấm đúp <b>BatTuThieuViHoa_Setup_1.6.1_Windows.exe</b> → Tiếp → Cài đặt. '
         'Nếu Windows SmartScreen cảnh báo (phần mềm chưa mua chứng thư số), bấm <i>Thông tin thêm → Vẫn chạy</i>. '
         'Biểu tượng có ở menu Start và màn hình nền.</td><td><i>Cài đặt → Ứng dụng → Bát Tự Thiệu Vĩ Hoa → Gỡ cài đặt</i>, hoặc menu Start → '
         '<i>Gỡ cài đặt Bát Tự Thiệu Vĩ Hoa</i>.</td></tr>'
         '<tr><td><b>macOS</b> 11 trở lên</td><td>Mở tệp <b>.dmg</b> (bản <i>arm64</i> cho máy chip Apple M, bản <i>x86_64</i> cho máy Intel), '
         'kéo biểu tượng vào thư mục <i>Applications</i>. Lần đầu mở: nhấp chuột phải vào ứng dụng → <i>Mở</i> → <i>Mở</i> '
         '(phần mềm chưa ký với Apple).</td><td>Kéo ứng dụng từ <i>Applications</i> vào Thùng rác.</td></tr></table>')
H.append('<p>Lá số đã lưu nằm trong tệp <b>la_so.sqlite</b>: Windows <i>%APPDATA%\\BatTuThieuViHoa</i>; macOS '
         '<i>~/Library/Application Support/BatTuThieuViHoa</i>. Gỡ phần mềm <b>không xóa</b> tệp này (để không mất lá số); muốn sao lưu '
         'hãy chép tệp này sang nơi khác.</p>')

H.append('<h2>2. Bố cục màn hình và Chế độ chuyên gia</h2><p>Thanh bên trái chia 3 nhóm. Công tắc <b>Chế độ chuyên gia</b> ở cuối thanh bên '
         '(phần mềm ghi nhớ lựa chọn) hiện thêm các bảng chi tiết dành cho người nghiên cứu.</p>'
         '<table width="100%" cellspacing="0"><tr><th width="26%">Mục</th><th>Dùng để</th></tr>'
         '<tr><td colspan="2"><b>Đọc lá số – luôn hiện (theo thứ tự trên thanh bên)</b></td></tr>'
         '<tr><td>✎ Nhập liệu</td><td><b>Màn hình đầu tiên</b> khi mở phần mềm: nhập ngày giờ sinh hoặc tứ trụ; sửa xong bấm <b>▶ Luận giải</b> (hoặc Enter, F5). Phần mềm tính ở nền nên vẫn dùng được trong lúc tính; lá đã tính được nhớ lại, mở lại không phải tính lại.</td></tr>'
         '<tr><td>☷ Lá số</td><td>Lá số trực quan: 4 trụ can chi màu theo ngũ hành, thập thần, tàng can, trạng thái nhật can; nạp âm năm; dải 10 đại vận (vận đang hành viền tím).</td></tr>'
         '<tr><td>★ Luận giải chung</td><td>9 dòng kết luận của 9 lĩnh vực; nút Nhập/sửa, Xuất/In, Lưu.</td></tr>'
         '<tr><td>❦ Luận giải chi tiết</td><td>Bài đọc liền mạch 8 mục + đánh giá tổng thể + 6 ý cho từng kết luận; chữ gạch chân có chú giải.</td></tr>'
         '<tr><td>↯ Vận trình</td><td>Đại vận (thêm Lưu niên 100 năm, Theo dõi ở chế độ chuyên gia).</td></tr>'
         '<tr><td>▤ Theo lĩnh vực</td><td>Các mục B–K theo lĩnh vực, trình bày dạng văn bản đọc (mỗi mục mở đầu bằng khung KẾT LUẬN, rồi căn cứ chi tiết) và ô tìm kiếm.</td></tr>'
         '<tr><td>☯ Ngũ hành thiếu</td><td>Mục I, II, ★, III, IV, V và 2 biểu đồ.</td></tr>'
         '<tr><td>✓ Kiểm chứng sự việc</td><td>Nhập sự việc đã xảy ra để đối chiếu với lá số.</td></tr>'
         '<tr><td>⚑ Đối chiếu thực tế</td><td>Ghi từng năm đã qua tốt hay xấu và luận giải đúng đến đâu; phần mềm cho biết PP1 hay PP2 khớp đời thực hơn.</td></tr>'
         '<tr><td colspan="2"><b>Chuyên gia – hiện khi bật công tắc</b></td></tr>'
         '<tr><td>◫ Trung gian · ⚖ Phân tích · ◎ Chuyên đề</td><td>Bảng chi tiết các bước tính (tứ trụ trung gian, PP1, PP2, thông căn, đối chiếu, dụng thần, 10 chuyên đề); thẻ Lưu niên 100 năm và Theo dõi trong Vận trình; số dòng/cột của bảng.</td></tr>'
         '<tr><td>◷ Không rõ giờ</td><td>Hiện khi bật "Không rõ giờ sinh" (hoặc ở chế độ chuyên gia).</td></tr>'
         '<tr><td colspan="2"><b>Tham khảo và quản lý</b></td></tr>'
         '<tr><td>❐ Tham khảo · ☰ Danh sách lá số · ⇄ So sánh</td><td>Trang chủ, Hướng dẫn, Tham khảo ngoài sách, Tiết khí; lưu/tìm/nhóm lá số; đặt hai lá cạnh nhau.</td></tr></table>')
anh('22_che_do_thuong_thanh_ben.png', 'Chế độ thường (trái) – bật Chế độ chuyên gia để hiện thêm Trung gian, Phân tích, Chuyên đề')
H.append('<p>Trong mọi bảng: rê chuột lên ô để xem toàn văn ô dài và <b>chú giải thuật ngữ</b> có trong ô.</p>')

H.append('<h2>3. Nhập lá số</h2>')
anh('24_nhap_phoi_ngau.png', 'Màn hình Nhập liệu: cột trái nhập (có ngày sinh vợ/chồng), cột phải tứ trụ và các kiểm tra, cảnh báo')
H.append('<p>Chọn một trong ba cách: <b>Dương lịch</b> (ngày, tháng, năm 1899–2101, giờ, phút, múi giờ – mặc định UTC+7), '
         '<b>Âm lịch</b> (đánh dấu <i>Tháng nhuận</i> nếu sinh tháng nhuận; phần mềm đổi sang dương lịch rồi tính), hoặc '
         '<b>Tự nhập can chi</b> 4 trụ. Chọn giới tính và <b>năm xem</b> (mặc định năm hiện tại). Không biết giờ sinh: bật '
         '<i>Không rõ giờ sinh</i> – phần mềm tính đủ 12 giờ (mục ◷ Không rõ giờ). <i>Tùy chọn nâng cao</i>: chọn vượng suy dùng cho kết luận, '
         'chế độ Ngũ hành thiếu, số ngày tới tiết Lập (sinh tháng Thìn/Mùi/Tuất/Sửu), tháng/ngày âm cho phương hướng kết hôn.</p>'
         '<p><b>Ngày sinh của vợ/chồng</b> (tùy chọn): nhập ngày, tháng, năm, giờ, phút dương lịch – hoặc chỉ nhập năm – để mục Hôn nhân '
         'xét hợp tuổi theo nạp âm và can chi. Trụ năm của người kia đổi tại Lập xuân (P2 tr.268): nếu chỉ nhập năm, phần mềm coi như sinh sau '
         'Lập xuân và có ghi cảnh báo.</p>'
         '<p>Ô màu vàng nhạt ở cột phải là <b>cảnh báo</b>: sát ranh giới giờ (±10 phút), sát giao tiết (±1 ngày), '
         'thiếu số ngày tới tiết Lập, PP1 cần luận giải, PP1 ≠ PP2.</p>')

H.append('<h2>4. Luận giải chung – đọc kết luận trước</h2>')
anh('01_tom_tat.png', 'Luận giải chung: mỗi thẻ là dòng KẾT LUẬN của một lĩnh vực; bấm "Chi tiết mục …" để mở căn cứ')
H.append('<p>Người mới học nên đọc theo thứ tự trên thanh bên: <b>Lá số → Luận giải chung → Luận giải chi tiết → Vận trình</b>. '
         'Kết luận và nội dung do phần mềm luận giải in <b>chữ đậm</b>. Câu sách và số trang (ví dụ <i>P1 tr.70</i> = Phần 1 trang 70) '
         'được tách ra dòng riêng, chữ thường, ẩn sau liên kết <b>📖 Căn cứ sách ▾</b> – bấm để xem, bấm lần nữa để ẩn. '
         'Bản in PDF ghi "Căn cứ sách: …" ngay dưới mỗi kết luận.</p>')

H.append('<h2>4b. Lá số trực quan</h2>')
anh('18_la_so_truc_quan.png', 'Lá số: can chi lớn tô màu theo hành (Mộc xanh lá, Hỏa đỏ, Thổ nâu, Kim vàng, Thủy xanh dương); trụ ngày (nhật chủ) nền tím')
H.append('<p>Mỗi thẻ đại vận ghi: tuổi và năm bắt đầu – kết thúc, thập thần của can vận và <b>KẾT LUẬN</b> theo cách của sách Thiệu Vĩ Hoa '
         '(Rất tốt, Tốt – chữ đỏ / Xấu, Rất xấu), kèm tóm tắt 10 năm (5 năm đầu theo can, 5 năm sau theo chi). '
         'Rê chuột lên thẻ để xem lý do từng nửa vận.</p>')
H.append('<h2>5. Theo lĩnh vực, Luận giải và tìm kiếm</h2>')
anh('02_linh_vuc_C_ban_menh.png', 'Mục C – Bản mệnh: khung KẾT LUẬN ở trên, căn cứ chi tiết ở dưới; chữ đỏ đậm = điểm tốt')
anh('05_tim_kiem.png', 'Tìm kiếm: gõ từ khóa rồi Enter; bấm một kết quả để nhảy tới chỗ đó')
anh('19_luan_giai_bai_doc.png', 'Luận giải chi tiết dạng bài đọc: nhãn KẾT LUẬN, CÁT TƯỜNG, [Sách], [Kinh nghiệm]; chữ gạch chân – rê chuột xem chú giải')
H.append('<p><b>Chú giải thuật ngữ</b> giải thích thuật ngữ bằng lời thường, '
         'có số trang sách; danh sách đầy đủ ở cuối trang Luận giải chi tiết.</p>')
anh('20_chu_giai_thuat_ngu.png', 'Bảng chú giải thuật ngữ cuối trang Luận giải chi tiết')
H.append('<h3>Màu sắc</h3><table width="100%" cellspacing="0">'
         '<tr><th width="30%">Dấu hiệu</th><th>Ý nghĩa</th></tr>'
         '<tr><td><font color="#c00000"><b>Chữ đỏ đậm</b></font></td><td>Điểm tốt/thuận: dụng thần, đại vận có xu thế Tốt, năm đáng chú ý trong lưu niên, '
         'người thân Tốt/Khá (lục thân, hôn nhân), THUẬN/NGHIÊNG THUẬN/VỀ GIÀU… (mệnh cục, sự nghiệp, sức khỏe), năm có cờ trong bảng vận trình của Luận giải chi tiết.</td></tr>'
         '<tr><td><span style="background:#ffe699">Nền vàng</span></td><td>Dòng của năm xem trong bảng Theo dõi.</td></tr>'
         '<tr><td><span style="background:#e2efda">Nền xanh</span> / <span style="background:#fbe4e1">nền hồng</span></td><td>Năm "Tốt" / "Xấu" trong bảng Theo dõi.</td></tr></table>')

H.append('<h2>6. Phân tích, Vận trình, Ngũ hành thiếu, Chuyên đề, Không rõ giờ</h2>')
anh('17_phan_tich.png', 'Phân tích: PP1 định tính, PP2 định lượng, Thông căn, Đối chiếu, Dụng – hỷ – kỵ thần')
anh('06_dai_van_dinh_dang_dk.png', 'Vận trình – Đại vận: đầu trang là Kết luận đại vận – mức của từng vận, vì sao 5 năm đầu / 5 năm sau, tài lộc, công danh; căn cứ sách ẩn sau liên kết')
anh('16_ngu_hanh_thieu.png', 'Ngũ hành thiếu: chọn chế độ PA2 – Theo sách / PA1 – Theo trang; biểu đồ năng lực và biểu đồ đại vận')
anh('15_chuyen_de_luc_than.png', 'Chuyên đề: 10 thẻ (Lục thân, Hôn nhân, Mệnh cục, Tài lộc, Tố chất, Sự nghiệp, Sức khỏe, Thần sát, Sao, Bảo phù)')

H.append('<h2>6b. Kiểm chứng sự việc</h2>')
anh('21_kiem_chung_su_viec.png', 'Nhập năm và loại sự việc đã xảy ra, bấm Áp dụng – các cột bên phải tính từ lưu niên của đúng năm đó')
H.append('<p>Nhập tối đa 30 sự việc: năm dương lịch, chọn <b>loại sự việc</b> (10 loại), ghi chú tùy ý; bấm <b>✔ Áp dụng &amp; tính lại</b>. '
         'Dòng "Kết quả" cho biết bao nhiêu sự việc có dấu hiệu cùng loại trong lá số. Sự việc được lưu cùng lá số. Đây là công cụ kiểm tra, không phải tiên đoán.</p>')
H.append('<h3>Không rõ giờ sinh – tìm giờ theo việc lớn đã xảy ra</h3>')
anh('25_viec_lon_tim_gio.png', 'Không rõ giờ → thẻ "Việc lớn đã xảy ra": nhập năm và loại việc, bấm Áp dụng')
H.append('<p>Theo cách của sách (P2 tr.445): khi không biết giờ sinh, căn cứ một số việc lớn đã xảy ra để tìm giờ. Nhập tối đa 12 việc '
         '(8 loại: kết hôn, ly hôn, bất lợi cho cha, cho mẹ, tai nạn – bệnh nặng, thăng tiến – phát tài, thất bại – kiện tụng, sinh con). '
         'Cột "Với giờ đang xem" cho biết việc đó có khớp dấu hiệu năm của giờ đang xem không; thẻ <b>Bảng 12 giờ</b> xếp giờ khớp nhiều việc nhất lên đầu.</p>')
H.append('<h2>6c. Đối chiếu thực tế</h2>')
anh('26_doi_chieu_thuc_te.png', 'Đối chiếu thực tế: ghi năm đó Tốt / Bình thường / Xấu, sự kiện và đánh giá luận giải; ✔/✗ = khớp xu thế PP1, PP2')
H.append('<p>Với các năm đã qua, chọn <b>Thực tế năm đó</b> (Tốt / Bình thường / Xấu), ghi sự kiện, chọn <b>Đánh giá luận giải</b> '
         '(Đúng / Một phần / Sai / Chưa rõ) rồi bấm <b>✔ Áp dụng &amp; tính lại</b>. Ô tóm tắt ở trên cho tỉ lệ khớp của PP1 và PP2; khi đã ghi '
         'từ 5 năm có sự việc rõ, phần mềm gợi ý nên dùng phương pháp nào (P1 tr.125) – đổi ở Nhập liệu → Tùy chọn nâng cao. Dữ liệu được lưu cùng lá số.</p>')
H.append('<h2>7. Lưu, tìm, so sánh lá số</h2>')
anh('27_danh_sach_sao_luu.png', 'Danh sách lá số: tìm theo tên, nhóm, ghi chú, tứ trụ, ngày sinh; lọc theo nhóm; sao lưu và khôi phục')
H.append('<p>Bấm <b>💾 Lưu lá số</b> (ở Luận giải chung hoặc Danh sách) → đặt tên, nhóm, ghi chú. <b>Mở</b> (hoặc bấm đúp) để nạp lại – phần mềm '
         'tính lại từ dữ liệu nhập và cho kết quả như lúc lưu (mỗi bản ghi lưu kèm phiên bản quy tắc; nếu quy tắc khác, phần mềm báo). '
         '<b>Ghi đè</b> thay dữ liệu nhập bằng lá đang xem; <b>Sửa</b> đổi tên/nhóm/ghi chú; <b>Xóa</b> không hoàn tác. '
         'Chọn 2 lá (giữ Ctrl) → <b>So sánh</b>.</p>'
         '<p><b>⤓ Sao lưu kho…</b> chép toàn bộ lá số ra một tệp (.sqlite) – nên làm định kỳ hoặc trước khi đổi máy. '
         '<b>⤒ Khôi phục từ bản sao lưu…</b> thêm các lá trong tệp sao lưu vào kho hiện tại; lá đã có thì bỏ qua, không xóa lá nào.</p>'
         '<p><b>⚑ Thống kê đối chiếu</b> cộng các năm đã ghi ở trang Đối chiếu thực tế của mọi lá đã lưu: tỉ lệ khớp của PP1, PP2 và tỉ lệ '
         'luận giải đúng trên cả kho; xuất được ra Excel. Càng nhiều lá có sự việc thật, con số càng có ý nghĩa.</p>')
anh('28_thong_ke_doi_chieu.png', 'Thống kê đối chiếu thực tế trên cả kho lá số')
anh('12_so_sanh_hai_la.png', 'So sánh hai lá: khối KẾT LUẬN đặt cạnh nhau, dòng khác nhau tô nền vàng')

H.append('<h2>8. Xuất và in</h2>')
anh('14_hop_xuat_in.png', 'Hộp thoại Xuất / In')
H.append('<table width="100%" cellspacing="0"><tr><th width="16%">Định dạng</th><th>Nội dung</th></tr>'
         '<tr><td>PDF / In</td><td>Thông tin lá số, Luận giải chung, 2 biểu đồ (năng lực ngũ hành, đại vận), các mục đã chọn (mặc định các mục B–K theo lĩnh vực và bảng Ngũ hành thiếu); '
         'mỗi trang có "Trang x/y", ngày in, phiên bản quy tắc. Bảng quá rộng (Lưu niên 100 năm, Theo dõi) xem ở bản XLSX.</td></tr>'
         '<tr><td>XLSX</td><td>Các bảng kết quả (mỗi phần một trang tính, giữ bố cục bảng, ô gộp, màu), không có công thức; trang đầu ghi dữ liệu nhập và phiên bản.</td></tr>'
         '<tr><td>JSON</td><td>Dữ liệu nhập và toàn bộ kết quả tính dạng dữ liệu (dùng cho lưu trữ, đối soát kỹ thuật).</td></tr></table>')

H.append('<h2>9. Câu hỏi thường gặp</h2><table width="100%" cellspacing="0">'
         '<tr><th width="34%">Câu hỏi</th><th>Trả lời</th></tr>'
         '<tr><td>Kết quả có đáng tin không?</td><td>Mọi phép tính theo bộ quy tắc 6.5 đã được kiểm chứng trên hàng chục triệu phép tính (xem Báo cáo kiểm thử).</td></tr>'
         '<tr><td>Dòng "Dụng thần theo cách bổ hành khuyết" là gì?</td><td>Khi trong phe dụng thần có nhóm không có chữ nào trong tứ trụ, phần mềm ghi thêm nhóm đó '
         '(mục B – Bản mệnh). Khi giải ví dụ, Thiệu Vĩ Hoa nhiều lần lấy nhóm đang khuyết làm dụng thần (P1 tr.97). Đây chỉ là ghi chú để cân nhắc, '
         'không thay dụng thần đã chọn.</td></tr>'
         '<tr><td>Vì sao PP1 và PP2 khác nhau?</td><td>Hai phương pháp của sách; xem mục ⚖ Phân tích → Đối chiếu. Lựa chọn dùng cho kết luận đặt ở Nhập liệu → Tùy chọn nâng cao.</td></tr>'
         '<tr><td>Sinh tháng nhuận âm lịch?</td><td>Chọn Âm lịch và đánh dấu Tháng nhuận; phần mềm đổi đúng sang dương lịch.</td></tr>'
         '<tr><td>Có cần Internet?</td><td>Không. Phần mềm không gửi dữ liệu đi đâu.</td></tr></table>')
H.append('<p class="nho">Nguồn sách: Thiệu Vĩ Hoa, <i>Dự đoán theo Tứ trụ</i> (Phần 1, Phần 2) và các sách tham khảo ghi ở mục Tham khảo.</p>')


def main():
    thu_muc, ra = sys.argv[1], sys.argv[2]
    app = tao_app([])
    from export.xuat_pdf import html_ra_pdf
    import re
    anh_ = {'anh://' + t: QImage(os.path.join(thu_muc, t)) for t in set(re.findall(r'anh://([\w.]+)', ''.join(H)))}
    for k, v in anh_.items():
        if v.isNull(): raise SystemExit('Thiếu ảnh ' + k)
    n = html_ra_pdf('<html><body>' + ''.join(H) + '</body></html>', ra, 'Hướng dẫn sử dụng – Bát Tự Thiệu Vĩ Hoa', anh_, CSS)
    print('Hướng dẫn sử dụng:', ra, n, 'trang')


if __name__ == '__main__':
    main()
