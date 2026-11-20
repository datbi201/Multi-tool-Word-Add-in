# Rev0.9.6 — Giao diện tiếng Anh và nút đổi ngôn ngữ Việt, Anh

1. **Nút Ngôn ngữ (VI | EN)** trên Ribbon, ở nhóm Cài đặt của thẻ Chuẩn hóa và ở thẻ Công cụ. Bấm một lần là đổi **ngay** Ribbon (nhãn nút, tên nhóm, tên thẻ, chú thích) và bảng điều khiển đang mở, **không cần mở lại Word**. Cũng chọn được ở **Cài đặt**, mục Giao diện (Ngôn ngữ giao diện).
2. **Phạm vi dịch:** toàn bộ Ribbon (166 chuỗi), hộp Cài đặt (nhãn, chú thích, hướng dẫn nhanh), bảng điều khiển (ba thẻ Chuẩn hóa kỹ thuật, Phương trình, Công cụ AI, danh sách rà soát, chân trang đếm token), hộp chọn màu, và các thông báo kết quả (thanh trạng thái Word, ô thông báo, hộp thoại): tính trong bảng, thao tác bảng, dàn trang và đoạn văn, định dạng số, chuẩn hóa kỹ thuật, rà soát, công cụ AI. Khoảng 1.085 chuỗi nguyên câu và 678 mảnh thông báo ghép (số và tên riêng giữ nguyên vị trí).
3. **Mặc định vẫn là tiếng Việt**, không tự nhận theo Windows, nên người đang dùng không bị đổi giao diện.
4. **Chữ chưa có bản dịch giữ tiếng Việt** (không lỗi). Cố ý giữ tiếng Việt: chữ đọc số và đơn vị bằng lời (Đọc số), luật rà soát chính tả, ngữ pháp, văn phong **của văn bản tiếng Việt** và các ví dụ tiếng Việt trong chú thích, kết quả do AI trả về (theo ngôn ngữ đích bạn chọn), nhật ký chẩn đoán, tên ngôn ngữ "Tiếng Việt" trong ô chọn.
5. Sửa nhỏ: hướng dẫn nhanh trong Cài đặt cập nhật theo ba thẻ hiện tại (bỏ nhắc thẻ "My tab" cũ và nút Xuất đã bỏ); danh sách rà soát: hai cột ký hiệu Sửa, Bỏ rộng hơn cho chữ tiếng Anh.
6. Kỹ thuật: lớp mới `Lang.cs`, dữ liệu `LangData.cs`. Ribbon chạy thật chuyển nhãn và chú thích tĩnh thành lệnh gọi lại (getLabel, getScreentip, getSupertip) nên đổi ngôn ngữ chỉ cần làm mới Ribbon. Bảng điều khiển dịch lại từ bản gốc tiếng Việt nên đổi qua lại nhiều lần vẫn chính xác từng chữ.
7. Kiểm thử: bộ mới `tests-lang` (67 kiểm tra): mặc định vi, lưu và nạp ngôn ngữ, giá trị lạ về vi, dịch nguyên câu và theo mảnh (không khớp giữa một từ khác), 100% nhãn và chú thích Ribbon có bản dịch, lệnh gọi lại Ribbon khớp hàm trên Connect, hộp Cài đặt và bảng điều khiển không còn chữ tiếng Việt chưa dịch, đổi en rồi về vi trả lại đúng từng chữ gốc. Cả 13 bộ đạt.
8. Chưa thử trong Word thật qua Ribbon của bản cài (cần `install.bat`); đã chụp bảng điều khiển và Cài đặt ở tiếng Anh bằng `dpi-check`/PaneShot (đặt `TROLY_LANG=en`).

Thêm chuỗi mới sau này: thêm cặp (tiếng Việt, tiếng Anh) vào `src\LangData.cs` (vào cả bảng Frag nếu là mảnh của thông báo ghép); bộ `tests-lang` sẽ báo nếu Ribbon, Cài đặt hoặc bảng điều khiển có chữ chưa dịch.

Phần của Rev0.9.5.14 (thẻ Công cụ) và Rev0.9.5.13 giữ nguyên bên dưới.

---

# Rev0.9.5.14 — Thẻ "Công cụ" xếp gọn, thay thẻ tự tạo của Word

1. **Thẻ mới "Công cụ"** (đứng đầu trong ba thẻ của add-in) gom các lệnh Word hay dùng vào một chỗ, xếp nút lớn nhỏ ngay hàng như thẻ có sẵn của Word. Thẻ tự tạo bằng Customize Ribbon của Word không chỉnh được kích thước nút hay chia hàng, nên bị lệch; thẻ nằm trong add-in thì tự quy định được.
   - **Chữ:** Format Painter lớn; hàng Font và cỡ chữ; hàng Đậm, Nghiêng, Gạch chân, Gạch ngang, Chỉ số dưới, Chỉ số trên; hàng màu chữ, tô nổi, đổi hoa thường, xóa định dạng.
   - **Đoạn văn:** nút lớn Chuẩn hóa đoạn văn; hàng gạch đầu dòng, đánh số, danh sách nhiều cấp, thụt dòng, hiện ký hiệu đoạn; hàng canh trái, giữa, phải, hai bên, giãn dòng; hộp thoại Paragraph.
   - **Trang:** Chuẩn hóa dàn trang lớn; Thêm trang dọc, Thêm trang ngang, Page Setup; Size, Orientation, Columns; Breaks, Page Borders (hộp Borders and Shading), Format Page Numbers.
   - **Chèn:** Symbol, Shapes, SmartArt; Caption, Cross-reference, Text Box; Equation lớn.
   - **Ảnh và xem:** Compress Pictures, Crop; One Page, Multiple Pages, Page Width.
   - **Bảng:** Table lớn; Shading, Borders, Text Direction; Distribute Rows, Distribute Columns, Repeat Header Rows; Split Cells, Split Table, Merge Cells; Convert Text to Table, AutoFit, Formula; lưới 9 nút canh chữ trong ô (trên, giữa, dưới nhân trái, giữa, phải).
2. Bốn nút của add-in trong thẻ này (Chuẩn hóa đoạn văn, Chuẩn hóa dàn trang, Thêm trang dọc, Thêm trang ngang) chạy cùng lệnh và dùng cùng icon với nút gốc ở thẻ Chuẩn hóa.
3. Thẻ có tên **Công cụ** để không trùng tên với thẻ **My tools** tự tạo của Word. Sau khi cài, bỏ dấu tích My tools trong File, Options, Customize Ribbon để khỏi hiện hai thẻ giống nhau.
4. Kiểm tra: 66 tên lệnh Word (idMso) trong Ribbon đều được xác nhận trên **Word thật** bằng `real-word-check.bat` (một tên sai làm Word bỏ cả Ribbon của add-in). Kiểm thử mới đảm bảo mọi tên lệnh nằm trong danh sách đã kiểm, không lặp lệnh, nhãn nhóm ngắn, bốn nút add-in nối đúng lệnh.
5. Chưa có trong thẻ: **Page Number (chèn số trang)** và **Columns dạng chọn nhanh một, hai, ba cột**, vì Word 16 trên máy này không có tên lệnh tương ứng; thay bằng Format Page Numbers và More Columns. **Crop Tools** dùng nút Crop. Chưa hiển thị thử qua Ribbon thật (cần cài bằng `install.bat`).

Phần của Rev0.9.5.13 (ô gộp, canh lề, nhiều đoạn) giữ nguyên bên dưới.

---

# Rev0.9.5.13 — Tính trong bảng có ô gộp, canh lề và nhiều đoạn

1. **Bảng có ô gộp:** Tổng, Trung bình, Nhỏ nhất, Lớn nhất, Đếm số và Tích theo định dạng nay tính đúng cột thẳng xuống dù hàng bên cạnh hoặc bên dưới có ô gộp. Nguyên nhân gốc: Word đánh số ô theo vị trí trong từng hàng, hàng gộp A:C chỉ có 3 ô nên ô D của hàng đó là ô thứ 3 chứ không phải thứ 5; mã cũ gọi `Cell(hàng, 5)` và thất bại nên báo "Không có ô số nào". Nay bảng không đều (nhận bằng `Table.Uniform` và số ô) dùng bản đồ ô theo độ rộng ngang: nguồn là các ô được chọn, kết quả vào ô thẳng cột ngay dưới ô cuối (hoặc ô ngay bên phải khi tính theo hàng), không ghi vào ô đã gộp. Hết bảng thì thêm hàng mới; nếu không thêm được cột thì báo rõ và không sửa bảng.
2. **Chuẩn hóa đoạn văn:** Cài đặt, mục Dàn trang, có thêm **Canh lề đoạn văn** (Hai bên, Trái, Giữa, Phải; mặc định Hai bên). Bôi đen nhiều đoạn thì áp cho **tất cả các đoạn đó cùng lúc** (bỏ qua tiêu đề và đoạn trong ô bảng; đoạn danh sách không thụt dòng đầu); không bôi đen thì chỉ đoạn tại con trỏ như bản 0.9.5.10. Vẫn tạm tắt AutomaticallyUpdate của từng style rồi khôi phục (cả khi lỗi).
3. Kiểm tra trên **Word thật** (phiên Word ẩn qua COM, `real-word-check.bat`): bảng gộp cho tổng 140, trung bình 35,00, nhỏ nhất 2, lớn nhất 70, đếm 4, Tích 2|20|48; bốn kiểu canh lề, nhiều đoạn, bỏ qua tiêu đề, khôi phục style, chế độ một đoạn. Phát hiện `Columns.Count` không báo lỗi trên Word 16 nên không dùng làm dấu hiệu.
4. Kiểm thử Word giả: mô phỏng ô gộp, thêm hàng, bộ sưu tập ô sống; kiểm thử mới trượt trên mã Rev0.9.5.12 (cả 5 phép tính rỗng), đạt trên mã mới.

Chưa kiểm tra qua nút Ribbon của bản cài.

---

## Rev0.9.5.12 (2026-11-18)

- Thẻ Phương trình có ba ô: nhập thô, xem trước Equation cùng dòng, xem trước Equation riêng dòng kiểu MathType với phân số nhiều tầng.
- Tích checkbox trong ô xem trước để chèn ngay kiểu đó vào Word. Ô tự bỏ tích sau thao tác để có thể chèn lại. Vùng bôi đen được giữ nguyên và công thức được chèn sau vùng đó.
- Build thành công, 12 bộ tự động đạt. Kiểm thử thẻ Phương trình có 15 ca, bắt được 3 điểm lệch trên bản cũ. Đã xem ảnh giao diện và thử thu phóng 50%, 100%, 130%. Đã thử qua COM trong Word thật với con trỏ giữa đoạn, xác nhận kiểu chèn, phân số đứng và Undo. Chưa thử thao tác qua Ribbon của bản cài mới.

## Rev0.9.5.11 (2026-11-18)

- Chỉnh thẻ Phương trình theo ảnh người dùng: ô chọn “Chèn công thức chuẩn MathType (phân số nhiều tầng)” đặt cùng khung xem trước. Tích chọn dùng phân số đứng khi bấm Chèn; bỏ chọn dùng phân số một dòng. Vẫn tạo Equation gốc của Word theo bộ chèn hiện có.
- Bỏ nút “Thay vùng chọn” và bỏ phần giải thích dưới công thức hợp lệ; vẫn hiện thông báo lỗi khi không nhận được công thức. Nút Chèn dùng vị trí con trỏ và không tự thay vùng bôi đen.
- Build thành công, 12 bộ tự động ALL PASS (180, 8654, 318, 272, 168, 131, 137, 11, 39, 446, 108, 23). Bộ mới TestComposePanel có 23 kiểm tra, bắt 9 lỗi trên giao diện cũ; kiểm tra ô chọn, một nút Chèn, thông báo, tạo công thức cả hai chế độ, không chồng ở bề rộng 300–700 và zoom 50%, 100%, 130%. Đã dựng và xem ảnh cửa sổ thật bằng Word mô phỏng. Chưa thử chèn qua bản cài trong Word thật.

# Rev0.9.5.10 — Chuẩn hóa đoạn tại con trỏ

Trong **Chuẩn hóa → Dàn trang**, nút **Chuẩn hóa đoạn văn** áp dụng cho một đoạn tại con trỏ.

- Căn đều hai bên (justify).
- Times New Roman dùng cỡ đã đặt (mặc định 13 pt), Arial dùng cỡ đã đặt (mặc định 11 pt). Với đoạn có font khác hoặc nhiều font, giữ cỡ chữ hiện có.
- Áp giãn dòng, khoảng cách trước/sau theo Cài đặt.
- Áp thụt đầu dòng hoặc thụt treo cho đoạn thường. Đoạn tiêu đề, danh sách và trong bảng giữ thụt lề hiện có.
- Khi có vùng chọn nhiều đoạn, xử lý một đoạn tại đầu đang hoạt động của vùng chọn.
- Một lần Ctrl+Z hoàn tác thao tác.

Đã thử trên Word thật qua COM với tài liệu 800 đoạn, Normal và style tùy chỉnh bật AutomaticallyUpdate, vùng chọn xuôi/ngược và đoạn trong bảng. Đã xác nhận định nghĩa style và đoạn lân cận giữ nguyên sau thao tác. Chưa thử nút Ribbon qua bản cài mới.

---

# Rev0.9.5.9 — Chuẩn hóa kỹ thuật theo ngôn ngữ và Công cụ AI

- **Chuẩn hóa kỹ thuật** mặc định dùng **Tự động theo ngôn ngữ quanh số**. Trong câu tiếng Anh, 12.5 là đúng và 12,5 được đề xuất sửa thành 12.5. Trong câu tiếng Việt, quy tắc đảo lại. Khi thiếu dấu hiệu từ chữ, dùng ngôn ngữ kiểm lỗi Word gán cho vùng gần số. Nếu vẫn chưa rõ, không ép đổi dấu. Có thể chọn kiểu Việt, Anh hoặc SI cố định.
- **Chuẩn hóa → Công cụ AI** trên Ribbon có các nút nhỏ **Thuật ngữ**, **Văn phong**, **Ngữ pháp**, **Từ đồng nghĩa**, **Viết lại**. Chức năng tra thuật ngữ nằm ở mục Thuật ngữ và Dịch nằm cùng Viết lại. Bảng điều khiển báo bằng chữ đỏ khi chưa có API key.
- Dấu −/+ và tỷ lệ được căn giữa. Các ô chọn Quét tăng chiều cao theo chữ ở nhiều mức thu/phóng.
- Đã thử quét và sửa số trong Word thật qua COM. Chưa thử bản cài qua Ribbon với API trực tuyến.

---

# Rev0.9.5.8 — Định dạng số trong bản dịch kỹ thuật

Khi bật **Dịch**, cả chế độ chỉ dịch và dịch kèm viết lại dùng dấu chấm thập phân cho tiếng Anh và dấu phẩy cho tiếng Việt. Ví dụ 12,50 MPa ↔ 12.50 MPa. Nhóm hàng nghìn có sẵn dùng khoảng trắng hẹp không ngắt. Add-in giữ dấu âm, chữ số 0 cuối, số mũ, ký hiệu đơn vị nhận diện được, số hiệu và trích dẫn. Kết quả được kiểm tra trước khi hiện bản xem trước và kiểm tra lại khi áp dụng vào Word. Nếu không xác định được ý nghĩa của dấu trong chuỗi mơ hồ như 1,234, add-in giữ nguyên chuỗi đó.

Quy ước trình bày số tham chiếu [SI Brochure, mục 5.4.4](https://www.bipm.org/documents/20126/41483022/SI-Brochure-9-EN.pdf). Phạm vi của lần sửa là định dạng số và kiểm tra dữ liệu kỹ thuật nhận diện được. Chất lượng dịch thuật ngữ và diễn đạt phụ thuộc model. Build và 11 bộ kiểm thử đạt. Đã thử áp dụng và hoàn tác cả hai chiều trong Word thật bằng bản dịch mô phỏng. Chưa thử API trực tuyến.

---

# Rev0.9.5.7 — Công cụ AI

Trong Word, mở **Chuẩn hóa → Công cụ AI**. Danh sách **Chọn công cụ AI** có ba mục.

- **Viết lại và Dịch**. Bôi đen tối đa 500 từ. Chọn một hoặc cả hai thao tác, chọn giọng văn hoặc ngôn ngữ đích, rồi chạy AI. Yêu cầu chỉ chứa chữ được bôi đen. Xem hoặc sửa bản xem trước rồi bấm **Áp dụng** để thay phần đã chọn. Add-in kiểm tra lại tài liệu và chữ gốc trước khi thay, giữ bố cục đoạn, đối chiếu số liệu và trích dẫn.
- **Văn phong**. Chọn các tiêu chí và quét phần bôi đen hoặc toàn bộ phần thân tài liệu. Có tìm câu dài, khẩu ngữ, cụm thừa, lặp từ, dấu câu, từ nối, xưng hô và gợi ý từ đồng nghĩa. **Sửa câu dài (AI)** áp dụng cho dòng được đánh dấu **Câu dài AI**.
- **Ngữ pháp**. Chọn tiếng Việt và/hoặc tiếng Anh, các tiêu chí chính tả, cấu trúc, hòa hợp, dấu câu, văn phong và thuật ngữ tiếng Anh. Quét phần bôi đen hoặc toàn bộ phần thân. Xem và áp dụng đề xuất trong danh sách. Nút **Sửa hết** chỉ xử lý các mục được phân loại là lỗi và có cách sửa hợp lệ.

**Thuật ngữ**, **Tra thuật ngữ** và **Gợi ý từ** dùng AI. Kho từ điển tích hợp và chức năng nạp CSV đã được gỡ. Bản mới không đọc các CSV cũ trong %AppData%\TroLyKyThuat.

Trong **Cài đặt**, nhập khóa API, bấm **Kiểm tra**, chọn model rồi bấm **Kiểm tra & Lưu**. Khi bấm quét toàn văn, add-in gửi toàn bộ phần thân chính, gồm chữ trong ô bảng, theo nhiều lô tới nhà cung cấp của model. Đầu/chân trang, chú thích và hộp văn bản nằm ngoài phạm vi quét. Model đánh giá theo từng lô.

Build thành công. Kiểm thử tự động 11 bộ. Đã thử thao tác đọc, ghi, hoàn tác và ánh xạ lỗi trong Word thật bằng tài liệu tạm và phản hồi mô phỏng. Chưa thử bản cài qua Ribbon với API trực tuyến.

---

# Rev0.9.5.6 — Công cụ AI (viết lại và dịch), giao diện, từ điển

Chưa thử trong Word thật. Sửa theo ảnh chụp người dùng gửi khi thử Rev0.9.5.3:

2. **Hai ô văn bản của thẻ mở rộng được**: kéo thanh nắm dưới mỗi ô, hoặc nhấp đúp để mở rộng gấp ba hoặc thu lại; chiều cao được nhớ.
3. **Nhóm "Vị trí chèn"** không còn bị chữ tiêu đề đè lên nút chọn (tự co theo cỡ chữ).
4. **Thu phóng xuống 50%** (trước đây 70%).
5. **Cài đặt**: nhãn "Kiểu lọc khi Tổng theo định dạng" đổi thành "Định dạng để tính tổng và tích"; mục Dàn trang bố trí lại bằng hàng nhãn và ô nhập tự co, không còn chồng chữ; thêm **Thụt dòng đầu (First line) 1,25 cm** và **Thụt treo (Hanging) 0 cm** cho nút Chuẩn hóa đoạn văn (không áp cho tiêu đề, đoạn trong bảng, đoạn danh sách; nếu cả hai > 0 thì ưu tiên thụt treo).
6. **Nút "Kiểm tra & Lưu" hiện đủ dấu &** (trước đây WinForms coi & là phím tắt nên mất ký tự); mọi nhãn và nút dùng GDI TextRenderer để chữ sắc nét hơn; hộp chọn màu vẽ chữ ClearType.
7. **Ngữ pháp**: bắt thêm lỗi modal + động từ sai dạng tổng quát (can improving, will reduces), "Anothers", và kiểm tra which + động từ theo tiền tố (these factors, which lead là đúng; properly, which lead nên là leads). Gỡ dòng cứng "which lead".

Kiểm thử Rev0.9.5.6 (15 bộ): 180, 8.654, 309, 251, 195, 131, 131, 294, 233, 98, 11, 125, 27, 446, 68: ALL PASS. Thử phá mã có chủ ý thì kiểm thử báo lỗi.

---

# Rev0.9.5.5 — Rà soát từ điển theo ngữ cảnh


# Rev0.9.5.4 — Từ điển Văn phong và Ngữ pháp

Từ điển Văn phong và Ngữ pháp Việt–Anh dùng chung cho hai thẻ. Mở bộ mặc định để xem danh sách tích hợp và mở tệp bổ sung để thêm hoặc ghi đè mục cùng ngôn ngữ, loại và cụm gốc. Quét đọc lại tệp bổ sung. Gợi ý văn phong và phương ngữ ở mức Nghi ngờ. Trong thẻ Văn phong, bật mục Khẩu ngữ để thấy gợi ý từ bộ này.

Chưa thử trong Word thật.

# Rev0.9.5.3 — Thu/phóng giao diện

Hai nút **−** và **+** ở đầu cửa sổ **Bảng điều khiển** và **Cài đặt** đổi tỷ lệ giao diện ngay. Mỗi lần bấm thay đổi 10 điểm phần trăm, từ 70% đến 130%. Mức mặc định là 100%. Add-in lưu tỷ lệ riêng cho từng cửa sổ. Tỷ lệ này áp dụng cho giao diện, còn mức zoom tài liệu Word do Word quản lý.

Build thành công và 15 bộ kiểm thử đạt, gồm 48 kiểm tra DPI và thu/phóng. Chưa thử trong Word thật.

**Rev0.9.5.2** (chưa thử trong Word thật) — Sửa hai lỗi người dùng báo khi thử Rev0.9.5.1:
- **Icon «Đọc số» sai ý nghĩa:** nút này viết số bằng chữ (123 thành một trăm hai mươi ba), không phải đọc thành tiếng. Icon cũ là cái loa; nay là số 123 ở trên, mũi tên xuống xanh, dòng chữ ở dưới.
- **Chữ trong Bảng điều khiển và hộp Cài đặt nhỏ, mờ, vỡ nét:** Windows có nhiều màn hình tỷ lệ khác nhau (máy này: 100% và 125%). Cửa sổ nhận biết DPI từng màn hình nhưng WinForms không tự co giãn khi cửa sổ sang màn hình khác tỷ lệ, nên chữ giữ cỡ 96 DPI (nhỏ, mảnh, thiếu nét). Nay form tự đổi cỡ chữ, kích thước và cột danh sách theo DPI thật của màn hình đang hiện, giữ trong vùng làm việc của màn hình. Nhật ký ghi thêm DPI và chế độ nhận biết DPI của cửa sổ để chẩn đoán; nếu Bảng điều khiển không ở chế độ nhận biết DPI từng màn hình, add-in tạo lại một lần.
- **Kiểm thử mới:** 	ests\TestDpi.cs (26 kiểm tra): gửi WM_DPICHANGED giả vào form thật và kiểm tra cỡ chữ, nút, danh sách co đúng 1,5 lần, về lại đúng cỡ cũ, bỏ qua DPI quá thấp, không vượt màn hình. Công cụ dpi-check.bat chụp Bảng điều khiển và hộp Cài đặt trên từng màn hình thật.
- Phiên bản 0.9.5.2; chỉ đổi icon Đọc số và co giãn theo DPI, không đổi chức năng khác.

**Rev0.9.5.1** (chưa thử trong Word thật) — Làm lại toàn bộ icon Ribbon theo phong cách Word hiện đại (phương án A):
- **Thay đổi giao diện icon:** Chuyển đổi toàn bộ icon từ kiểu ô vuông bo góc màu xanh đặc có chữ trắng sang phong cách phẳng, line art nét viền xám đậm `RGB(68,68,68)` dày 1,5–2 px kèm tối đa một màu nhấn (xanh dương `RGB(0,120,212)`, xanh lá `RGB(46,158,91)` cho thêm/xác nhận, đỏ `RGB(209,52,56)` cho xóa/lỗi).
- **Mỗi icon một ẩn dụ rõ ràng:**
  - Bảng: lưới nét xám với hàng tiêu đề xanh; áp dụng mọi bảng kèm huy hiệu ✓ xanh lá; áp dụng bảng chọn có khung nét đứt xanh; lặp tiêu đề có mũi tên xoay tròn.
  - Thêm/xóa hàng cột: hàng/cột mới tô xanh nhạt kèm huy hiệu "+" xanh lá đúng vị trí; hàng/cột xóa tô đỏ nhạt kèm huy hiệu "✕" đỏ.
  - Tính toán: khung tính bo góc nền xanh nhạt, ký hiệu toán nét đậm (Σ, x̄, min mũi tên xuống vạch sàn, max mũi tên lên vạch trần, #, π).
  - Định dạng số: số đại diện sắc nét "1,2", "1.2", "1 2" (bỏ chữ US/VN/SI); tăng/giảm số lẻ kèm dấu +/-; căn lề bảng trái/giữa/phải trên trang giấy nhỏ.
  - Màu sắc: icon riêng cho từng nút (màu chữ A có thanh màu, màu nền có bảng nhỏ và giọt màu, màu viền và dòng chẵn); swatch màu phẳng.
  - Ngôn ngữ & Dàn trang: Thuật ngữ (sách gáy xanh kèm tích), Tra thuật ngữ (kính lúp), Gợi ý từ (hai mũi tên hoán đổi), Ngữ pháp (trang giấy kèm sóng xanh), Dàn trang (trang có lề nét đứt), Cài đặt (bánh răng 8 răng có lỗ xanh), Bảng điều khiển (cửa sổ có thanh bên xanh).
  - Icon dự phòng: khung nét đứt xám có dấu "?" ở giữa cho mã lạ.
- **Icon chuyên biệt cho 6 nút trước đây dùng icon dự phòng:** `btnBdClr`, `btnHdrBg`, `btnHdrFg`, `btnSpDec`, `btnSpInc`, `btnZebraClr`.
- **Phiên bản & Tiêu đề:** `AssemblyInfo.cs` đặt 0.9.5.1; tiêu đề Bảng điều khiển hiển thị đầy đủ `Trợ lý Kỹ thuật - Rev0.9.5.1`.
- **Kiểm thử (14 bộ):** 180, 8.654, 309, 238, 195, 131, 131, 294, 233, 70, 11, 35, 27, và bộ mới 444 (TestIcons kiểm tra toàn diện 8 tiêu chí a–h): ALL PASS. Chỉ đổi giao diện icon, không đổi chức năng. Chưa thử trong Word thật.

---

**Rev0.9.5** (chưa thử trong Word thật) sửa các lỗi người dùng báo khi thử Rev0.9.3 và 0.9.4:
1. **Ô chọn màu không có tác dụng (lỗi nghiêm trọng):** khi bấm một ô màu, hộp chọn màu ghi kết quả OK rồi đóng, nhưng việc đóng làm phát sinh sự kiện mất tiêu điểm của chính hộp, và sự kiện đó đổi kết quả thành Hủy, nên màu vừa chọn bị bỏ. Đã sửa (không ghi đè khi đã chọn màu) và add-in nay tin vào màu đã chọn thay vì kết quả hộp. Có kiểm thử mô phỏng đúng cơ chế: trượt 3/11 trên mã Rev0.9.4, đạt 11/11 trên Rev0.9.5.
2. **Tách thẻ My tab thành hai thẻ:**
   - **My table:** Bảng, Ký tự trong bảng, Hàng và cột, **Tính trong bảng** (đổi tên; các nút nhỏ riêng lẻ: Tổng, Trung bình, Nhỏ nhất, Lớn nhất, Đếm số, Tổng theo định dạng, Tích theo định dạng). Đã bỏ xuất CSV và Excel. Khi con trỏ nằm trong một bảng, nhãn thẻ hiện "My table ●". Ribbon của Office không cho đổi màu chữ của nhãn thẻ, nên dùng dấu chấm tròn thay cho màu xanh.
   - **Chuẩn hóa:** Dàn trang, Định dạng số, Chuẩn hóa kỹ thuật, Ngôn ngữ, Cài đặt.
3. **Bỏ ô "Riêng một dòng"** ở thẻ Phương trình. Công thức luôn chèn cùng dòng; chỉ khi con trỏ ở đoạn hoàn toàn trống, hoặc bôi đen đúng cả một đoạn, thì chèn Equation riêng một dòng.
4. **AI không còn dịch sang tiếng Việt:** Viết lại và Sửa câu dài (AI) nay dùng lời nhắc tiếng Anh bắt buộc giữ nguyên ngôn ngữ của đoạn đầu vào ("NEVER translate"); nếu AI vẫn trả sai ngôn ngữ thì tự thử lại một lần, rồi từ chối áp dụng. Gợi ý đồng nghĩa AI cũng phải cùng ngôn ngữ với từ gốc.
5. **Thẻ Viết lại: chọn vị trí chèn** (Thay thế đoạn bôi đen, Chèn sau đoạn bôi đen, Chèn tại vị trí con trỏ), lưu qua các phiên, Hoàn tác đúng cho cả ba chế độ.
6. **Sửa câu dài (AI) nay hiện dấu ✔** trên dòng có đề xuất AI; bấm ✔ thay đúng câu trong Word và giữ định dạng ký tự.
7. **Tiêu đề Bảng điều khiển** hiện đúng số bản (trước đây ghi nhầm Rev0.9.3).
8. **Từ điển mở rộng** (`TextDataExt.cs`, do Claude biên soạn, chưa có chuyên gia ngôn ngữ duyệt): thêm 150+ nhóm đồng nghĩa tiếng Việt và 150+ nhóm tiếng Anh cho nút Gợi ý từ (đối xứng hai chiều, không bắc cầu giữa các nhóm), 100+ cặp chính tả và phương ngữ (nổ/nỗ, triển/triễn, s/x, ch/tr, d/gi/r, hỏi/ngã...), 70+ cụm thừa nghĩa. Gỡ một dòng lỗi cũ "xông pha → xông pha" có thể báo nhầm từ đúng. Người dùng bổ sung bằng `%AppData%\TroLyKyThuat\tu-dong-nghia.csv` và `tu-dien-nguoi-dung.csv`.

Kiểm thử Rev0.9.5 (13 bộ): 180, 8.654, 309, 238, 195, 131, 131, 294, 233, 70, 11 (hộp chọn màu), 35 (từ điển), 27 (Ribbon): ALL PASS. Chưa thử trong Word thật; cần kiểm tra Ribbon hai thẻ có hiện đúng, nhãn "My table ●" có đổi khi con trỏ vào và ra khỏi bảng.

---

**Rev0.9.3** (chưa thử trong Word thật) nâng cấp toàn diện từ Rev0.9.2 theo các nhóm **A–H**:
- **D (Đổi tên module):** nhóm Ribbon và thẻ Bảng điều khiển từ *Rà soát* sang **Chuẩn hóa kỹ thuật**.
- **A, B, C (Thao tác bảng & kiểu chuẩn trên Ribbon):** thêm nút căn trái/giữa/phải cho bảng (`AlignTables`); thay spin box cỡ viền bằng ComboBox chuẩn + nhập tay `0.25–6 pt`; thay ComboBox font bằng Gallery 4 font phổ biến kèm xem trước từng font và ComboBox cỡ chữ Word (`8–72 pt`); thay spin box `Thêm N` bằng ComboBox (`0–3` và nhập tay đến `20`).
- **F (Bảng chọn màu & Dàn trang):** thêm bảng chọn màu hợp nhất (`ColorPickerForm` với *Tự động*, *Màu giao diện 10×6*, *Màu chuẩn 10×1*, *Thêm màu...*) cho 4 nút màu bảng; thêm nhóm **Dàn trang** bên trái cùng thẻ với 4 nút: *Chuẩn hóa dàn trang* (A4, lề trên/dưới/trái/phải và đầu/chân trang cấu hình được), *Chuẩn hóa đoạn văn* (Times New Roman `13 pt`, Arial `11 pt`, giữ nguyên font khác, giãn dòng `1.2`, trước/sau `6 pt`), *Thêm 1 trang ngang* và *Thêm 1 trang dọc* (`wdSectionBreakNextPage`).
- **E (Văn phong, Ngữ pháp & Đồng nghĩa ngoại tuyến):** mở rộng kiểm tra chính tả/nhầm vùng miền (`ch/tr`, `s/x`, `d/gi/r`, `l/n`, hỏi/ngã, vần cuối), trùng lặp Hán-Việt, luật ghép vần âm tiết tiếng Việt (`ViSyllable`) có bảo vệ từ tiếng Anh/thuật ngữ/viết tắt/đơn vị, lỗi nhầm ngữ cảnh (`chong/trong`, `bảng/bản`, `mẫu/mẩu`, `phương/phần`, `giành/dành`, `bàng quan/bàng quang`, `chuyển/truyền`), mở rộng kho từ đồng nghĩa Việt/Anh (≥100 mục mỗi ngôn ngữ), thêm `tu-dien-nguoi-dung.csv`, thay ComboBox bằng `ListBox` bấm đúp/Enter thay trực tiếp và giữ nguyên định dạng ký tự khi thay từ.
- **G (Tính năng AI tùy chọn):** tự nhận diện nhà cung cấp qua tiền tố khóa API (`sk-ant-` Anthropic, `AIza` Gemini, `sk-` OpenAI), kiểm tra & lấy danh sách model trực tiếp, lưu khóa mã hóa DPAPI (tuyệt đối không lưu/in ra mã nguồn hay log), tùy chọn *Dùng AI* ở Văn phong (gửi câu dài kèm 2 câu trước/sau) và Ngữ pháp (gợi ý đồng nghĩa nâng cao), thêm thẻ **Viết lại** (giới hạn ≤500 từ, chọn giọng văn, bảo toàn định dạng đoạn và hỗ trợ hoàn tác), cùng chân trang thống kê token phiên làm việc.
- **H (DPI sắc nét):** bật Per-Monitor DPI Awareness V2 (`SetThreadDpiAwarenessContext` + `AutoScaleMode.Dpi`) cho Bảng điều khiển, Cài đặt và Bảng chọn màu.

---

**Rev0.9.2** (chưa thử trong Word thật) gộp nhánh 0.9.1D (Tích theo định dạng nhiều hàng, viền độc lập, các nút nhóm Bảng có tác dụng ngay) với dòng chính 0.9.7 (Thuật ngữ, Văn phong), và thêm **Ngữ pháp offline cho tiếng Việt và tiếng Anh** (nút Ngữ pháp trong nhóm Ngôn ngữ, thẻ Ngữ pháp trong Bảng điều khiển):
- **Tiếng Anh:** chính tả thường gặp (recieve, seperate, corrossion...), a/an có ngoại lệ phát âm (an MRI, a UT, an hour, a user), lặp từ, hòa hợp chủ vị chắc chắn, dấu câu và khoảng cách, văn phong (in order to, due to the fact that, từ yếu, câu dài, bị động lạm dụng), thống nhất Anh/Mỹ và cách viết thuật ngữ (pipeline, wall thickness...).
- **Tiếng Việt:** cặp liên từ sai (không những ... nhưng, mặc dù ... mà), dư từ (các những, rất nhất, một số các), lặp của/để/việc dày, chính tả thuật ngữ kỹ thuật (áp xuất, xác xuất, ứng xuất...), dấu câu và khoảng cách, câu bị cụt.
- Tự nhận ngôn ngữ từng đoạn. Bỏ qua tiêu đề, chữ trong ngoặc kép (tiếng Việt), trích dẫn [12]. Mức: Sai (chắc chắn, có sửa), Chưa chuẩn, Nghi ngờ (không tự sửa hàng loạt). Tùy chọn từng nhóm ở thẻ Ngữ pháp.
- **Giới hạn:** chỉ bắt lỗi có luật rõ ràng, ưu tiên không báo nhầm nên sẽ bỏ sót. Không hiểu nghĩa câu, không bắt lỗi ngữ pháp sâu hay logic; phần đó để mô-đun AI trực tuyến (Rev1.0). Bảng từ và luật là bộ khởi đầu chưa được chuyên gia ngôn ngữ duyệt.
- Kiểm thử: 157, 8.654, 309, 206, 156, 131, 122, 294 (ngữ pháp Anh), 196 (ngữ pháp Việt): ALL PASS. Thử trên đoạn văn đúng (Việt và Anh) không có báo nhầm; phá mã có chủ ý thì kiểm thử báo lỗi.

---

**Rev0.9.1D** (chưa thử trong Word thật):
1. **Sửa viền bảng:** màu, kiểu nét, độ dày viền nay áp dụng độc lập. Đổi màu viền chỉ ghi màu, không còn kéo theo kiểu nét và độ dày. Mọi bảng và Bảng này vẫn áp đủ cả ba.
2. **Tích theo định dạng nhiều hàng:** bôi đen một cột gồm nhiều ô (thường là cột cuối) rồi bấm nút: mỗi ô nhận tích của các ô đúng định dạng trong hàng của nó (trừ chính ô đó). Hàng không có ô khớp thì giữ nguyên ô đích. Các ô bôi đen phải cùng một cột. Một lần Ctrl+Z hoàn tác tất cả. Đặt con trỏ trong một ô vẫn tính cho một hàng như Rev0.9.1C.
Kiểm thử: 157, 8.654, 309, 206 (lệnh bảng), 139, 131: ALL PASS.

---

# Trợ lý Kỹ thuật, Rev0.9.2

**Rev0.9.1C thêm hai thay đổi** (chưa thử trong Word thật):
1. **Nút "Tích theo định dạng"** (menu Tính, cạnh "Tổng theo định dạng"): nhân các số đúng kiểu lọc định dạng (chữ đậm, tô sáng, màu chữ, màu nền ô, chọn trong Cài đặt) nằm trong **hàng chứa con trỏ**, **không tính ô con trỏ**, và ghi kết quả vào **chính ô con trỏ**. Không thêm hàng hoặc cột. Không ô nào khớp thì dừng, bảng không đổi. Bôi đen nhiều ô thì báo cần đặt con trỏ trong một ô. Hoàn tác bằng Ctrl+Z một bước. Số chữ số thập phân khi để Tự động bằng tổng số chữ số thập phân của các thừa số (1,5 x 1,5 = 2,25), tối đa 6.
2. **Các điều khiển trong nhóm Bảng và Ký tự trong bảng có tác dụng ngay** lên bảng đang chứa con trỏ (hoặc mọi bảng nằm trong phần bôi đen): màu nền, màu chữ, căn chữ tiêu đề; độ dày, màu, kiểu nét viền; tô xen kẽ và màu dòng chẵn; font, cỡ chữ, cách đoạn; căn ô chữ, căn ô số. Mỗi điều khiển chỉ ghi đúng nhóm thuộc tính của nó. Con trỏ không ở trong bảng thì chỉ lưu thông số. **"Mọi bảng" và "Bảng này"** vẫn rà soát toàn bộ thông số đang đặt và áp dụng một lần (cho mọi bảng hoặc bảng đang chọn).

Kiểm thử Rev0.9.1C: 157 lõi, 8.654 ma trận số, 309 hóa học, 170 lệnh bảng (137 cũ + 33 mới), 139 rà soát và Word giả, 131 đọc số: ALL PASS. Bộ thứ bảy (122 bộ máy văn bản) không có trong bản này vì mã Thuật ngữ và Văn phong chỉ có ở dòng chính (Rev0.9.5 trở đi). Chưa thử trong Word thật.

---

# Trợ lý Kỹ thuật, Rev0.9.1A (nhánh thử nghiệm từ Rev0.9)

Add-in Word chạy offline cho báo cáo kỹ thuật. **Rev0.9.1A thử nghiệm nút "Tổng theo định dạng"**: chỉ cộng các ô đậm (hoặc tô sáng, màu chữ, màu nền ô, chọn trong Cài đặt, mục Tính). Chưa thử trong Word thật.

Rev0.9 nâng nút Đọc số: đọc cả đơn vị đứng sau số (xem mục dưới). Rev0.8 sửa lỗi nghiêm trọng làm đổi giá trị số. Rev0.7 sửa theo phản hồi khi thử Rev0.6: bỏ hộp thoại sau mỗi lệnh, icon tông xanh da trời, thêm nút Đọc số, Đồng bộ số lẻ 0 đến 5 kèm kiểu kỹ thuật, bảng màu giống Word, danh sách font lấy từ máy, thẻ đổi tên thành "My tab".

## Cài đặt (làm một lần, hoặc khi nâng bản)

1. Đóng Word.
2. Bấm chuột phải vào `install.bat`, chọn Run as administrator.
3. Thấy "BIEN DICH THANH CONG" rồi "CAI DAT XONG" là được. Mở Word, thẻ "My tab" nằm ở cuối dãy thẻ.
4. Nếu lỗi: chạy `chan-doan.bat`, rồi báo Claude. Nhật ký ở `%TEMP%\TroLyKyThuat\log.txt`.

Gỡ: chạy `uninstall.bat` bằng quyền quản trị. Thử logic không cần Word: chạy `run-tests.bat`, kết quả mong đợi là sáu dòng "ALL PASS" (150, 8654, 309, 132, 139, 131 kiểm tra).

## Rev0.9: Đọc số kèm đơn vị

Đọc số nay đọc cả đơn vị đứng ngay sau số (cách tối đa một dấu cách), viết sau đơn vị, trong ngoặc:
- `20 kPa` thành `20 kPa (hai mươi kilôpascan)`
- `1.000.000 VNĐ` thành `1.000.000 VNĐ (một triệu Việt Nam đồng)`
- `12 m/s` thành `mười hai mét trên giây`, `9,81 m/s²` thành `chín phẩy tám một mét trên giây bình phương`, `50 kN·m`, `2,5 kg/m³`, `W/(m·K)`, `-15 °C` (âm mười lăm độ Celsius), `30 %` (phần trăm), `$100` (một trăm đô la Mỹ).

**Tên đơn vị** theo Nghị định 134/2007/NĐ-CP về đơn vị đo lường hợp pháp của Việt Nam: kilôgam, niutơn, pascan, jun, oát, vôn, ampe, culông, fara, ôm, simen, vebe, tesla, henry, héc, becơren, kenvin, độ Celsius; tiền tố kilô, mega, giga, tera, mili, micrô, nanô, picô... (kilôpascan, megapascan, kilôoát, micrômét). Số mũ: mét vuông, mét khối (đơn vị độ dài); giây bình phương, kilôgam trên mét khối (đơn vị khác); từ bậc 4 trở lên đọc "mũ bốn". Phép chia đọc "trên", phép nhân đọc cách nhau (niutơn mét).
- Nếu thích tên quen dùng (kilopascal, newton, pascal, watt, volt, độ C): Cài đặt, mục Số, "Tên đơn vị khi Đọc số", chọn "Thông dụng". Mặc định là theo luật.
- Có thêm đơn vị kỹ thuật dầu khí và cơ khí hay gặp: bar, psi, psig, barg, atm, mmHg, inch, foot, pound, thùng (bbl), bbl/d, MMscf/d, Nm³/h, ppm, ppb, tấn, kWh, MW, cSt, mD...; tiền VND, VNĐ, đ, ₫, USD, EUR, $, €.

**Giới hạn cần biết**: chỉ đọc đơn vị viết đúng chữ hoa chữ thường của ký hiệu (kPa, không phải kpa) và đứng ngay sau số; chữ thường không phải đơn vị ("25 máy", "3 người") được bỏ qua, chỉ đọc số. Bôi đen một phần thì chỉ đọc đơn vị nằm trong phần bôi đen. Con trỏ nằm trong số, trong đơn vị, hoặc ngay sau phần đọc đã viết đều được. Bảng tên đơn vị lấy từ nội dung Nghị định đã đọc trên mạng; chưa đối chiếu lại văn bản hiện hành (có thể đã được thay thế), nên cách viết tiền tố deci, centi và một số đơn vị ít gặp cần bạn xem lại.

## Sửa lỗi Rev0.8: số bị đổi giá trị

Tái hiện từ video: con trỏ trong `48,5`, bấm Thêm số lẻ vài lần, rồi Dấu, số thành `4.850.000`. Hai nguyên nhân chồng nhau:
1. **Vùng chọn sau khi sửa** (nút Dấu, Thêm số lẻ, Bớt số lẻ, Đồng bộ khi chỉ đặt con trỏ): bản Rev0.7 dựng lại vùng chọn thành đúng một ký tự giữa số (`48,5` thêm số lẻ thành `48,50`, vùng chọn là chữ `5`). Lần bấm sau sửa trúng ký tự đó nên số bị phá: `48,5,0`, `48,5,000`. Nay chỉ đặt con trỏ thì sau lệnh vẫn là con trỏ, nằm ngay cuối số vừa sửa. Bôi đen nhiều số vẫn giữ nguyên vùng bôi đen.
2. **Bộ đọc số quá dễ dãi**: `48,5,000` bị hiểu là 485000 vì dấu phẩy lặp được coi là dấu nghìn mà không kiểm tra cách nhóm. Nay dấu nhóm nghìn phải đúng quy tắc (nhóm đầu 1 đến 3 chữ số, các nhóm sau đúng 3 chữ số). Chuỗi sai quy tắc như `48,5,000`, `1.2,3`, `12,34,567` không phải số, nên mọi lệnh bỏ qua và không sửa.

Kiểm thử tái hiện đúng chuỗi thao tác trong video (đã xác nhận bản Rev0.7 trượt các kiểm thử này, bản Rev0.8 đạt).

## Thay đổi so với Rev0.6

1. **Bỏ hộp thoại sau mỗi lệnh**: kết quả chỉ hiện trên thanh trạng thái của Word (góc dưới bên trái), và trong Bảng điều khiển nếu đang mở. Muốn có lại hộp thoại: Cài đặt, mục Thông báo. Lỗi bất ngờ cũng đi qua thanh trạng thái và nhật ký.
2. **Icon tông xanh da trời** giống giao diện Colorful của Office 365 (xanh 0078D4 và các sắc gần nó). Đỏ chỉ còn ở nút xóa, vàng và cam chỉ là điểm nhấn nhỏ.
3. **Đọc số** (nút mới, nhóm Định dạng số): đặt con trỏ trong hoặc sát một số rồi bấm, chữ được viết ngay sau số, trong ngoặc: `1.234 (một nghìn hai trăm ba mươi bốn)`. Bôi đen nhiều số thì đọc từng số. Bấm lần hai không viết thêm.
   - Quy ước: hai mươi mốt, hai mươi lăm, mười lăm, một trăm linh năm, không trăm linh một (khi có hàng cao hơn), bốn (không dùng "tư"), số thập phân đọc từng chữ số sau "phẩy" (3,05 là ba phẩy không năm), số âm đọc "âm", đến 30 chữ số (nghìn tỷ, tỷ tỷ), `2,3 × 10⁶` đọc "hai phẩy ba nhân mười mũ sáu".
4. **Đồng bộ số lẻ 0 đến 5 chữ số** và **kiểu kỹ thuật**: mũi tên của nút có 0 đến 5 và mục "Viết kiểu kỹ thuật". Kiểu kỹ thuật viết `2300000` thành `2,3 × 10⁶`, `0,00367` thành `3,67 × 10⁻³` (dấu thập phân theo kiểu đang dùng). Số chữ số thập phân của phần đầu chỉnh ở Cài đặt, mục Số (0 đến 5, mặc định 2).
   - Số mũ là ký tự chỉ số trên của Unicode, không phải định dạng, nên số vừa viết không bị nhận lại thành một số khác. Số đã ở dạng `a × 10ⁿ` (cũng nhận `x`, `·`, `^`) được viết lại theo mức mới, không lồng thêm × 10. Đồng bộ số lẻ thường chỉ sửa phần đầu.
5. **Bảng màu giống Word** cho Màu nền, Màu chữ tiêu đề, Màu viền, Màu dòng chẵn: Tự động (hoặc Không màu), Màu giao diện (10 cột × 6 hàng: Nền 1, Chữ 1, Nền 2, Chữ 2, Nhấn 1 đến 6, kèm các hàng sáng hơn, tối hơn), Màu chuẩn (10 màu), Màu khác (hộp chọn màu của Windows).
   - Bộ màu giao diện lấy từ tài liệu đang mở (đọc lúc nạp Ribbon); không đọc được thì dùng bộ mặc định của Office hiện nay (Nhấn 1 là 156082).
   - Mã màu ghi vào tài liệu là RGB nên máy không cài add-in vẫn hiển thị đúng. Mỗi ô đọc được tên và mã màu (ví dụ "Nhấn 1, sáng hơn 80% (…)").
6. **Font trong bảng**: hộp Font nay liệt kê mọi font đã cài trong máy (cùng nguồn với danh sách của Word), vẫn gõ được tên font.
7. **Đổi tên thẻ** thành "My tab".

## Giới hạn đã biết

- Chưa thử trong Word thật. Kiểm thử bằng Word giả: 150 lõi, bảng màu và đọc số hợp lệ, 8.654 ma trận số, 309 hóa học, đơn vị, LaTeX, OMML khứ hồi, 132 lệnh bảng, 130 rà soát, đọc số, đồng bộ kiểu kỹ thuật, chuỗi thao tác bấm liên tiếp, 76 đọc số và kiểu kỹ thuật theo chuỗi.
- Bảng màu: màu gốc và hầu hết ô sáng hơn, tối hơn trùng công thức Office (HSL). Vài ô có thể lệch 1 đơn vị ở một kênh màu so với Word (cách làm tròn nội bộ khác nhau). Hàng biến thể của màu rất sáng, rất tối là phỏng theo cách Word đang hiển thị. Màu khác dùng hộp chọn màu chuẩn của Windows, không phải hộp "Colors" riêng của Word.
- Bộ màu giao diện chỉ đọc lại khi Ribbon nạp lại (mở Word, hoặc khi chọn màu). Đổi theme của tài liệu giữa chừng thì bảng màu chưa đổi ngay.
- Số kỹ thuật dùng ký tự chỉ số trên của Unicode: font không có đủ ký tự ⁴ đến ⁹ sẽ được Word tự thay font cho riêng các ký tự đó.
- Hai chỗ dễ sai nhất trong Word thật từ Rev0.6: (1) chèn Equation bằng InsertXML (nếu lỗi tự quay về định dạng chỉ số thường); (2) đổi Equation về chữ thường. Nếu lỗi, gửi lại dòng log.
- Ribbon của Word không có ô số dạng mũi tên chồng; dùng ô nhập kèm hai nút mũi tên lên, xuống.

## Các bản trước

### Rev0.6 (so với Rev0.5)

1. **Màu nền dòng tiêu đề không lan sang bảng khác**: "Bảng này" trước đây ghi cả vào Table Style dùng chung nên mọi bảng đã gán master style đổi theo. Nay "Bảng này" chỉ tạo style nếu chưa có và chỉ ghi vào bảng đang chọn; "Mọi bảng" vẫn cập nhật style.
2. **Lặp tiêu đề** (đổi tên): bật hoặc tắt cho mọi bảng nằm trong phần đang chọn (con trỏ trong một bảng thì bảng đó; bôi đen qua nhiều bảng thì tất cả). Bấm lần nữa để tắt. Mũi tên bên dưới: bật hoặc tắt cho mọi bảng.
3. **Giữ vùng bôi đen**: Dấu, Thêm số lẻ, Bớt số lẻ, Đồng bộ số lẻ đều đặt lại đúng vùng bôi đen sau khi sửa, nên bấm liên tiếp được. Vùng chọn rời nhau bằng Ctrl thì Word không cho dựng lại (sẽ chọn thừa ô), nên giữ nguyên cách cũ.
4. **Đồng bộ số lẻ** (nút mới, nhóm Định dạng số): mọi số đang bôi đen có đúng 1, 2 hoặc 3 chữ số thập phân. Số ít hơn được thêm số 0 (kể cả số nguyên), số nhiều hơn được làm tròn. Số ở góc biểu tượng là mức đang dùng; mũi tên bên dưới chọn 1, 2 hoặc 3 và áp dụng ngay.
5. **Bỏ qua tương tự, Sửa tương tự**: với số, "cùng dạng" là đủ (1.1 và 2.3; 5.678 và 12.345). Mỗi số vẫn được sửa theo giá trị riêng.
6. **Kiểu chỉ số**: luân phiên hai kiểu: Equation và font thường có chỉ số trên, dưới. Equation có sẵn được đọc thẳng từ OMML (giữ chỉ số dưới thẳng, chữ nghiêng) rồi đổi lại thành chữ thường.
7. **Ribbon**:
   - Nhóm Bảng: Tiêu đề (màu nền, màu chữ, căn chữ), Viền (độ dày với mũi tên lên xuống, màu, kiểu nét), Tô xen kẽ (hộp chọn, màu). Mỗi nút màu có 8 màu gợi ý và Màu khác.
   - Nhóm Ký tự trong bảng: Cỡ chữ và Cách đoạn có ô nhập cùng mũi tên lên, xuống.
   - Biểu tượng Rà soát đổi thành kính lúp; biểu tượng Căn ô chữ, Căn ô số rõ hơn.
8. **Ô nhập phương trình** trong Bảng điều khiển: đổi font ô nhập (Cambria Math làm chữ bị cắt, nhìn như mất chữ).
9. **Nhận diện**: mã LaTeX dính đơn vị như `Ω\cdotcm²` nay được nhận ra.

