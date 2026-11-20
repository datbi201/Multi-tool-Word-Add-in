# Lịch sử các bản (Rev)

Mỗi bản nằm trong một thư mục riêng, ví dụ `Rev0.1`. Bản mới không ghi đè bản cũ.

## Rev0.9.6 (2026-11-20)

- Thêm giao diện tiếng Anh và nút đổi ngôn ngữ Việt, Anh (Ribbon, Cài đặt, bảng điều khiển, hộp chọn màu, thông báo kết quả). Đổi ngay không cần mở lại Word. Mặc định tiếng Việt.
- Lớp `Lang` dịch ở biên giao diện: Ribbon dùng lệnh gọi lại thay nhãn tĩnh, hộp thoại dịch từ bản gốc tiếng Việt, thông báo ghép dịch theo mảnh (không khớp giữa một từ khác). Chuỗi chưa dịch giữ tiếng Việt.
- Chữ giữ tiếng Việt có chủ ý: đọc số bằng lời, luật rà soát văn bản tiếng Việt, kết quả AI.
- Hướng dẫn nhanh trong Cài đặt cập nhật theo ba thẻ hiện tại.
- Kiểm thử mới `tests-lang` (67 kiểm tra); 13 bộ đạt. Chưa thử qua Ribbon của bản cài.

## Rev0.9.5.14 (2026-11-20)

- Thêm thẻ **Công cụ** vào add-in, dựng từ các lệnh Word có sẵn (idMso), thay thẻ My tools tự tạo bằng Customize Ribbon bị lệch (Word không cho chỉnh kích thước nút và chia hàng). Sáu nhóm: Chữ, Đoạn văn, Trang, Chèn, Ảnh và xem, Bảng; kèm bốn nút của add-in (Chuẩn hóa đoạn văn, Chuẩn hóa dàn trang, Thêm trang dọc, Thêm trang ngang).
- Đã xác nhận 66 tên lệnh trên Word 16 thật (công cụ `tools\RealWordMso.cs`). Không có tên lệnh cho Page Number và Columns nhanh trên máy này, nên dùng Format Page Numbers và More Columns.
- Kiểm thử mới trong TestRibbon (65 kiểm tra) và bí danh icon trong TestIcons; 12 bộ kiểm thử đạt. Chưa thử qua Ribbon của bản cài.

## Rev0.9.5.13 (2026-11-20)

- Sửa tính toán trong bảng có ô gộp (người dùng báo ngày 20/11, ảnh bảng có hàng 5 gộp): Word đánh số ô theo vị trí trong hàng nên mã cũ không tìm được ô D của hàng gộp. Thêm bản đồ ô theo độ rộng cho bảng không đều; áp cho Tổng, Trung bình, Nhỏ nhất, Lớn nhất, Đếm số, Tổng và Tích theo định dạng.
- Chuẩn hóa đoạn văn: thêm Canh lề (hai bên, trái, giữa, phải) trong Cài đặt; bôi đen nhiều đoạn thì áp cho tất cả cùng lúc (đảo lại quy định Rev0.9.5.10 chỉ xử lý đoạn đầu hoạt động).
- Kiểm tra trên Word thật qua COM; kiểm thử Word giả cập nhật; kiểm thử ô gộp trượt trên mã cũ.
- Mô hình Word giả từng giả định sai `Columns.Count` báo lỗi với ô gộp; đã sửa theo kết quả đo trên Word thật.

## Rev0.9.5.12 (2026-11-18)

- Thẻ Phương trình có ba ô: nhập thô, xem trước Equation cùng dòng, xem trước Equation riêng dòng kiểu MathType với phân số nhiều tầng.
- Tích checkbox trong ô xem trước để chèn ngay kiểu đó vào Word. Ô tự bỏ tích sau thao tác để có thể chèn lại. Vùng bôi đen được giữ nguyên và công thức được chèn sau vùng đó.
- Build thành công, 12 bộ tự động đạt. Kiểm thử thẻ Phương trình có 15 ca, bắt được 3 điểm lệch trên bản cũ. Đã xem ảnh giao diện và thử thu phóng 50%, 100%, 130%. Đã thử qua COM trong Word thật với con trỏ giữa đoạn, xác nhận kiểu chèn, phân số đứng và Undo. Chưa thử thao tác qua Ribbon của bản cài mới.

## Rev0.9.5.11 (2026-11-18)

- Chỉnh thẻ Phương trình theo ảnh người dùng: ô chọn “Chèn công thức chuẩn MathType (phân số nhiều tầng)” đặt cùng khung xem trước. Tích chọn dùng phân số đứng khi bấm Chèn; bỏ chọn dùng phân số một dòng. Vẫn tạo Equation gốc của Word theo bộ chèn hiện có.
- Bỏ nút “Thay vùng chọn” và bỏ phần giải thích dưới công thức hợp lệ; vẫn hiện thông báo lỗi khi không nhận được công thức. Nút Chèn dùng vị trí con trỏ và không tự thay vùng bôi đen.
- Build thành công, 12 bộ tự động ALL PASS (180, 8654, 318, 272, 168, 131, 137, 11, 39, 446, 108, 23). Bộ mới TestComposePanel có 23 kiểm tra, bắt 9 lỗi trên giao diện cũ; kiểm tra ô chọn, một nút Chèn, thông báo, tạo công thức cả hai chế độ, không chồng ở bề rộng 300–700 và zoom 50%, 100%, 130%. Đã dựng và xem ảnh cửa sổ thật bằng Word mô phỏng. Chưa thử chèn qua bản cài trong Word thật.

## Rev0.9.5.10 (2026-11-18)

- Chuẩn hóa đoạn văn chỉ xử lý đoạn tại con trỏ, căn đều hai bên và áp thông số cỡ chữ, giãn dòng, khoảng cách, thụt lề hiện có. Vùng chọn nhiều đoạn chỉ xử lý đoạn tại đầu hoạt động.
- Gỡ vòng duyệt toàn văn của thao tác này. Định dạng trực tiếp, tạm ngừng AutomaticallyUpdate của style nếu bật và khôi phục cả khi lỗi.
- Kiểm thử tái hiện trên mã cũ phát hiện 4 lỗi. Build thành công, 11 bộ tự động đạt. Word thật qua COM xác nhận Normal, đoạn lân cận, nội dung, vị trí chọn và Undo trên tài liệu 800 đoạn, cả style tự cập nhật. Chưa thử nút Ribbon qua bản cài mới.

## Rev0.9.5.9 (2026-11-14)

- Căn giữa dấu −/+ và tỷ lệ thu/phóng. Nhóm tùy chọn Quét tăng chiều cao theo chữ, kiểm tra từ 50% đến 130% ở DPI 144 và 192.
- Chuẩn hóa kỹ thuật nhận ngôn ngữ theo ngữ cảnh từng số, kể cả các vế Việt–Anh trong cùng đoạn. Dùng ngôn ngữ kiểm lỗi Word khi chữ quanh số chưa đủ rõ và giữ tùy chọn dấu cố định. Bảo vệ số nhóm nghìn có khoảng trắng SI.
- Đổi nhóm Ngôn ngữ trên Ribbon thành Công cụ AI, với năm nút nhỏ. Gộp Thuật ngữ và Từ đồng nghĩa vào cùng thẻ AI trong Bảng điều khiển. Hiện hướng dẫn màu đỏ khi thiếu API key.
- Thử quét, nhận ngôn ngữ Word và sửa vùng chọn trên Word thật qua COM. Chưa thử bản cài qua Ribbon với API trực tuyến.

## Rev0.9.5.8 (2026-11-14)

- Chuẩn hóa dấu thập phân cho bản dịch sang tiếng Anh và tiếng Việt, giữ giá trị, số chữ số thập phân, dấu và số mũ. Nhóm nghìn dùng khoảng trắng hẹp không ngắt khi nguồn đã nhóm.
- Đối chiếu trước khi hiện bản xem trước và trước khi ghi vào Word. Bảo vệ ký hiệu đơn vị nhận diện được, số hiệu tiêu chuẩn, trích dẫn và số điều khoản. Giữ cách viết gốc nếu số có thể hiểu theo hai cách và thiếu bằng chứng chọn quy ước.
- Build thành công, 11 bộ tự động đạt, bộ AI có 137 kiểm tra. Thử áp dụng, chặn sai đơn vị và hoàn tác ở hai chiều trên Word thật qua COM với văn bản mô phỏng. Chưa thử bản cài qua Ribbon với API trực tuyến.

## Rev0.9.5.7 (2026-11-14)

- Gộp Văn phong, Ngữ pháp, Viết lại và Dịch vào thẻ Công cụ AI. Quét phần bôi đen hoặc toàn bộ phần thân chính, gồm chữ trong bảng, qua model đã cấu hình. Viết lại và Dịch chỉ gửi và thay vùng đã bôi đen.
- Gỡ từ điển tích hợp và chức năng nạp CSV. Thuật ngữ, Tra thuật ngữ và Gợi ý từ chuyển sang AI.
- Đối chiếu JSON, vị trí trích dẫn, nội dung cũ và trạng thái tài liệu. Từ chối đầu ra bị cắt và thay đổi số liệu hoặc trích dẫn đã phát hiện. Gợi ý ngữ cảnh duyệt từng mục.

## Rev0.1 (2026-11-07)
- Khung add-in Word (COM, late binding, build bằng csc có sẵn của Windows, không cần Visual Studio).
- Module Bảng: master style (Table Style TroLy_Master) áp dụng cho mọi bảng; công cụ cho bảng đang chọn (thêm hàng cột, xóa, tính tổng và trung bình điền một lần, quy ước số, bố cục, tô nền, xuất CSV).
- Kiểm thử: lõi tính toán (68 kiểm tra) và mô hình Word giả (36 kiểm tra) đều đạt.
- Chưa thử trong Word thật.

## Rev0.1.1 (2026-11-07)
- Sửa lỗi Word báo "running into problems with the add-in" rồi đóng ngay: khai báo giao diện Ribbon (IRibbonExtensibility) sai kiểu, làm Word gọi nhầm hàm.
- Thêm tệp nhật ký `%APPDATA%\TroLyKyThuat\log.txt` và bắt mọi lỗi trên luồng của bảng điều khiển.
- install.bat xóa danh sách add-in bị Word vô hiệu hóa.
- Chưa thử trong Word thật.

## Rev0.1.2 (2026-11-07)
- Word vẫn đóng khi nạp ở Rev0.1.1, và không có nhật ký, nghĩa là lỗi xảy ra trước khi mã add-in chạy. Rev0.1.1 đã đổi giao diện Ribbon sang kiểu sai (bản PIA của Microsoft là giao diện kép), nên Rev0.1.2 trả về đúng kiểu.
- Khai báo IDTExtensibility2 và IRibbonExtensibility giống hệt bản PIA (TypeLibType 4160, có DispId).
- Nhật ký ghi sớm nhất có thể (ngay khi DLL được nạp) và ghi vào %TEMP%\TroLyKyThuat\log.txt, nơi luôn ghi được.
- Chưa thử trong Word thật.
- Kết quả chẩn đoán trên máy bạn: Word 64-bit (Office 365, 16.0.20326), .NET 4.8, đăng ký COM đúng. Word sập trong clr.dll (lỗi nội bộ .NET, mã 0x80131506, truy cập bộ nhớ sai) ngay lúc nạp, trước khi mã add-in ghi được dòng nhật ký nào.
- Rev0.1.2 bỏ hẳn lớp chuyển đổi kiểu của CLR cho các tham số của IDTExtensibility2 (dùng con trỏ thô IntPtr), các hàm của add-in cài đặt dạng tường minh.
- install.bat gỡ đăng ký bản cũ trước khi cài.

## Rev0.1.3 (2026-11-07)
- Add-in đã nạp được trong Word thật (Rev0.1.2). Lần thử đầu báo các lỗi của module Bảng, đã sửa:
- Màu nền, màu chữ dòng tiêu đề, cỡ chữ và tô xen kẽ không hiện: nguyên nhân là các thiết lập nằm trong Table Style bị đè bởi Paragraph Style Normal và bởi nền tự động ghi trực tiếp. Nay định dạng trực tiếp từng ô.
- Kiểu viền Đứt hiện nét đôi: đã dùng nhầm mã 15 (ThinThickLargeGap), nay dùng mã 3 (DashSmallGap).
- Thêm hàng, cột nay xử lý riêng từng hàng, cột đã chọn (kể cả chọn rời nhau bằng Ctrl nếu Word cho phép), thêm ô nhập số hàng, cột. Xóa hàng, cột cũng theo đúng các hàng, cột đã chọn.
- Xóa định dạng chỉnh tay không còn làm mất chỉ số trên và dưới.
- Kiểm thử: 68 kiểm tra lõi, 71 kiểm tra mô hình Word giả, đều đạt. Chưa thử trong Word thật.


## Rev0.2 (2026-11-07)
- Ribbon có đủ lệnh hay dùng (áp dụng master style, thêm, xóa hàng cột, tính, định dạng số, rà soát). Bảng điều khiển chỉ còn để chỉnh thông số và duyệt danh sách rà soát.
- Tốc độ: mọi lệnh Word chạy trên luồng chính của Word; đọc rồi so sánh, chỉ ghi thuộc tính khác; tạm tắt cập nhật màn hình, phân trang nền, kiểm tra chính tả. Thông báo kèm thời gian chạy.
- Sửa định dạng số: chụp vùng chọn thành tọa độ trước khi ghi; cả ba kiểu nhóm nghìn từ 4 chữ số; SI dùng dấu cách không ngắt; số mơ hồ (1,234) quyết định bằng bằng chứng từ các số rõ ràng trong vùng chọn; tính bằng decimal.
- Bỏ các nút Chọn nhanh cả hàng, cột, bảng và Làm mới vùng chọn. Chọn theo số (2,4,6-8) dùng cho cả thêm và xóa. Ghi log cách Word báo vùng chọn rời nhau (dòng SEL).
- Tab Chuẩn hóa: rà soát toàn văn bản (hóa học, toán, ký hiệu theo ngữ cảnh, số kỹ thuật), Sửa, Bỏ qua, Sửa tất cả, equation gốc của Word hoặc chỉ số trên dưới.
- Kiểm thử: 71 lõi, 8.654 ma trận số, 57 hóa học và nhận diện, 81 lệnh bảng với Word giả, 18 rà soát với Word giả. Chưa thử trong Word thật.

## Rev0.3 (2026-11-08)
- Chọn nhiều hàng, cột rời nhau bằng Ctrl: bộ đếm giờ gom các vùng chọn (Word chỉ báo vùng cuối). Thêm cột, hàng, xóa, tô nền, tính đều dùng tập ô đã gom. Log SEL có thêm gomCtrl.
- Chuẩn hóa: nhận H_2O, O_2, SO_4^{2-}, O^{2-} là hóa học (không còn là biến); nhận diện số bằng ngữ cảnh (năm, ngày, số mục, trích dẫn, số hiệu), bỏ qua số 4 chữ số không manh mối; đại lượng in nghiêng, công thức hóa học đứng thẳng; tô vàng mục đang xem và trả lại màu cũ; công thức riêng một dòng luôn dùng Equation; công thức cùng dòng chọn "Định dạng chỉ số trên dưới" hoặc "Dùng font chữ nhỏ và thấp" (glyph Cambria).
- Ribbon: thẻ ở cuối dãy; icon vẽ bằng mã; thêm nút Lặp tiêu đề, Tô nền, Bỏ tô nền, Xuất CSV, Sao chép, Soạn công thức, ô Chọn theo số.
- Bảng điều khiển: bỏ nút trùng Ribbon, nội dung giãn theo cửa sổ, tab Soạn công thức mới, tùy chọn gắn sát cạnh phải Word (mô phỏng).
- Soạn công thức: hóa học (cân bằng tùy chọn) và toán, xem trước, chèn tại con trỏ hoặc thay vùng bôi đen, chín mẫu.
- Nghiên cứu task pane thật: cần ICustomTaskPaneConsumer và điều khiển ActiveX đăng ký COM; rủi ro cao và ảnh hưởng trình đọc màn hình nên hoãn.
- Kiểm thử: 71 lõi, 8.654 ma trận số, 113 hóa học, nhận diện, soạn công thức, 89 lệnh bảng, 44 rà soát và chèn công thức. Chưa thử trong Word thật.

## Rev0.4 (2026-11-08)
- Bỏ gắn sát cạnh phải cửa sổ Word. Sửa tính theo cột (hướng tự nhận, bỏ lựa chọn hướng). Bỏ tô nền ô.
- Mô-đun Định dạng số (Ribbon): nút Dấu luân chuyển Việt, Mỹ, SI; Thêm số lẻ, Bớt số lẻ (tối đa 5); làm việc trên dãy ô trong bảng hoặc các số trong văn bản, nhiều số cùng lúc.
- Chuẩn hóa: từ điển đơn vị kỹ thuật (rpm, kPa, MPa, ..., hoa thường sai thì báo), mã LaTeX hiện lộ (\Delta, \tau_w, \Delta G_{ads}°), P_CO2 (cả cụm CO2 là chỉ số dưới), phát hiện công thức và phương trình hóa học sai hóa trị kèm đề xuất sửa và cân bằng; mặc định chọn hết kiểu quét.
- Nút Kiểu chỉ số: tự tìm công thức trong vùng bôi đen và luân chuyển chỉ số định dạng, ký tự Cambria Math, Equation cùng dòng.
- Soạn phương trình (đổi tên, bỏ mẫu, tự nhận loại, xem trước đúng chỉ số, hộp cách hiểu và sửa, chèn tại con trỏ).
- Ribbon và Bảng điều khiển: ít chữ, nhiều biểu tượng, chú thích chuyển thành tooltip; bỏ chức năng trùng; thẻ Cài đặt có quy ước số, nhà cung cấp AI, khóa API mã hóa DPAPI và Trợ giúp.
- Kiểm thử: 73 lõi, 8.654 ma trận số, 230 hóa học, đơn vị và phương trình, 108 lệnh bảng, 66 rà soát và định dạng. Chưa thử trong Word thật.

## Rev0.5 (2026-11-08)
- Nút Chuẩn hóa nhanh: sửa mã lỗi tại con trỏ hoặc vùng bôi đen (R_cyl, R_{cyl}, t_req, \tau_w, Re_{cyl}); công thức dài thành Equation Word.
- Thư viện: LaTeX lộ (\frac, mũ âm, chữ Hy Lạp), Re/Nu/Pr thẳng, số mục tiêu đề, số hiệu Bảng và Hình, khoảng cách số và đơn vị.
- Equation: bỏ ("corr"), hóa học chữ thẳng không nháy, phân số đứng kiểu MathType (tùy chọn), trả con trỏ về Word sau khi chèn.
- Ribbon: thêm Ký tự trong bảng, menu Dòng tiêu đề, Viền, Tô xen kẽ, nút tăng giảm số lượng, nhóm Xuất bên phải, Cài đặt thành hộp thoại riêng (kèm Xóa định dạng chỉnh tay).
- Bảng điều khiển: hai thẻ, kiểu dấu VN/US/SI độc lập, cột ✔ ✖, Sửa tương tự, Bỏ qua tương tự.
- Kiểm thử: 90 lõi, 8.654 ma trận số, 276 hóa học và LaTeX, 108 lệnh bảng, 81 rà soát. Chưa thử InsertXML trong Word thật.

## Rev0.6 (2026-11-08)
- Sửa: "Bảng này" không còn đổi màu nền dòng tiêu đề của các bảng khác (không ghi vào Table Style dùng chung).
- Lặp tiêu đề: làm được cho nhiều bảng (phần bôi đen hoặc mọi bảng), bấm lần nữa để tắt.
- Dấu, Thêm số lẻ, Bớt số lẻ giữ nguyên vùng bôi đen sau khi sửa.
- Nút mới Đồng bộ số lẻ (1, 2 hoặc 3 chữ số thập phân cho mọi số đang bôi đen).
- Bỏ qua tương tự, Sửa tương tự nhận ra số đánh số (cùng dạng).
- Kiểu chỉ số luân phiên Equation và font thường; đọc ngược Equation từ OMML.
- Ribbon: Tiêu đề, Viền, Tô xen kẽ dạng điều khiển nhỏ; Cỡ chữ, Cách đoạn, Viền có mũi tên lên xuống; Rà soát biểu tượng kính lúp.
- Ô nhập phương trình hết bị mất chữ. Nhận ra Ω\cdotcm².
- Kiểm thử: 105 lõi, 8.654 ma trận số, 309 hóa học và OMML, 132 lệnh bảng, 88 rà soát. Chưa thử trong Word thật.


## Rev0.7 (2026-11-08)
- Bỏ hộp thoại sau mỗi lệnh: kết quả chỉ ghi lên thanh trạng thái Word (bật lại ở Cài đặt, mục Thông báo).
- Icon đổi sang tông xanh da trời kiểu Office 365 Colorful.
- Nút mới Đọc số: viết số bằng chữ tiếng Việt ngay sau số, trong ngoặc (đến 30 chữ số, thập phân, số âm, số kỹ thuật).
- Đồng bộ số lẻ 0 đến 5 chữ số, thêm kiểu kỹ thuật (2,3 × 10⁶) với số lẻ chỉnh ở Cài đặt; nhận lại và viết lại số đã ở dạng a × 10ⁿ.
- Bảng màu giống Word cho bốn nút màu: Tự động, Màu giao diện 10 × 6 (lấy theo theme tài liệu), Màu chuẩn, Màu khác.
- Hộp Font liệt kê các font đã cài trong máy. Thẻ đổi tên thành "My tab".
- Kiểm thử: 130 lõi và bảng màu, 8.654 ma trận số, 309 hóa học và OMML, 132 lệnh bảng, 111 rà soát, đọc số, kiểu kỹ thuật, 76 đọc số và kiểu kỹ thuật theo chuỗi. Chưa thử trong Word thật.

## Rev0.8 (2026-11-08)
- Sửa lỗi nghiêm trọng làm đổi giá trị số (video 16:08): con trỏ trong 48,5, bấm Thêm số lẻ nhiều lần rồi Dấu thành 4.850.000.
  - Nguyên nhân 1: sau lệnh định dạng số khi chỉ đặt con trỏ, vùng chọn bị dựng thành một ký tự giữa số; lần bấm sau sửa trúng ký tự đó. Nay giữ con trỏ, đặt cuối số.
  - Nguyên nhân 2: bộ đọc số nhận 48,5,000 là 485000 (không kiểm tra nhóm nghìn). Nay dấu nhóm phải đúng quy tắc 1 đến 3 chữ số rồi từng nhóm 3 chữ số; chuỗi sai không phải số và không bị sửa.
- Kiểm thử mới tái hiện đúng chuỗi trong video; bản Rev0.7 trượt, Rev0.8 đạt. 150 lõi, 8.654 ma trận số, 309 hóa học và OMML, 132 lệnh bảng, 130 rà soát và Word giả, 76 đọc số. Chưa thử trong Word thật.

## Rev0.9 (2026-11-08)
- Đọc số nay đọc cả đơn vị đứng ngay sau số, viết sau đơn vị trong ngoặc: `20 kPa (hai mươi kilôpascan)`, `1.000.000 VNĐ (một triệu Việt Nam đồng)`, `12 m/s (mười hai mét trên giây)`, `$100 (một trăm đô la Mỹ)`.
- Tên đơn vị theo Nghị định 134/2007/NĐ-CP (kilôgam, niutơn, pascan, oát, vôn, độ Celsius, tiền tố kilô, mega, micrô...). Cài đặt, mục Số, có lựa chọn "Tên đơn vị khi Đọc số": theo luật (mặc định) hoặc thông dụng (kilopascal, newton, watt).
- Nhận đơn vị đơn, có số mũ (m², m³, s⁻¹, m3), đơn vị ghép (m/s, N·m, kg/m³, W/(m·K)), phần trăm, độ C, tiền tệ, đơn vị dầu khí (bar, psi, bbl/d, MMscf/d, Nm³/h...). Chữ thường không phải đơn vị ("25 máy") bị bỏ qua.
- Con trỏ trong số, trong đơn vị hoặc ngay sau phần đọc đều được; bấm lại không viết thêm. Bôi đen một phần thì chỉ đọc đơn vị nằm trong phần bôi đen.
- Chưa đối chiếu lại hiệu lực hiện hành của Nghị định 134/2007/NĐ-CP; cách viết tiền tố deci, centi và đơn vị ít gặp cần xem lại. Chưa thử trong Word thật.
- Kiểm thử: 150 lõi, 8.654 ma trận số, 309 hóa học và OMML, 132 lệnh bảng, 139 rà soát và Word giả, 131 đọc số và đơn vị.

## Rev0.9.2 (2026-11-12)
- Gộp nhánh 0.9.1D với dòng chính 0.9.7 bằng ghép ba chiều (không xung đột đáng kể): có cả Tích nhiều hàng, viền độc lập, nút áp dụng ngay và Thuật ngữ, Văn phong.
- Thêm Ngữ pháp offline tiếng Việt và tiếng Anh (Grammar.cs, GrammarEn.cs, GrammarVi.cs), nút Ngữ pháp, thẻ Ngữ pháp, cài đặt GrVi, GrEn, GrSpell, GrAgree, GrPunc, GrStyle, GrTerms, GrLongN.
- Kiểm thử: chín bộ ALL PASS (157, 8.654, 309, 206, 156, 131, 122, 294, 196). Chưa thử trong Word thật.

## Rev0.9.5.6 (2026-11-14)
- Cài đặt: đổi nhãn "Định dạng để tính tổng và tích"; bố trí lại mục Dàn trang không chồng chữ; thêm Thụt dòng đầu và Thụt treo cho Chuẩn hóa đoạn văn; nút "Kiểm tra & Lưu" hiện đủ dấu &; chữ GDI và ClearType sắc nét hơn.
- Kiểm thử: 15 bộ ALL PASS (180, 8.654, 309, 251, 195, 131, 131, 294, 233, 98, 11, 125, 27, 446, 68). Chưa thử trong Word thật.

## Rev0.9.5.5 (2026-11-14)

## Rev0.9.5.4 (2026-11-14)
- Từ điển Văn phong và Ngữ pháp Việt–Anh dùng chung cho hai thẻ. Mở bộ mặc định để xem danh sách tích hợp và mở tệp bổ sung để thêm hoặc ghi đè mục cùng ngôn ngữ, loại và cụm gốc. Quét đọc lại tệp bổ sung. Gợi ý văn phong và phương ngữ ở mức Nghi ngờ.
- Kiểm thử hồi quy từ ảnh: 17 ca trượt trên mã Rev0.9.5.3. Bản mới đạt 15 bộ, gồm 68 kiểm tra từ điển. Chưa thử trong Word thật.

## Rev0.9.5.3 (2026-11-13)
- Thêm hai nút thu/phóng giao diện ở Bảng điều khiển và Cài đặt. Tỷ lệ 70–130%, bước 10 điểm phần trăm, mặc định 100% và lưu riêng cho từng cửa sổ.
- Build thành công, 15 bộ ALL PASS. Bộ DPI/zoom đạt 48 kiểm tra, gồm nút bấm WinForms và DPI 200% giả lập trên hai form. Chưa thử trong Word thật.

## Rev0.9.5.2 (2026-11-13)
- Sửa icon «Đọc số» (tnRead): trước là cái loa (đọc thành tiếng), sai ý nghĩa vì nút viết số bằng chữ; nay là số 123, mũi tên xuống xanh, dòng chữ bên dưới.
- Sửa chữ Bảng điều khiển và hộp Cài đặt nhỏ, mờ, vỡ nét. Nguyên nhân tái hiện được: cửa sổ ở chế độ DPI từng màn hình nhưng WinForms (không có app.config trong tiến trình Word) không co giãn khi cửa sổ sang màn hình tỷ lệ khác; ở màn 125% chữ giữ cỡ 96 DPI. Sửa: DpiUtil xử lý WM_DPICHANGED và đồng bộ DPI khi hiện, co giãn cỡ chữ (kể cả điều khiển có font riêng), kích thước, cột danh sách; cột cố định lấy từ chính cột thay cho số cứng; cửa sổ không vượt vùng làm việc của màn hình. Ghi log chẩn đoán và tạo lại Bảng điều khiển nếu không ở chế độ DPI từng màn hình.
- Kiểm thử: thêm TestDpi (26 kiểm tra) vào run-tests; PaneShot và dpi-check.bat chụp thật trên từng màn hình. Kiểm tra trên máy có màn 100% và 125%: tạo ở 100%, tạo ở 125%, kéo từ 100% sang 125%, hộp Cài đặt ở 125%: chữ đúng cỡ, danh sách vừa bề rộng, không tràn màn hình.
- Mười lăm bộ kiểm thử ALL PASS (180, 8654, 309, 238, 195, 131, 131, 294, 233, 70, 11, 35, 27, 446, 26). Chưa thử trong Word thật.

## Rev0.9.5.1 (2026-11-13)
- Làm lại toàn bộ icon Ribbon theo phong cách Word hiện đại (phương án A): chuyển từ kiểu ô vuông bo góc màu xanh đặc có chữ trắng sang phong cách phẳng, line art nét xám đậm `RGB(68,68,68)` dày 1,5–2 px trên nền trong suốt, tối đa một màu nhấn (xanh dương `RGB(0,120,212)` làm chủ đạo, xanh lá `RGB(46,158,91)` cho thêm/xác nhận, đỏ `RGB(209,52,56)` cho xóa/lỗi).
- Mỗi icon một ẩn dụ rõ ràng và đồng nhất:
  - Bảng: lưới nét xám với hàng tiêu đề xanh; áp dụng mọi bảng kèm huy hiệu ✓ xanh lá; áp dụng bảng chọn có khung nét đứt xanh; lặp tiêu đề có mũi tên xoay tròn.
  - Thêm/xóa hàng cột: hàng/cột mới tô xanh nhạt kèm huy hiệu "+" xanh lá đúng vị trí; hàng/cột xóa tô đỏ nhạt kèm huy hiệu "✕" đỏ.
  - Tính toán: khung tính bo góc nền xanh nhạt, ký hiệu toán nét đậm (Σ, x̄, min mũi tên xuống vạch sàn, max mũi tên lên vạch trần, #, π).
  - Định dạng số: số đại diện sắc nét "1,2", "1.2", "1 2" (bỏ chữ US/VN/SI); tăng/giảm số lẻ kèm dấu +/-; căn lề bảng trái/giữa/phải trên trang giấy nhỏ.
  - Màu sắc: icon riêng cho từng nút (màu chữ A có thanh màu, màu nền có bảng nhỏ và giọt màu, màu viền và dòng chẵn); swatch màu phẳng đúng màu tại tâm ô.
  - Ngôn ngữ & Dàn trang: Thuật ngữ (sách gáy xanh kèm tích), Tra thuật ngữ (kính lúp), Gợi ý từ (hai mũi tên hoán đổi), Ngữ pháp (trang giấy kèm sóng xanh), Dàn trang (trang có lề nét đứt), Cài đặt (bánh răng 8 răng có lỗ xanh), Bảng điều khiển (cửa sổ có thanh bên xanh).
  - Icon dự phòng: khung nét đứt xám có dấu "?" ở giữa cho mã lạ.
- Tạo icon riêng cho 6 nút trước đây dùng icon dự phòng: `btnBdClr`, `btnHdrBg`, `btnHdrFg`, `btnSpDec`, `btnSpInc`, `btnZebraClr`.
- Phiên bản & Tiêu đề: `AssemblyInfo.cs` đặt 0.9.5.1; tiêu đề Bảng điều khiển hiển thị đầy đủ `Trợ lý Kỹ thuật - Rev0.9.5.1`.
- Thêm bộ kiểm thử `TestIcons` (446 kiểm tra) xác thực 8 tiêu chí a–h (mọi nút Ribbon có case riêng, diện tích không trong suốt trong khoảng 4%–60%, 4 góc trong suốt, xanh đặc <30%, phân biệt các icon không phải bí danh, mã lạ trả icon dự phòng an toàn, tâm swatch đúng màu, tiêu đề phiên bản).
- Mười bốn bộ kiểm thử ALL PASS (180, 8.654, 309, 238, 195, 131, 131, 294, 233, 70, 11, 35, 27, 446).

## Rev0.9.5 (2026-11-12)
- Sửa nguyên nhân gốc lỗi "chọn màu không đổi" (Rev0.9.3 và 0.9.4): sự kiện mất tiêu điểm của hộp chọn màu ghi đè kết quả OK thành Hủy khi hộp đang đóng. Kiểm thử mô phỏng modal: trượt 3/11 trên 0.9.4, đạt 11/11 trên 0.9.5.
- Tách thẻ My tab thành "My table" (Bảng, Ký tự trong bảng, Hàng và cột, Tính trong bảng với nút nhỏ riêng lẻ; bỏ CSV và Excel; nhãn "My table ●" khi con trỏ trong bảng) và "Chuẩn hóa" (các nhóm còn lại).
- Bỏ ô "Riêng một dòng" ở thẻ Phương trình (tự nhận: đoạn trống hoặc bôi đen cả đoạn thì chèn Equation riêng một dòng, còn lại cùng dòng).
- AI giữ nguyên ngôn ngữ (không còn dịch EN sang VI), có kiểm tra và thử lại; thẻ Viết lại thêm lựa chọn vị trí chèn; Sửa câu dài (AI) hiện ✔ áp dụng được; tiêu đề Bảng điều khiển sửa đúng số bản.
- Từ điển mở rộng (TextDataExt.cs): 150+ nhóm đồng nghĩa Việt, 150+ nhóm Anh, 100+ cặp chính tả phương ngữ, 70+ cụm thừa; gỡ cặp lỗi "xông pha".
- Kiểm thử: 13 bộ ALL PASS (180, 8.654, 309, 238, 195, 131, 131, 294, 233, 70, 11, 35, 27); thử phá mã có chủ ý thì kiểm thử báo lỗi. Chưa thử trong Word thật.

## Rev0.9.4 (2026-11-12)
- Sửa lỗi Word không hiện thẻ My tab sau khi cài Rev0.9.3 (add-in vẫn được nạp). Nguyên nhân nghi ngờ: thư viện font galFont dùng thuộc tính showInRibbon và 
ows không có trong đặc tả MS-CUSTOMUI; Word bỏ qua cả Ribbon khi XML sai lược đồ. Sửa: gỡ hai thuộc tính đó, thêm size="normal" để là thư viện có kích thước theo đặc tả.
- Chỉ đổi phần thư viện font; các thay đổi khác giữ như Rev0.9.3.
- Kiểm thử: build thành công; mười bộ ALL PASS (180, 8654, 309, 238, 156, 131, 131, 294, 233, 45). XML Ribbon đọc được bằng trình phân tích XML.
- Chưa thử trong Word thật (cần cài lại và kiểm tra thẻ My tab xuất hiện).
## Rev0.9.3 (2026-11-12)
- Đổi tên nhóm Ribbon `grpRv` và thẻ Bảng điều khiển `pRv` từ *Rà soát* sang **Chuẩn hóa kỹ thuật** (mục D).
- Cải tiến nhóm Bảng & Cỡ chữ trên Ribbon (mục A, B, C): thêm 3 nút căn bảng trái/giữa/phải (`AlignTables`), thay spin box viền bằng ComboBox `cbBd` (`0.25–6 pt`), thay ComboBox font bằng Gallery `galFont` (4 font kèm xem trước), thay spin box cỡ chữ bằng ComboBox `cbSize` (`8–72 pt`), thay spin box `Thêm N` bằng ComboBox `cbN` (`0–3`, nhập tay đến `20`).
- Bảng chọn màu hợp nhất (`ColorPickerForm.cs` với *Tự động*, *Màu giao diện 10×6*, *Màu chuẩn 10×1*, *Thêm màu...*) và nhóm **Dàn trang** (`WordLayout.cs` ở bên trái cùng của tab với 4 nút: *Chuẩn hóa dàn trang*, *Chuẩn hóa đoạn văn*, *Thêm 1 trang ngang*, *Thêm 1 trang dọc*) (mục F).
- Mở rộng Văn phong, Ngữ pháp & Đồng nghĩa ngoại tuyến (mục E): thêm `ViSpellData.cs` (luật âm tiết tiếng Việt `ViSyllable`, ≥360 cặp chính tả/phương ngữ, ≥60 cụm trùng lặp Hán-Việt, bảo vệ thuật ngữ/viết tắt/đơn vị, lỗi nhầm ngữ cảnh), mở rộng kho đồng nghĩa Việt/Anh (≥100 từ mỗi ngôn ngữ, không bắc cầu chéo nhóm), thêm `tu-dien-nguoi-dung.csv`, thay `cbAlt` bằng `ListBox` bấm đúp/Enter thay ngay và giữ nguyên định dạng ký tự khi thay thế.
- Tích hợp AI tùy chọn & kiểm soát an toàn (mục G): `AiClient.cs` tự nhận diện khóa API (`sk-ant-`, `AIza`, `sk-`), truy vấn danh sách model, gọi API qua HTTP thuần + TLS 1.2, lưu khóa bằng DPAPI (không lộ khóa), *Dùng AI* ở Văn phong (sửa câu dài có ngữ cảnh) và Ngữ pháp (gợi ý đồng nghĩa), thẻ mới **Viết lại** (giới hạn ≤500 từ, chọn giọng văn, bảo toàn định dạng đoạn và Hoàn tác), thanh chân trang thống kê token phiên làm việc.
- Bật Per-Monitor DPI Awareness V2 (`DpiUtil.cs`) cho Bảng điều khiển, Cài đặt và Bảng chọn màu (mục H).
- Kiểm thử: mười bộ ALL PASS (180, 8.654, 309, 238, 156, 131, 131, 294, 233, 45). Chưa thử trong Word thật.

## Rev0.9.1D (2026-11-12)
- Sửa viền bảng: màu, kiểu nét, độ dày viền áp dụng độc lập (trước đây đổi một thuộc tính thì cả ba bị áp).
- Tích theo định dạng nhiều hàng: bôi đen một cột nhiều ô, mỗi ô nhận tích của hàng mình; hàng không khớp giữ nguyên; một lần hoàn tác.
- Kiểm thử: 157, 8.654, 309, 206, 139, 131 ALL PASS; phá mã có chủ ý thì hơn 20 kiểm thử mới báo lỗi. Chưa thử trong Word thật.

## Rev0.9.1C (2026-11-11) — làm tiếp nhánh 0.9.1 (từ Rev0.9.1B)
- Thêm phép nhân (prod) và nút "Tích theo định dạng": nhân các ô đúng kiểu lọc định dạng trong hàng chứa con trỏ, không tính ô con trỏ, ghi vào ô con trỏ, không thêm hàng cột.
- Nhóm Bảng và Ký tự trong bảng: mọi điều khiển có tác dụng ngay lên bảng đang chọn (mỗi điều khiển chỉ ghi nhóm thuộc tính của nó). Mọi bảng và Bảng này vẫn áp dụng toàn bộ thông số.
- Kiểm thử: 157, 8.654, 309, 170 (137 cũ + 33 mới), 139, 131 ALL PASS; thử phá mã có chủ ý thì 17 kiểm thử mới báo lỗi. Bộ 122 (văn bản) không có ở nhánh này. Chưa thử trong Word thật.

## Rev0.9.1A (2026-11-09) — nhánh thử nghiệm từ Rev0.9 (0.9.0)
- Ý tưởng: Tính theo định dạng ô. Nút mới "Tổng theo định dạng" (nhóm Tính) chỉ cộng các ô khớp kiểu lọc chọn trong Cài đặt: chữ đậm, tô sáng (highlight), màu chữ, hoặc màu nền ô.
- Nếu không ô nào khớp, lệnh dừng trước khi sửa bảng, không thêm cột trống.
- Lọc không tính ô chữ đậm một phần (ví dụ "12" đậm trong "12,5") là không đậm; chữ đậm do kiểu bảng (hàng tiêu đề) được tính là đậm.
- Cờ CalcFmtOn được đặt lại mỗi lần bấm nút Tổng, Trung bình, Nhỏ nhất, Lớn nhất, Đếm số để không làm lệch nút thường.
- Chưa làm: điền kết quả vào đúng ô con trỏ (hiện vẫn điền vào cột bên phải hoặc hàng bên dưới như nút Tổng thường); chưa thử trong Word thật.
- Kiểm thử: 150 lõi, 8.654 ma trận số, 309 hóa học, 137 lệnh bảng (gồm 5 kiểm tra mới cho lọc định dạng), 139 rà soát, 131 đọc số — ALL PASS. Chưa thử trong Word thật.

## Rev0.9.7 (2026-11-09)
- Nạp bộ từ điển thuật ngữ chuyên ngành mở rộng (167 thuật ngữ, 12 nhóm) từ dự án "Tu dien chuyen nganh" dùng chung: **104 thuật ngữ có ≥2 cách gọi** (dùng được cho cả kiểm tra Thuật ngữ và Tra thuật ngữ) được nạp vào `tu-dien-thuat-ngu.csv` mẫu; 63 thuật ngữ còn lại chỉ có 1 cách gọi tiếng Việt nên bị bộ đọc CSV của add-in tự loại (một dạng thì không có gì để thống nhất) — chúng vẫn có trong file database `Tu dien chuyen nganh` cho mục đích tra cứu/dùng ở app khác, chỉ không hiện trong My tab trừ khi bạn thêm cách gọi thứ hai.
- Đã dò tự động 38 trường hợp một cụm dài (vd "mối hàn vòng", "ứng suất chảy dòng", "chiều dày danh định") chứa cụm ngắn của nhóm khác (vd "mối hàn", "ứng suất chảy", "chiều dày") làm nền; đối chiếu từng trường hợp xác nhận đều là khái niệm kỹ thuật thực sự khác nhau (đúng theo thiết kế: cụm dài giữ đúng nhóm riêng của nó, không làm sai lệch cách đếm của nhóm chung).
- Kiểm thử: 7 bộ vẫn ALL PASS (150, 8654, 309, 132, 148, 131, 122). Thêm kiểm thử riêng (ngoài 7 bộ) dựng 3 kịch bản dễ xung đột nhất (chiều dày chung/riêng, mối hàn chung/riêng, ứng suất chảy/ứng suất chảy dòng) — cả ba đều báo đúng, không bị cụm mới làm lệch.
- Chưa thử trong Word thật.

## Rev0.9.6 (2026-11-09)
- Bổ sung từ điển thuật ngữ mẫu (`tu-dien-thuat-ngu.csv`) theo chuyên ngành Ăn mòn và Đánh giá toàn vẹn đường ống: thêm 29 khái niệm (ăn mòn điện hóa/khe/xói mòn/chua/ngọt, nứt do hydro (HIC), rộp hydro, nứt do ứng suất sunfua (SSC), hệ số an toàn, hệ số thiết kế, chiều dày danh định, chiều dày còn lại thực tế, chiều dày tối thiểu yêu cầu (t_min), hệ số bền còn lại (RSF), áp suất an toàn (P_safe), NDT, UT, ILI, pig thông minh, khuyết tật, vùng mất kim loại, đánh giá toàn vẹn đường ống, đánh giá rủi ro, chu kỳ kiểm tra lại, HAZ, ứng suất dọc trục, áp suất thử thủy tĩnh, đường kính ngoài/trong); thêm "khí cacbonic" vào nhóm CO2.
- Khi soạn, hai biến thể suýt trùng đúng câu mẫu trong bộ kiểm thử Thuật ngữ ("chiều dày thiết kế", "độ dày còn lại") đã bị bỏ hoặc đổi chữ để không lấn vào nhóm "chiều dày|độ dày|bề dày" có sẵn.
- Từ điển mẫu vẫn chỉ là bộ khởi đầu, chưa được chuyên gia lĩnh vực duyệt lại toàn bộ; đổi tên vẫn theo ASME B31G, API 579, DNV-RP-F101, NACE như AGENTS.md yêu cầu nhưng chưa đối chiếu câu chữ chính thức của từng tiêu chuẩn.
- Kiểm thử: 150 lõi, 8.654 ma trận số, 309 hóa học và LaTeX, 132 lệnh bảng, 148 rà soát và Word giả, 131 đọc số, 122 bộ máy văn bản — bảy bộ ALL PASS. Chưa thử trong Word thật.

## Rev0.9.5 (2026-11-08)
- Thêm kiểm tra Thuật ngữ (thống nhất cách gọi theo từ điển CSV và theo đa số trong văn bản, chính tả hóa/hoá, kỹ/kĩ, i/y, dấu thanh, viết tắt) và Văn phong (câu dài, khẩu ngữ, lặp từ, dấu câu, cụm văn mẫu, từ nối lặp dày, xưng hô). Offline hoàn toàn.
- Nhóm Ribbon "Ngôn ngữ": Thuật ngữ, Văn phong, Tra thuật ngữ, Gợi ý từ. Hai thẻ mới trong Bảng điều khiển dùng chung bộ duyệt với Chuẩn hóa (WordReview chế độ term, style).
- Tệp người dùng sửa được: tu-dien-thuat-ngu.csv, tu-dong-nghia.csv trong %AppData%\TroLyKyThuat.
- Kiểm thử: thêm 122 kiểm tra bộ máy văn bản (TestText) và 17 kiểm tra Word giả; tổng bảy bộ ALL PASS (150, 8.654, 309, 132, 148, 131, 122). Chưa thử trong Word thật.
