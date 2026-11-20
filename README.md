# Trợ lý Kỹ thuật cho Word (Technical Assistant for Word)

**Tiếng Việt** | [English](#english)

Add-in cho Microsoft Word dành cho người viết báo cáo kỹ thuật: dàn trang, bảng, số liệu, công thức hóa học và toán, đơn vị, cùng các công cụ rà soát văn bản. Chạy trên máy của bạn, không cần máy chủ riêng.

> Giao diện có **tiếng Việt và tiếng Anh**. Đổi bằng nút **Ngôn ngữ (VI | EN)** trên Ribbon hoặc trong Cài đặt, có hiệu lực ngay, không cần mở lại Word. Mặc định là tiếng Việt. Chữ chưa có bản dịch giữ tiếng Việt; chữ đọc số bằng lời và luật kiểm tra văn bản tiếng Việt vẫn là tiếng Việt vì đó là chức năng cho văn bản tiếng Việt.

## Chức năng chính

Add-in thêm ba thẻ vào Ribbon của Word:

- **Công cụ:** gom các lệnh Word hay dùng (chữ, đoạn văn, trang, chèn, ảnh, bảng) vào một thẻ xếp gọn, kèm các nút chuẩn hóa của add-in và nút đổi ngôn ngữ.
- **My table:** áp một kiểu bảng thống nhất cho mọi bảng (tiêu đề, viền, tô xen kẽ, font, cỡ chữ, căn ô), thêm hoặc xóa hàng và cột, tính Tổng, Trung bình, Nhỏ nhất, Lớn nhất, Đếm số, Tích theo định dạng ngay trong bảng (kể cả bảng có ô gộp).
- **Chuẩn hóa:** chuẩn hóa dàn trang và đoạn văn (có chọn canh lề, áp cho nhiều đoạn cùng lúc), định dạng số (kiểu Việt, Mỹ, SI, số chữ số thập phân, kiểu kỹ thuật), đọc số và đơn vị bằng chữ, rà soát công thức hóa học, toán, ký hiệu và đơn vị, soạn phương trình, và các công cụ AI (viết lại, dịch, văn phong, ngữ pháp, thuật ngữ, từ đồng nghĩa).

## Yêu cầu

- Windows 10 trở lên, Microsoft Word bản máy tính (2016 hoặc Microsoft 365).
- .NET Framework 4.x (có sẵn trong Windows). Không cần cài thêm công cụ lập trình: `build.bat` dùng trình biên dịch `csc.exe` đi kèm Windows.

## Cài đặt

1. Tải hoặc clone repo này.
2. Đóng Word.
3. Bấm phải vào `install.bat`, chọn **Run as administrator**.
4. Mở Word, các thẻ mới nằm trên Ribbon.

Gỡ cài đặt: chạy `uninstall.bat` với quyền quản trị.

## Công cụ AI và quyền riêng tư

- Công cụ AI là tùy chọn. Bạn tự nhập **khóa API của mình** (Anthropic, Google Gemini, OpenAI hoặc dịch vụ tương thích OpenAI) trong **Cài đặt**.
- Khóa API được mã hóa bằng DPAPI của Windows và lưu trên máy bạn, không nằm trong repo.
- Chỉ văn bản bạn chủ động gửi bằng các nút AI (thường là phần đang bôi đen) mới được gửi tới nhà cung cấp bạn chọn. Các chức năng khác chạy hoàn toàn trên máy.

## Kiểm thử

`run-tests.bat` chạy các bộ kiểm thử (dùng Word giả). `real-word-check.bat` chạy kiểm tra trên Word thật qua COM. Lịch sử thay đổi: [CHANGELOG.md](CHANGELOG.md).

Add-in được viết và thử chủ yếu trên Word 16 (Microsoft 365). Các bản Word khác có thể chưa được thử.

## Đóng góp

Báo lỗi và đề xuất qua Issues. Khi báo lỗi, hãy ghi rõ bản Word, các bước làm, và thông báo hiển thị.

## Giấy phép

[MIT](LICENSE).

---

## English

A Microsoft Word add-in for technical report writers: page and paragraph layout, tables, numbers, chemical and math formulas, units, and text-review tools. It runs locally on your PC and needs no server.

> The interface is available in **Vietnamese and English**. Switch with the **Language (VI | EN)** button on the ribbon or in Settings; it takes effect immediately, no need to restart Word. Vietnamese is the default. Text without a translation stays in Vietnamese; the number-to-words reader and the checks for Vietnamese text remain Vietnamese because they work on Vietnamese text.

### Features

The add-in adds three tabs to the Word ribbon:

- **Công cụ (Tools):** frequently used Word commands (font, paragraph, page, insert, pictures, tables) in one neatly arranged tab, plus the add-in's own standardization buttons and the language switch.
- **My table:** apply one consistent table style to all tables (header, borders, zebra rows, font, size, cell alignment), add or delete rows and columns, and calculate Sum, Average, Min, Max, Count and format-based Product directly in a table (including tables with merged cells).
- **Chuẩn hóa (Standardize):** page and paragraph standardization (selectable alignment, applies to several selected paragraphs at once), number formatting (Vietnamese, US, SI styles, decimal places, scientific style), reading numbers and units in words, review of chemical and math formulas, symbols and units, an equation composer, and optional AI tools (rewrite, translate, style, grammar, terminology, synonyms).

### Requirements

- Windows 10 or later and desktop Microsoft Word (2016 or Microsoft 365).
- .NET Framework 4.x (included with Windows). No developer tools needed: `build.bat` uses the `csc.exe` compiler that ships with Windows.

### Install

1. Download or clone this repository.
2. Close Word.
3. Right-click `install.bat` and choose **Run as administrator**.
4. Open Word; the new tabs appear on the ribbon.

To remove it, run `uninstall.bat` as administrator.

### AI tools and privacy

- The AI tools are optional. You enter **your own API key** (Anthropic, Google Gemini, OpenAI, or an OpenAI-compatible service) in **Settings**.
- The key is encrypted with Windows DPAPI and stored on your machine, never in this repository.
- Only the text you send with an AI button (usually the selected text) is sent to the provider you chose. Everything else runs locally.

### Tests

`run-tests.bat` runs the test suites (using a mock Word). `real-word-check.bat` runs checks against a real Word instance through COM. See [CHANGELOG.md](CHANGELOG.md) for the history (in Vietnamese).

The add-in is developed and tested mainly on Word 16 (Microsoft 365). Other Word versions may be untested.

### License

[MIT](LICENSE).
