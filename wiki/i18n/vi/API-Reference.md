<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# Tham chiếu API của docxi: các lớp và phương thức

docxi là một thư viện Ruby, vì vậy trang này liệt kê API Ruby của nó. Không có API HTTP. Tùy chọn của từng phương thức được mô tả trong [Tính năng](Features), còn các lớp được giải thích trong [Kiến trúc](Architecture).

## <a id="docxidocument"></a>Docxi::Document

| Phương thức | Trả về | Ghi chú |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` là lề trang. Xem [Cấu hình](Configuration) |
| `#document` | `Docxi::Word::Document` | Phần thân để bạn thêm nội dung |
| `#render` | `StringIO` | Toàn bộ gói `.docx`. `.string` cho ra các byte |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Phương thức | Trả về | Thêm vào |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Một đoạn văn |
| `text(string, options = {})` | Paragraph | Một đoạn văn có một run văn bản |
| `br(page: true)` | Break | Một ngắt trang |
| `tab(page: true)` | Tab | Một đoạn văn chứa một ký tự tab |
| `table(options = {}) { \|table\| }` | Table | Một bảng |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Các mục của mục lục |
| `frame(text, options = {})` | Paragraph | Một đoạn văn chứa hộp văn bản |
| `add_media(io_path_or_url, options = {})` | Media | Một hình ảnh để đặt bằng `image` |
| `add_hyperlink(url, options = {})` | String | Một mã quan hệ dùng cho `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Đầu trang từ trang 2 trở đi và cho trang 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Chân trang từ trang 2 trở đi và cho trang 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Đoạn văn (`Docxi::Word::Contents::Paragraph`)

| Phương thức | Thêm vào |
|---|---|
| `text(string, options = {})` | Một run văn bản. Tùy chọn của đoạn văn là giá trị mặc định của nó |
| `br(times: 1)` | Ngắt dòng |
| `tab(times: 1)` | Ký tự tab |
| `image(media, options = {})` | Một hình ảnh nội dòng hoặc định vị |
| `hyperlink(text, url, id)` | Một liên kết dùng mã quan hệ `id` lấy từ `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | Một hộp văn bản |

## <a id="tables"></a>Bảng

| Phương thức | Thêm vào |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Một hàng. Tùy chọn của bảng là giá trị mặc định của nó |
| `TableRow#tc(options = {}) { \|tc\| }` | Một ô. Tùy chọn của hàng là giá trị mặc định của nó |
| `TableCell#text(string, options = {})` | Một đoạn văn. Tùy chọn của ô là giá trị mặc định của nó |
| `TableCell#p(options = {}) { \|p\| }` | Một đoạn văn có tùy chọn riêng |
| `TableCell#image(media, width:, height:)` | Một hình ảnh |
| `TableCell#page_numbers` | Một trường số trang |

## <a id="table-of-contents"></a>Mục lục

| Phương thức | Thêm vào |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Một mục trong mục lục |

## <a id="headers-and-footers"></a>Đầu trang và chân trang

| Lớp | Phương thức |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Khung

`Docxi::Word::Contents::Frame` chấp nhận `text`, `br`, `tab`, `image` và `frame` bên trong khối của nó.

## <a id="constants-and-helpers"></a>Hằng số và phương thức hỗ trợ

| Tên | Ý nghĩa |
|---|---|
| `Docxi::VERSION` | Phiên bản đã cài, ví dụ `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Ghi một part vào tệp nén với kích thước được đặt trước. Mọi part đều dùng phương thức này ở bên trong |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Phương thức hỗ trợ nội bộ theo quy tắc `blank?` của ActiveSupport |

## <a id="related-pages"></a>Trang liên quan

- [Tính năng](Features)
- [Ví dụ sử dụng](Usage-Examples)
- [Kiến trúc](Architecture)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
