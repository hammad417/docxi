<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# Tính năng của docxi: đoạn văn, văn bản, bảng, hình ảnh, đầu trang và chân trang

Trang này liệt kê mọi thành phần mà docxi có thể thêm vào tài liệu Word, cùng các tùy chọn mà từng thành phần chấp nhận. Các ví dụ đều dựa trên phần thiết lập sau:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Bạn có thể bỏ qua những tùy chọn không cần dùng. Màu là giá trị thập lục phân không có `#`, ví dụ `"8B0000"`. Về đơn vị đo và thiết lập trang, xem [Cấu hình](Configuration).

## <a id="paragraphs"></a>Đoạn văn

`document.p(options) { |p| ... }` thêm một đoạn văn.

| Tùy chọn | Giá trị | Tác dụng |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Căn chỉnh. `"both"` căn đều hai bên |
| `fill` | Màu thập lục phân | Tô bóng phía sau đoạn văn và văn bản của nó |
| `left`, `right` | Một phần hai mươi point | Thụt lề. Hãy đặt cả hai, vì chỉ đặt một giá trị sẽ bị bỏ qua |
| `ul` | `true` | Biến đoạn văn thành một mục danh sách có dấu đầu dòng |
| `bottom` | `true` | Vẽ viền bên dưới đoạn văn |

Các tùy chọn văn bản như `size:`, `bold:` hoặc `color:` cũng có thể truyền cho `p`. Chúng trở thành giá trị mặc định cho mọi run văn bản trong đoạn.

Bên trong khối, bạn có thể gọi `text`, `br`, `tab`, `image`, `hyperlink` và `frame`. `document.text(string, options)` là cách viết tắt cho một đoạn văn chỉ có một run.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Văn bản

`p.text(string, options)` thêm một run văn bản.

| Tùy chọn | Giá trị | Tác dụng |
|---|---|---|
| `bold` | `true` | In đậm |
| `italic` | `true` | In nghiêng |
| `underline` | `true` | Gạch chân đơn |
| `color` | Màu thập lục phân | Màu chữ |
| `size` | Point | Cỡ chữ |
| `font` | Tên phông chữ | Phông chữ, ví dụ `"Georgia"` |
| `fill` | Màu thập lục phân | Tô bóng phía sau văn bản |
| `space` | `true` | Giữ khoảng trắng ở đầu và cuối văn bản |
| `br` | `true` | Thêm ngắt dòng sau văn bản |

Chuỗi `nil` sẽ được bỏ qua.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Ngắt dòng, ngắt trang và tab

- `p.br`, hoặc `p.br(times: 3)`, thêm ngắt dòng bên trong đoạn văn.
- `document.br(page: true)` bắt đầu một trang mới.
- `p.tab`, hoặc `p.tab(times: 2)`, thêm ký tự tab bên trong đoạn văn.
- `document.tab(page: true)` thêm một đoạn văn chỉ chứa một ký tự tab.

Đừng gọi `document.br` hoặc `document.tab` khi không có `page: true`. Khi đó chúng đặt một run văn bản trực tiếp vào thân tài liệu, và đây không phải là WordprocessingML hợp lệ.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Bảng

`document.table(options)` nhận một khối. Bên trong khối, `table.tr(options)` thêm hàng và `tr.tc(options)` thêm ô.

**Tùy chọn của bảng**

| Tùy chọn | Giá trị | Tác dụng |
|---|---|---|
| `columns_width` | Mảng số | Độ rộng cột, được nhân với 14.92 để ra phần hai mươi point. Đây cũng là độ rộng mặc định của ô |
| `height` | Một phần hai mươi point | Chiều cao tối thiểu của mỗi hàng |
| `iwidth` | Một phần hai mươi point | Khoảng thụt của bảng so với lề trái |
| `style` | Mã kiểu | Kiểu bảng. docxi định nghĩa `TableNormal` và `TableCustom` |
| `width` | Số | Độ rộng ưu tiên của bảng, được ghi với kiểu `auto`. Nếu không có `columns_width`, mỗi ô cũng nhận độ rộng này, nhân với 14.92 |
| `align` | `"left"`, `"center"`, `"right"` | Căn chỉnh bảng. Văn bản thêm bằng `tc.text` cũng được căn chỉnh như vậy |
| `padding` | `true` | Thêm khoảng đệm cố định bên trong các ô |
| `borders` | Hash | Vấn đề đã biết: được áp dụng cho từng ô thay vì cho đường viền ngoài của bảng |

**Tùy chọn của hàng**: `height:` đặt chiều cao tối thiểu của hàng. Mọi tùy chọn ô truyền cho `tr` đều áp dụng cho từng ô trong hàng. Tùy chọn của bảng cũng được truyền xuống từng hàng và từng ô theo cách tương tự.

**Tùy chọn của ô**

| Tùy chọn | Giá trị | Tác dụng |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Độ dày viền tính bằng một phần tám point. `nil` nghĩa là cạnh đó không có viền |
| `fill` | Màu thập lục phân | Tô bóng ô. Ô không có màu tô sẽ có màu trắng (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Căn chỉnh theo chiều dọc |
| `colspan` | Số nguyên | Số cột mà ô trải qua |
| `rowspan` | `true` | Bắt đầu gộp ô theo chiều dọc |
| `merged` | `true` | Tiếp tục phần gộp dọc từ ô phía trên. Nội dung của ô này bị bỏ qua |
| `width` | Số | Độ rộng ô, được nhân với 14.92 |

**Nội dung ô**

- `tc.text(string, options)`: các tùy chọn của ô và của hàng, như `bold:` hoặc `align:`, cũng áp dụng cho văn bản.
- `tc.p(options) { |p| ... }`: một đoạn văn chỉ dùng tùy chọn riêng của nó.
- `tc.image(media, width:, height:)`: một hình ảnh. Xem vấn đề về kích thước ở mục [Hình ảnh](#images).
- `tc.page_numbers`: một trường số trang. Vấn đề đã biết: phía trước có thêm chữ "GlobalOptions".

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>Hình ảnh

`media = document.add_media(source)` đăng ký một hình ảnh. `source` có thể là:

- một đối tượng IO, chẳng hạn `File` hoặc `StringIO`
- một đường dẫn tệp
- một URL `http`, `https` hoặc `ftp`

Hình ảnh được đọc khi bạn gọi `render` và luôn được lưu dưới dạng một part PNG, vì vậy tệp PNG là lựa chọn an toàn.

**Hình ảnh nội dòng**: `p.image(media, options)` đặt hình ảnh trong dòng văn bản.

| Tùy chọn | Giá trị | Tác dụng |
|---|---|---|
| `style` | Hash các thuộc tính CSS | Kích thước, ví dụ `{ width: "120pt", height: "60pt" }` |
| `title` | Chuỗi | Tiêu đề của hình ảnh |
| `fill_color` | Màu VML như `"#00FF00"` | Màu tô phía sau hình ảnh |
| `opacity` | Chuỗi như `"50%"` | Độ mờ đục của màu tô |
| `wrap` | `true` | Tắt việc bao văn bản quanh hình ảnh |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Hình ảnh định vị**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` neo hình ảnh tại một độ lệch so với cột (`pH`) và đoạn văn (`pV`), tính bằng EMU. Phải truyền đủ cả `pH`, `pV` và `wrap`, nếu không hình ảnh sẽ được đặt nội dòng.

**Hình ảnh trong đầu trang**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` tắt việc bao văn bản. Mọi giá trị khác đều khiến văn bản bao quanh hình ảnh theo khung vuông.

**Vấn đề đã biết trong 0.1.0**: hình ảnh định vị, hình ảnh trong đầu trang và hình ảnh trong ô bảng tính cả chiều rộng lẫn chiều cao bằng `width * height * 14.92` EMU. Chúng luôn có dạng hình vuông và thường rất nhỏ. Hình ảnh nội dòng được đặt kích thước bằng `style:` không bị ảnh hưởng. Xem [Thực hành tốt](Best-Practices#images) để biết cách khắc phục tạm thời.

## <a id="headers-and-footers"></a>Đầu trang và chân trang

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- Đầu trang có thể chứa `text`, `br`, `table`, `image` và `page_numbers`. Chân trang có thể chứa `text`, `br`, `table` và `page_numbers`.
- Các tùy chọn truyền cho `Header.new` hoặc `Footer.new` áp dụng cho văn bản bên trong.
- `add_header` và `add_footer` áp dụng từ trang 2 trở đi. Trang 1 luôn dùng `add_first_header` và `add_first_footer`. Xem [Cấu hình](Configuration#first-page).

## <a id="page-numbers"></a>Số trang

`page_numbers(name: "Page", align: "center")` ghi phần văn bản trong `name`, theo sau là một trường PAGE. Số trang hiển thị bằng Arial 10 pt, màu xám đậm. Giá trị mặc định của `align` là `"right"`.

Vấn đề đã biết: mã trường được ghi thiếu dấu gạch chéo ngược trong `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Mục lục

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Khi có số trang, mục đó nhận một điểm dừng tab căn phải tại `pos` (tính bằng một phần hai mươi point; 9360 bằng 6,5 inch), được điền bằng ký tự trong `leader`. Các giá trị ký tự dẫn lấy từ WordprocessingML, chẳng hạn `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` và `"none"`.
- Khi không có số trang, mục đó là phần văn bản đứng sau một ký tự tab.
- Số trang là những số bạn truyền vào. docxi không tạo trường TOC của Word và không tự tính trang.

## <a id="hyperlinks"></a>Siêu liên kết

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` trả về một mã quan hệ, còn `hyperlink(text, url, id)` đặt liên kết. Đích của liên kết lấy từ quan hệ, nên hãy truyền đúng mã được trả về cho URL bạn muốn. Chỉ hỗ trợ liên kết ra bên ngoài.

## <a id="frames-text-boxes"></a>Khung (hộp văn bản)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Khung là các hộp văn bản VML có màu tô xám nhạt bán trong suốt và đường viền đen cố định. Văn bản nằm phía trên và phía dưới khung.
- Bên trong khung, bạn có thể gọi `text`, `br`, `tab`, `image` và `frame`.
- `spacing: true` thêm khoảng cách trước văn bản.

## <a id="output"></a>Đầu ra

`word.render` trả về một `StringIO` chứa toàn bộ gói `.docx`. Tệp nén được ghi mà không có bản ghi Zip64, nên Google Docs và các phiên bản LibreOffice cũ hơn đều mở được. Trang [Kiến trúc](Architecture) liệt kê từng part trong gói.

## <a id="related-pages"></a>Trang liên quan

- [Ví dụ sử dụng](Usage-Examples)
- [Tham chiếu API](API-Reference)
- [Lộ trình](Roadmap): các vấn đề đã biết dự kiến khắc phục trong 0.2.0

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
