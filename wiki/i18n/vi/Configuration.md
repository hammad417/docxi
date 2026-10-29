<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# Cấu hình docxi: lề trang, thiết lập trang và đơn vị đo

docxi không có tệp cấu hình, biến môi trường hay thiết lập toàn cục. Mọi thứ được đặt theo từng tài liệu hoặc từng thành phần. Trang này trình bày các thiết lập cấp tài liệu và các đơn vị mà tùy chọn sử dụng. Tùy chọn của từng thành phần được liệt kê trong [Tính năng](Features).

## <a id="page-margins"></a>Lề trang

Lề được truyền vào khi bạn tạo tài liệu:

```ruby
require "docxi"

# Mọi giá trị tính bằng một phần hai mươi point (1440 = 1 inch).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Tùy chọn | Mặc định | Ý nghĩa |
|---|---|---|
| `top` | `720` (0,5 inch) | Lề trên |
| `right` | `1440` (1 inch) | Lề phải |
| `bottom` | `720` (0,5 inch) | Lề dưới |
| `left` | `1440` (1 inch) | Lề trái |
| `header` | `190` | Khoảng cách từ mép trên của trang đến đầu trang |
| `footer` | `200` | Khoảng cách từ mép dưới của trang đến chân trang |

Hash bạn truyền vào sẽ thay thế toàn bộ giá trị mặc định, vì vậy hãy đặt `top`, `right`, `bottom` và `left` cùng lúc. Giá trị nào trong bốn giá trị này bị bỏ trống sẽ được ghi thành giá trị rỗng. `header` và `footer` quay về 190 và 200 khi bị thiếu hoặc để trống.

## <a id="page-size-and-orientation"></a>Kích thước và hướng trang

Trang có khổ US Letter theo hướng dọc: 12240 × 15840 phần hai mươi point, tức 8,5 × 11 inch. Trong 0.1.0, bạn không thể đổi kích thước và hướng trang.

## <a id="first-page"></a>Trang đầu tiên

docxi luôn bật thiết lập "Trang đầu khác" của Word. Đầu trang và chân trang thêm bằng `add_header` và `add_footer` được dùng từ trang 2 trở đi. Trang 1 dùng `add_first_header` và `add_first_footer`, vì vậy hãy đặt cả hai nếu trang đầu cần đầu trang hoặc chân trang. Xem [Tính năng](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>Mọi tài liệu đều có gì

- Phông chữ mặc định là Arial 12 pt.
- Các kiểu được định nghĩa: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` và `NoList`.
- Mục danh sách có dấu đầu dòng (`ul: true`) dùng dấu gạch ngang làm dấu đầu dòng.
- Thuộc tính tài liệu ghi tên ứng dụng là "Docxi Ruby Gem" và lưu thời điểm tài liệu được kết xuất.

## <a id="units"></a>Đơn vị đo

| Tùy chọn | Đơn vị |
|---|---|
| `size:` của văn bản | Point (docxi nhân đôi để thành nửa point theo cách Word lưu) |
| Lề, `left:` và `right:` của đoạn văn, `iwidth:` của bảng, `height:` của hàng | Một phần hai mươi point (1440 mỗi inch) |
| Độ dày `borders:` của ô | Một phần tám point |
| `columns_width:` và `width:` của ô | Được nhân với 14.92 để ra phần hai mươi point, gần với pixel ở 96 dpi |
| `style:` của hình ảnh nội dòng | Độ dài CSS như `"120pt"` |
| `pH:` và `pV:` của hình ảnh định vị | EMU (914400 mỗi inch) |
| Màu chữ và màu tô | Giá trị thập lục phân không có `#`, ví dụ `"8B0000"` |
| `fill_color:` của hình ảnh | Màu VML có `#`, ví dụ `"#00FF00"` |

## <a id="related-pages"></a>Trang liên quan

- [Tính năng](Features)
- [Tham chiếu API](API-Reference)
- [Thực hành tốt](Best-Practices)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
