<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Bắt đầu với docxi: tài liệu Word đầu tiên của bạn bằng Ruby

Hướng dẫn này đưa bạn từ một thư mục trống đến một tệp `.docx` đã định dạng. Hướng dẫn giả định bạn nắm được kiến thức Ruby cơ bản. Bạn không cần Rails hay Microsoft Word.

## <a id="1-install-the-gem"></a>1. Cài đặt gem

```sh
gem install docxi
```

Trong dự án có Gemfile, hãy thêm `gem "docxi", "~> 0.1"` rồi chạy `bundle install` thay cho lệnh trên. Xem chi tiết tại [Cài đặt](Installation).

## <a id="2-create-a-document"></a>2. Tạo tài liệu

Lưu đoạn mã sau thành `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Chạy `ruby hello.rb` (hoặc `bundle exec ruby hello.rb` trong dự án dùng Bundler) rồi mở `hello.docx`.

## <a id="3-what-each-line-does"></a>3. Mỗi dòng làm gì

- `Docxi::Document.new` tạo một tài liệu trống với lề trang mặc định.
- `word.document` là phần thân của tài liệu. Mọi thứ bạn thêm vào đều đi qua đây.
- `document.p` thêm một đoạn văn. Các tùy chọn truyền cho `p`, chẳng hạn `align:`, áp dụng cho đoạn văn và cho văn bản bên trong.
- `p.text` thêm một run văn bản. `size:` tính bằng point.
- `word.render` dựng gói `.docx` và trả về một `StringIO`. Thuộc tính `.string` của đối tượng này chứa các byte của tệp.

## <a id="4-add-a-list-and-a-table"></a>4. Thêm danh sách và bảng

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

Có hai điều cần lưu ý:

- `space: true` giữ lại khoảng trắng ở đầu hoặc cuối một run. Nếu thiếu tùy chọn này, Word sẽ xóa khoảng trắng và các từ bị dính vào nhau.
- Tùy chọn đặt trên `tr` áp dụng cho mọi ô trong hàng đó, nên toàn bộ hàng tiêu đề được tô bóng và in đậm.

## <a id="5-where-to-go-next"></a>5. Tiếp theo nên đọc gì

- [Ví dụ sử dụng](Usage-Examples): công thức hoàn chỉnh cho báo cáo, hóa đơn và thư
- [Tính năng](Features): mọi tùy chọn cho đoạn văn, văn bản, bảng, hình ảnh, đầu trang và chân trang
- [Tích hợp](Integrations): gửi tài liệu từ controller của Rails
- [Thực hành tốt](Best-Practices): tránh trước những bất ngờ thường gặp

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
