<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# Khắc phục sự cố docxi: lỗi cài đặt, lỗi khi chạy và đầu ra bị hỏng

Hãy tìm triệu chứng của bạn bên dưới. Phần lớn sự cố ở ứng dụng cũ đến từ docxi 0.0.x, và nâng cấp lên 0.1.0 sẽ khắc phục chúng mà không cần sửa mã.

## <a id="installation"></a>Cài đặt

**Bundler không tìm được phiên bản `nokogiri` tương thích trong ứng dụng Rails 6+.**
docxi 0.0.6 ghim `nokogiri < 1.7`, trong khi mọi phiên bản Rails từ 6.0 đều cần nokogiri 1.8.5 trở lên. Hãy dùng `gem "docxi", "~> 0.1"`.

**Bundler không tìm được phiên bản `rubyzip` tương thích.**
docxi 0.0.6 ghim `rubyzip ~> 0.9.9`, nên xung đột với bất kỳ gem nào cần rubyzip mới hơn. docxi 0.1.0 chấp nhận rubyzip từ 1.3 đến 3.x.

**Bundler cứ cài bản 0.0.6.**
Hoặc Gemfile đang cố định docxi ở bản 0.0.x, chẳng hạn bằng `"~> 0.0.6"`, hoặc Ruby cũ hơn 2.5, phiên bản mà 0.1.0 yêu cầu. Hãy đổi dòng trong Gemfile thành `gem "docxi", "~> 0.1"` và kiểm tra `ruby -v`.

**Biên dịch nokogiri thất bại.**
Không có gem nokogiri biên dịch sẵn nào khớp với nền tảng của bạn, nên gem đang được dựng từ mã nguồn. Hãy cài một trình biên dịch C và làm theo [hướng dẫn cài đặt nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>Lỗi khi dựng tài liệu

**Lỗi `NoMethodError: undefined method 'present?'` bên ngoài Rails.**
docxi 0.0.6 mặc định rằng ActiveSupport đã được nạp. 0.1.0 không cần điều đó.

**Lỗi `NameError` khi gọi `tab`, `frame` hoặc `page_numbers` của đầu trang.**
Các phương thức này bị hỏng trong 0.0.6 và hoạt động trong 0.1.0, kể cả `document.frame("text")`.

**Lỗi `Errno::ENOENT` khi thêm hình ảnh.**
Đường dẫn không tồn tại so với thư mục mà chương trình của bạn đang chạy. Trên Ruby 3 với docxi 0.0.6, URL hình ảnh cũng thất bại theo cách này, vì 0.0.6 dùng `Kernel#open`. 0.1.0 tải URL bằng `URI#open`.

## <a id="documents-that-dont-look-right"></a>Tài liệu hiển thị không đúng

**Google Docs hoặc LibreOffice không mở được tệp.**
docxi 0.0.6 dùng với rubyzip 3 sẽ ghi bản ghi Zip64 vào mọi mục, và các chương trình này từ chối chúng. 0.1.0 ghi tệp nén mà không có các bản ghi đó.

**Trang đầu không có đầu trang hoặc chân trang.**
docxi dành cho trang đầu một đầu trang và chân trang riêng. Hãy đặt chúng bằng `add_first_header` và `add_first_footer`. Xem [Cấu hình](Configuration#first-page).

**Thiếu một lề, hoặc bố cục trang trông sai.**
Khi đặt lề, hãy truyền `top`, `right`, `bottom` và `left` cùng lúc. Những giá trị bạn bỏ qua sẽ được ghi rỗng.

**Khoảng trắng giữa các run văn bản bị mất.**
Word xóa khoảng trắng ở đầu hoặc cuối một run. Hãy thêm `space: true` vào lời gọi `text` đó.

**Hình ảnh trong đầu trang, trong ô bảng hoặc dùng `pH`/`pV` rất nhỏ hoặc có dạng hình vuông.**
Vấn đề đã biết trong 0.1.0: kích thước của chúng được tính bằng `width * height * 14.92` EMU cho cả hai chiều. Hãy dùng hình ảnh nội dòng được đặt kích thước bằng `style:`, hoặc xem cách khắc phục tạm thời trong [Thực hành tốt](Best-Practices#images).

**Hình ảnh bị mất khi kết xuất cùng một tài liệu lần thứ hai.**
Hình ảnh thêm từ một `File` đang mở hoặc IO khác được đọc trong lúc `render`, và lần kết xuất thứ hai sẽ đọc phải một luồng rỗng. Hãy truyền đường dẫn thay thế, hoặc chỉ kết xuất một lần.

**Viền đặt trên `table` xuất hiện quanh từng ô.**
Vấn đề đã biết: `borders:` ở cấp bảng được áp dụng cho từng ô thay vì cho đường viền ngoài của bảng. Để chỉ vẽ đường viền ngoài, hãy đặt `borders:` trên các ô ở rìa bảng, chỉ cho các cạnh phía ngoài của chúng.

**Số trang trong ô bảng bắt đầu bằng "GlobalOptions".**
Đây là vấn đề đã biết trong `TableCell#page_numbers`. Hãy đặt số trang ở chân trang hoặc đầu trang thay thế.

**Hình ảnh JPEG hoặc GIF hiển thị sai hoặc không xuất hiện.**
Mọi hình ảnh đều được lưu dưới dạng part PNG với kiểu nội dung PNG, bất kể định dạng thật là gì, nên một số chương trình có thể không hiển thị các định dạng khác. Hãy chuyển hình ảnh sang PNG trước khi thêm vào.

## <a id="inspecting-a-generated-file"></a>Kiểm tra một tệp đã tạo

Tệp `.docx` là một tệp nén zip, nên bạn có thể xem bên trong:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Nếu sự cố vẫn còn, hãy [mở một issue](https://github.com/hammad417/docxi/issues) kèm một script ngắn tái hiện sự cố và phiên bản của Ruby, docxi, nokogiri và rubyzip.

## <a id="related-pages"></a>Trang liên quan

- [Câu hỏi thường gặp](FAQ)
- [Thực hành tốt](Best-Practices)
- [Cài đặt](Installation)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
