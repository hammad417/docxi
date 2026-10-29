<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# Thực hành tốt với docxi: tránh những bất ngờ thường gặp

Những thói quen này xuất phát trực tiếp từ cách docxi hoạt động, và phần lớn giúp ngăn một sự cố được liệt kê trong [Khắc phục sự cố](Troubleshooting). Chúng được nhóm theo chủ đề. Lời khuyên về thiết lập trang áp dụng cho mọi người, còn lời khuyên về bảo mật và hiệu năng quan trọng nhất với các ứng dụng chạy production.

## <a id="page-setup"></a>Thiết lập trang

- **Truyền cả bốn lề cùng lúc.** Hash bạn truyền cho `Docxi::Document.new` thay thế các giá trị mặc định, và giá trị nào trong `top`, `right`, `bottom` và `left` bị bỏ qua sẽ được ghi rỗng.
- **Đặt đầu trang và chân trang cho trang đầu.** docxi luôn bật chế độ trang đầu riêng. Nếu trang 1 cần đầu trang hoặc chân trang, hãy gọi thêm `add_first_header` và `add_first_footer`.

## <a id="text"></a>Văn bản

- **Dùng `space: true` cho khoảng trắng ở đầu hoặc cuối.** Nếu không, Word sẽ xóa chúng và các từ của hai run sẽ dính vào nhau.
- **Đặt định dạng chung ở cấp đoạn văn.** Các tùy chọn truyền cho `p`, như `size:` hoặc `font:`, áp dụng cho mọi run bên trong, nên bạn ít phải lặp lại hơn.
- **Dùng `p.br` để ngắt dòng và `document.br(page: true)` để ngắt trang.** `document.br` không kèm `page: true` sẽ tạo ra XML không hợp lệ.

## <a id="tables"></a>Bảng

- **Đặt viền trên hàng hoặc ô, không đặt trên bảng.** Trong 0.1.0, `borders:` ở cấp bảng được áp dụng cho từng ô thay vì cho đường viền ngoài của bảng.
- **Khai báo `columns_width` cho mọi cột.** Các ô lấy độ rộng từ đó, nên bảng giữ được bố cục như bạn mong muốn.

## <a id="images"></a>Hình ảnh

- **Dùng tệp PNG.** Mọi hình ảnh đều được lưu dưới dạng part PNG.
- **Đặt kích thước hình ảnh nội dòng bằng `style:`**, ví dụ `style: { width: "120pt", height: "60pt" }`. Đây là cách đặt kích thước hình ảnh duy nhất hoạt động đúng như mong đợi trong 0.1.0.
- **Truyền đường dẫn thay vì một `File` đang mở nếu tài liệu có thể được kết xuất nhiều lần.** Hình ảnh từ IO được đọc trong lúc `render`, và lần kết xuất thứ hai sẽ nhận một hình ảnh rỗng.
- **Hình ảnh trong đầu trang, trong ô bảng và hình ảnh định vị cho đến bản 0.2.0.** Chiều rộng và chiều cao của chúng đều được tính bằng `width * height * 14.92` EMU, nên kết quả luôn là hình vuông. Để có một hình vuông cạnh `S` inch, hãy truyền `width` và `height` bằng `Math.sqrt(S * 914400 / 14.92).round`, tức khoảng 248 cho một inch.

## <a id="security"></a>Bảo mật

- **Không truyền chuỗi không đáng tin cậy vào `add_media`.** Đường dẫn khiến docxi đọc tệp đó từ máy chủ của bạn, còn URL khiến máy chủ của bạn tải nó xuống. Hãy truyền một IO do chính bạn mở, hoặc kiểm tra giá trị trước. Xem [Bảo mật](Security).
- **Văn bản được thoát ký tự tự động.** Các chuỗi truyền cho `text` được ghi dưới dạng văn bản XML, nên dữ liệu người dùng nhập không thể thay đổi cấu trúc tài liệu.

## <a id="performance-and-reliability"></a>Hiệu năng và độ tin cậy

- **Dựng tài liệu lớn trong tác vụ nền.** Toàn bộ tài liệu được dựng trong bộ nhớ khi bạn gọi `render`. Xem [Tích hợp](Integrations).
- **Kết xuất một lần, rồi dùng lại các byte.** Hãy giữ `word.render.string` nếu bạn cần gửi hoặc lưu tài liệu nhiều lần.
- **Ghim một phiên bản tương thích** bằng `gem "docxi", "~> 0.1"`.
- **Kiểm thử những gì bạn tạo ra.** Tệp `.docx` là một zip chứa các tệp XML, nên bài kiểm thử có thể giải nén và kiểm tra XML. Chính các bài kiểm thử của docxi cũng hoạt động theo cách này. Xem [Kiến trúc](Architecture).

## <a id="related-pages"></a>Trang liên quan

- [Tính năng](Features)
- [Cấu hình](Configuration)
- [Khắc phục sự cố](Troubleshooting)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
