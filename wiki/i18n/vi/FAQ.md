<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# Câu hỏi thường gặp về docxi: tài liệu Word trong Ruby và Rails

Câu trả lời ngắn cho những câu hỏi được hỏi nhiều nhất. Mỗi câu trả lời đều có liên kết đến phần giải thích chi tiết hơn.

**Tôi có cần cài Microsoft Word hoặc LibreOffice không?**
Không. docxi tự ghi XML và tệp nén zip.

**docxi có cần Rails không?**
Không. docxi chạy được trong mọi chương trình Ruby. Các bài kiểm thử cũng được chạy với Rails từ 5.2 đến 8.1 được nạp. Xem [Tích hợp](Integrations).

**Làm sao để tạo tài liệu Word trong Ruby on Rails?**
Dựng tài liệu trong một controller hoặc một tác vụ nền, rồi gửi `word.render.string` bằng `send_data`. Xem [Tích hợp](Integrations).

**Những phiên bản Ruby và Rails nào được hỗ trợ?**
Ruby 2.5 trở lên. Rails là tùy chọn, và CI kiểm thử Rails từ 5.2 đến 8.1. Xem [Cài đặt](Installation).

**docxi có thể chỉnh sửa một tệp .docx có sẵn hoặc điền vào mẫu Word không?**
Không, docxi chỉ tạo tài liệu mới. Với mẫu có trường trộn thư, hãy dùng một gem dựa trên mẫu như sablon.

**Mục lục có được tạo tự động từ các đầu đề không?**
Không. Bạn tự thêm từng mục và số trang của mục đó. Xem [Tính năng](Features#table-of-contents).

**Tôi có thể dùng danh sách đánh số hoặc kiểu đầu đề không?**
Không có trong 0.1.0. Đoạn văn có thể là mục có dấu đầu dòng (`ul: true`), và bạn có thể định dạng đầu đề bằng các tùy chọn văn bản như `size:` và `bold:`.

**Tôi có thể đổi kích thước hoặc hướng trang không?**
Chưa được. Trang có khổ US Letter theo hướng dọc, và bạn chỉ có thể đổi lề. Xem [Cấu hình](Configuration).

**Những định dạng hình ảnh nào dùng được?**
PNG. Mọi hình ảnh đều được lưu dưới dạng part PNG.

**Tôi có thể thêm hình ảnh từ URL không?**
Có. `add_media` chấp nhận URL `http`, `https` và `ftp`, và tải chúng xuống khi bạn gọi `render`. Đừng truyền URL đến từ người dùng mà không kiểm tra. Xem [Bảo mật](Security).

**Tôi có thể ghi tài liệu thẳng vào tệp không?**
`render` trả về một `StringIO`. Hãy ghi thuộc tính `.string` của nó bằng `File.binwrite("out.docx", word.render.string)`.

**docxi có thể xuất tài liệu rất lớn theo kiểu streaming không?**
Không. Toàn bộ tài liệu được dựng trong bộ nhớ khi bạn gọi `render`. Xem [Triển khai](Deployment).

**Tôi có thể đặt đoạn văn viết từ phải sang trái không?**
0.1.0 không có tùy chọn viết từ phải sang trái. Văn bản được ghi đúng như bạn truyền vào.

**Các vấn đề đã biết được liệt kê ở đâu?**
Trong [Lộ trình](Roadmap) và trong mục "Limitations and known issues" (giới hạn và vấn đề đã biết) của README.

## <a id="related-pages"></a>Trang liên quan

- [Khắc phục sự cố](Troubleshooting)
- [Tính năng](Features)
- [Bắt đầu](Getting-Started)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
