<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# Wiki docxi: tạo tài liệu Word (.docx) bằng Ruby

docxi là một gem Ruby tạo tệp Microsoft Word `.docx` từ mã nguồn. Gem hỗ trợ đoạn văn, văn bản có định dạng, bảng, hình ảnh, đầu trang và chân trang, số trang, siêu liên kết và một mục lục đơn giản. docxi không cần Microsoft Word, LibreOffice hay tệp mẫu, và chạy được cả trong Ruby thuần lẫn trong ứng dụng Ruby on Rails.

[README](https://github.com/hammad417/docxi#readme) là phần giới thiệu ngắn. Wiki này trình bày mọi tùy chọn, cung cấp các công thức có thể sao chép ngay, giải thích cách docxi hoạt động bên trong và giúp bạn khi có sự cố.

Phiên bản hiện tại: **0.1.1**, dành cho Ruby 2.5 trở lên. Xem [Nhật ký thay đổi](Changelog).

## <a id="start-here"></a>Bắt đầu từ đây

**Nếu bạn mới dùng docxi**

- [Bắt đầu](Getting-Started): cài gem và tạo tài liệu đầu tiên
- [Cài đặt](Installation): yêu cầu, Bundler và nâng cấp từ 0.0.x
- [Ví dụ sử dụng](Usage-Examples): báo cáo, hóa đơn, thư và bảng dựng từ dữ liệu, sẵn sàng để sao chép

**Xây dựng tài liệu**

- [Tính năng](Features): mọi thành phần và các tùy chọn của chúng
- [Cấu hình](Configuration): lề, thiết lập trang và đơn vị đo
- [Tham chiếu API](API-Reference): các lớp và phương thức trong nháy mắt
- [Thực hành tốt](Best-Practices): những thói quen giúp tránh các bất ngờ thường gặp

**Khi có sự cố**

- [Khắc phục sự cố](Troubleshooting): thông báo lỗi và tài liệu hiển thị không đúng
- [Câu hỏi thường gặp](FAQ): câu trả lời ngắn cho các câu hỏi phổ biến
- [Bảo mật](Security): báo cáo sự cố và xử lý dữ liệu đầu vào không đáng tin cậy

**Tìm hiểu sâu hơn**

- [Kiến trúc](Architecture): cách các lời gọi Ruby trở thành một gói `.docx`
- [Tích hợp](Integrations): controller của Rails, ứng dụng Rack và tác vụ nền
- [Triển khai](Deployment): chạy docxi trong môi trường production và cách phát hành các phiên bản
- [Đóng góp](Contributing) và [Thiết lập môi trường phát triển](Development-Setup)
- [Lộ trình](Roadmap), [Nhật ký thay đổi](Changelog) và [Ghi công](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>docxi làm được gì và không làm được gì

docxi tạo tài liệu mới. docxi không thể mở hay chỉnh sửa một tệp `.docx` có sẵn, không điền vào mẫu Word và không thực hiện trộn thư. docxi không có danh sách đánh số, kiểu đầu đề hay mục lục tự động, và trang luôn có khổ US Letter theo hướng dọc. Các vấn đề đã biết dự kiến khắc phục trong 0.2.0 được liệt kê ở [Lộ trình](Roadmap).

## <a id="credits"></a>Ghi công

docxi do [Hammad Maqbool](https://github.com/hammad417) tại [Phaedra Solutions](https://www.phaedrasolutions.com/) bảo trì. Dự án bắt đầu vào năm 2016 từ [docxer](https://github.com/gtimti/docxer) của Timur Gabdrakipov. Xem [Ghi công](Credits) để biết thêm chi tiết.

## <a id="languages"></a>Ngôn ngữ

Bản tiếng Anh là bản tham chiếu của wiki này. Liên kết đến các bản dịch nằm ở thanh bên và ở đầu mỗi trang.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Báo cáo sự cố](https://github.com/hammad417/docxi/issues)
