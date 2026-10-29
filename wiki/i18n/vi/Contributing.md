<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# Đóng góp cho docxi

Mọi đóng góp như báo cáo lỗi, bản sửa, tài liệu và bản dịch đều được hoan nghênh. Trang này giải thích cách báo cáo sự cố và những gì một pull request cần có. Để thiết lập bản sao làm việc, xem [Thiết lập môi trường phát triển](Development-Setup).

## <a id="reporting-a-bug"></a>Báo cáo lỗi

Hãy mở một issue tại [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) kèm theo:

- một script Ruby ngắn tái hiện sự cố
- phiên bản của Ruby, docxi, nokogiri và rubyzip (và Rails, nếu bạn dùng)
- điều bạn mong đợi và điều bạn nhận được. Với tài liệu bị hỏng, hãy cho biết bạn đã mở chúng bằng chương trình nào

Vấn đề bảo mật thì hãy báo cáo riêng. Xem [Bảo mật](Security).

## <a id="pull-requests"></a>Pull request

1. Fork kho mã và tạo một nhánh.
2. Thực hiện thay đổi, và giữ đúng phong cách của mã xung quanh.
3. Thêm hoặc cập nhật bài kiểm thử, rồi chạy `bundle exec rake test`.
4. Nếu XML được tạo ra thay đổi có chủ đích, hãy chạy `bundle exec rake golden` và giải thích thay đổi trong pull request. Các tệp golden bảo vệ đầu ra mà người dùng hiện tại đang dựa vào.
5. Cập nhật `CHANGELOG.md` và tài liệu nếu hành vi thay đổi.
6. Mở pull request. CI chạy các bài kiểm thử trên Ruby từ 2.5 đến 4.0 và Rails từ 5.2 đến 8.1, đồng thời kiểm tra LibreOffice có mở được các tài liệu mẫu hay không.

Những thay đổi làm thay đổi tài liệu hiện có, chẳng hạn các bản sửa dự kiến cho 0.2.0, nên nằm trong một pull request riêng để phần khác biệt của tệp golden dễ xem xét.

## <a id="documentation-and-translations"></a>Tài liệu và bản dịch

- README tiếng Anh là nguồn của 15 bản dịch trong `docs/i18n/`. Khi thay đổi README, hãy làm theo `docs/i18n/TRANSLATION_GUIDE.md` trong kho mã.
- Wiki này được viết trong thư mục `wiki/` của kho mã và được phát hành từ đó. `wiki/README.md` giải thích cách làm.

## <a id="related-pages"></a>Trang liên quan

- [Thiết lập môi trường phát triển](Development-Setup)
- [Kiến trúc](Architecture)
- [Lộ trình](Roadmap)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
