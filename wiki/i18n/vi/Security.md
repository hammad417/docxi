<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# Bảo mật trong docxi: báo cáo sự cố và xử lý dữ liệu không đáng tin cậy

Trang này giải thích cách báo cáo một vấn đề bảo mật, và những điểm bảo mật cần biết khi ứng dụng của bạn truyền dữ liệu người dùng nhập vào docxi.

## <a id="reporting-a-vulnerability"></a>Báo cáo lỗ hổng

Vui lòng không báo cáo vấn đề bảo mật trong các issue công khai. Thay vào đó, hãy gửi email cho người bảo trì. Địa chỉ email nằm trong [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Hãy ghi rõ phiên bản docxi và một cách ngắn gọn để tái hiện sự cố.

## <a id="upgrade-from-00x"></a>Nâng cấp từ 0.0.x

docxi 0.0.x truyền chuỗi đường dẫn hình ảnh cho `Kernel#open`. `Kernel#open` chạy một chuỗi bắt đầu bằng `|` như một lệnh shell, nên kẻ tấn công kiểm soát được đường dẫn hình ảnh có thể chạy lệnh trên máy chủ của bạn. 0.1.0 mở tệp bằng `File.open` và tải URL bằng `URI#open`. Cả hai đều không chạy lệnh.

## <a id="untrusted-input"></a>Dữ liệu không đáng tin cậy

**Truyền văn bản là an toàn.** Văn bản, nội dung bảng và các chuỗi khác được ghi dưới dạng văn bản XML hoặc giá trị thuộc tính, và nokogiri sẽ thoát ký tự cho chúng. Dữ liệu người dùng nhập không thể thêm phần tử vào tài liệu.

**Nguồn hình ảnh cần được cẩn trọng.** `add_media` chấp nhận:

| Nguồn | docxi làm gì |
|---|---|
| Một IO do bạn mở | Đọc IO đó |
| Một đường dẫn tệp | Đọc tệp đó từ ổ đĩa của máy chủ |
| Một URL `http`, `https` hoặc `ftp` | Tải URL đó xuống từ máy chủ của bạn |

Nếu người dùng chọn được đường dẫn, họ có thể khiến bất kỳ tệp nào đọc được trên máy chủ của bạn, chẳng hạn tệp cấu hình hoặc khóa, bị nhúng vào tài liệu. Nếu người dùng chọn được URL, máy chủ của bạn sẽ gửi yêu cầu tới đó, và yêu cầu này có thể chạm tới các địa chỉ nội bộ (giả mạo yêu cầu phía máy chủ, SSRF). Hãy truyền một IO do chính bạn mở, hoặc chỉ cho phép những đường dẫn và máy chủ đã biết.

**Siêu liên kết** trỏ tới bất cứ nơi nào URL chỉ định. Hãy kiểm tra URL do người dùng cung cấp trước khi truyền vào `add_hyperlink`.

## <a id="how-releases-are-protected"></a>Cách các phiên bản được bảo vệ

- Các phiên bản được phát hành từ GitHub Actions bằng cơ chế xuất bản tin cậy (trusted publishing) của RubyGems, nên kho mã không lưu khóa API nào.
- Siêu dữ liệu của gem đặt `rubygems_mfa_required`, nên các lần đẩy gem bằng khóa API cá nhân đều cần xác thực đa yếu tố.

## <a id="related-pages"></a>Trang liên quan

- [Thực hành tốt](Best-Practices)
- [Triển khai](Deployment)
- [Đóng góp](Contributing)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
