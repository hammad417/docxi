<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# Nhật ký thay đổi và lịch sử phát hành của docxi

Trang này tóm tắt từng phiên bản. Ghi chú đầy đủ nằm trong [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md) của kho mã, và đó là nguồn tham chiếu nếu hai nơi có khác biệt.

## <a id="releases"></a>Các phiên bản

| Phiên bản | Ngày | Ghi chú |
|---|---|---|
| 0.1.1 | 2026-10-29 | Tóm tắt và mô tả gem chính xác hơn, thêm liên kết tới tài liệu. Không thay đổi mã |
| 0.1.0 | 2026-10-28 | Hoạt động với các phiên bản Ruby, Rails và rubyzip hiện tại. Tạo ra những tài liệu giống hệt 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Bản phát hành trước của 0.1.0 |
| 0.0.6 | 2016-10-18 | Phiên bản cuối cùng trước bản cập nhật 0.1.0 |
| 0.0.2 | 2016-10-14 | Không có ghi chú phát hành nào được lưu lại |
| 0.0.1 | 2016-10-14 | Phiên bản đầu tiên. Không có ghi chú phát hành nào được lưu lại |

Các phiên bản từ 0.0.3 đến 0.0.5 chưa từng được phát hành lên RubyGems.

## <a id="011"></a>0.1.1

Không có thay đổi mã, và tài liệu được tạo ra giống hệt 0.1.0.

- Tóm tắt và mô tả gem giờ đã nói đúng docxi làm gì. Trước đây chúng gọi docxi là trình tạo mẫu.
- RubyGems liên kết tới wiki này làm tài liệu của gem.

## <a id="010"></a>0.1.0

**Khả năng tương thích**

- Cần Ruby 2.5 trở lên. Ứng dụng dùng Ruby cũ hơn sẽ tiếp tục ở bản 0.0.6.
- Hoạt động có hoặc không có Rails, và được kiểm thử cùng Rails từ 5.2 đến 8.1.
- nokogiri `>= 1.8, < 2` thay cho `>= 1.6.3, < 1.7`, khoảng cũ xung đột với mọi phiên bản Rails từ 6.0.
- rubyzip `>= 1.3, < 4` thay cho `~> 0.9.9`.
- Không còn gọi `blank?` và `present?` của ActiveSupport, vốn chỉ hoạt động khi Rails đã nạp chúng.

**Các bản sửa**

- Tài liệu ghi bằng rubyzip 3 không còn chứa bản ghi Zip64, loại bản ghi mà Google Docs và các phiên bản LibreOffice cũ hơn từ chối.
- `add_media` với URL hình ảnh hoạt động trở lại trên Ruby 3.
- `add_media` không còn chạy một chuỗi bắt đầu bằng `|` như lệnh shell.
- `tab` và `frame` không còn gây lỗi `NameError`.
- `document.frame("text")` tạo một hộp văn bản chứa văn bản đó.
- `page_numbers` trong đầu trang không còn gây lỗi `NameError`.

**Gem**

- Người bảo trì mới, cùng liên kết mới tới trang chủ và nhật ký thay đổi.
- Việc phát hành yêu cầu MFA.

## <a id="related-pages"></a>Trang liên quan

- [Lộ trình](Roadmap)
- [Cài đặt](Installation#upgrading-from-00x)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
