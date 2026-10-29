<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# Chạy docxi trong môi trường production và phát hành phiên bản

docxi là một thư viện chứ không phải một dịch vụ, nên bản thân nó không có gì để triển khai. Trang này trình bày những điều quan trọng khi bạn đưa vào hoạt động một ứng dụng dùng docxi, và cách các phiên bản docxi mới được phát hành.

## <a id="shipping-an-app-that-uses-docxi"></a>Đưa vào hoạt động một ứng dụng dùng docxi

- **Phụ thuộc.** `bundle install` cài nokogiri và rubyzip. nokogiri cung cấp gem biên dịch sẵn cho các nền tảng Linux, macOS và Windows phổ biến. Trên các nền tảng khác, ví dụ một số image container tối giản, gem sẽ được biên dịch từ mã nguồn và cần một trình biên dịch C. Xem [hướng dẫn cài đặt nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **Không cần phần mềm văn phòng.** Máy chủ và container không cần Microsoft Word hay LibreOffice.
- **Bộ nhớ.** `render` dựng toàn bộ tài liệu trong bộ nhớ và trả về dưới dạng một `StringIO`. Mức dùng bộ nhớ tăng theo số trang và kích thước hình ảnh.
- **Mạng và tệp.** docxi chỉ đọc hệ thống tệp cho những đường dẫn hình ảnh bạn truyền vào `add_media`, và chỉ dùng mạng cho những URL hình ảnh bạn truyền vào. Xem [Bảo mật](Security).
- **Tài liệu dài.** Hãy dựng chúng trong một tác vụ nền thay vì trong lúc xử lý yêu cầu web. Xem [Tích hợp](Integrations).

## <a id="how-docxi-releases-are-published"></a>Cách các phiên bản docxi được phát hành

Các bước này dành cho người bảo trì. Các phiên bản được đưa lên RubyGems qua GitHub Actions bằng cơ chế xuất bản tin cậy (trusted publishing), nên kho mã không lưu khóa API nào.

1. Cập nhật `lib/docxi/version.rb` và `CHANGELOG.md`, rồi hợp nhất vào `master`.
2. Gắn thẻ cho commit bằng số phiên bản có tiền tố `v`, ví dụ `v0.1.1`, rồi đẩy thẻ đó lên.
3. `.github/workflows/release.yml` đóng gói gem và đẩy lên RubyGems.

Bản phát hành trước làm theo đúng các bước này với số phiên bản như `0.2.0.rc1`. Bundler chỉ cài bản phát hành trước khi được yêu cầu rõ ràng. CI (`.github/workflows/ci.yml`) chạy ở mỗi lần push lên `master` và ở mỗi pull request.

## <a id="related-pages"></a>Trang liên quan

- [Cài đặt](Installation)
- [Thiết lập môi trường phát triển](Development-Setup)
- [Nhật ký thay đổi](Changelog)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
