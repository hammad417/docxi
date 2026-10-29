<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Thiết lập môi trường phát triển: chạy bài kiểm thử và công cụ của docxi

Trang này giúp bạn chạy một bản sao làm việc của docxi, đồng thời giải thích các bài kiểm thử, tài liệu mẫu và thiết lập CI.

## <a id="get-the-code"></a>Lấy mã nguồn

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Mọi phiên bản Ruby từ 2.5 trở lên đều dùng được để phát triển. Các phụ thuộc phát triển là minitest và rake.

## <a id="rake-tasks"></a>Các tác vụ Rake

| Lệnh | Tác dụng |
|---|---|
| `bundle exec rake test` | Chạy bộ kiểm thử (tác vụ mặc định) |
| `bundle exec rake samples` | Ghi các tài liệu mẫu vào `tmp/samples/*.docx` để mở bằng Word, LibreOffice hoặc Google Docs |
| `bundle exec rake golden` | Ghi lại `test/golden` từ mã hiện tại. Chỉ dùng sau một thay đổi đầu ra có chủ đích |
| `bundle exec rake build` | Đóng gói gem vào `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Chạy bài kiểm thử với Rails được nạp

Thư mục `gemfiles/` có một Gemfile cho mỗi phiên bản Rails, cùng một số Gemfile cho các phiên bản rubyzip và nokogiri cũ hơn:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` nạp toàn bộ framework trước khi chạy bài kiểm thử. Biến này cũng so sánh quy tắc `blank?` của docxi với quy tắc của ActiveSupport.

## <a id="what-the-tests-cover"></a>Bài kiểm thử bao quát những gì

| Tệp | Kiểm tra |
|---|---|
| `test/test_golden_output.rb` | Mọi part của hai tài liệu mẫu khớp với những gì 0.0.6 tạo ra |
| `test/test_package.rb` | XML hợp lệ về cú pháp, các quan hệ trỏ tới part có thật, mọi part đều có kiểu nội dung và không có bản ghi Zip64 |
| `test/test_media.rb` | Hình ảnh từ đường dẫn, đối tượng IO và URL, và việc một chuỗi `\|` không bao giờ bị chạy như lệnh |
| `test/test_fixed_features.rb` | `tab`, `frame` và số trang trong đầu trang |
| `test/test_util.rb` | `Docxi::Util.blank?` tuân theo quy tắc của ActiveSupport |
| `test/test_rails_independence.rb` | docxi chạy được mà không cần ActiveSupport |
| `test/test_warnings.rb` | Quá trình kết xuất không phát sinh cảnh báo lỗi thời (deprecation) nào từ `lib/` |

Các tài liệu mẫu nằm trong `test/support/sample_documents.rb`, và `test/support/package.rb` đọc tệp `.docx` trực tiếp từ cấu trúc zip của nó.

## <a id="continuous-integration"></a>Tích hợp liên tục

`.github/workflows/ci.yml` chạy ở mỗi lần push lên `master` và ở mỗi pull request. Phạm vi gồm:

- Ruby 2.5, 2.6 và 4.0 không có Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) và 8.1 (4.0)
- Rails `main`, được phép thất bại
- rubyzip 1.3 và 2.4, cùng nokogiri 1.11
- một tác vụ chuyển các tài liệu mẫu sang PDF bằng LibreOffice

Việc phát hành được trình bày trong [Triển khai](Deployment).

## <a id="related-pages"></a>Trang liên quan

- [Đóng góp](Contributing)
- [Kiến trúc](Architecture)
- [Triển khai](Deployment)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
