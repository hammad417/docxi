<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# Cài đặt docxi: yêu cầu, thiết lập Bundler và nâng cấp

docxi là một gem viết hoàn toàn bằng Ruby. Khi cài đặt, gem chỉ kéo theo hai phụ thuộc là nokogiri và rubyzip, ngoài ra không có gì khác. Không cần phần mềm văn phòng nào.

## <a id="requirements"></a>Yêu cầu

| Yêu cầu | Phiên bản |
|---|---|
| Ruby | 2.5 trở lên |
| nokogiri | `>= 1.8, < 2`, được cài tự động |
| rubyzip | `>= 1.3, < 4`, được cài tự động |
| Rails | Không bắt buộc |

## <a id="install-with-bundler"></a>Cài đặt bằng Bundler

Thêm gem vào Gemfile của bạn:

```ruby
gem "docxi", "~> 0.1"
```

Sau đó cài đặt:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Cài đặt không dùng Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Kiểm tra cài đặt

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Trong dự án dùng Bundler, hãy chạy lệnh dưới dạng `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`. Lệnh sẽ in ra phiên bản đã cài, ví dụ `0.1.0`.

## <a id="upgrading-from-00x"></a>Nâng cấp từ 0.0.x

Nếu Gemfile của bạn cố định docxi ở bản 0.0.x, chẳng hạn bằng `"~> 0.0.6"`, hãy đổi dòng đó thành `gem "docxi", "~> 0.1"`. Sau đó chạy `bundle update docxi`. Bạn không cần sửa mã, và 0.1.0 tạo ra những tài liệu giống hệt 0.0.6: bộ kiểm thử so sánh đầu ra với các tệp do 0.0.6 tạo ra.

Những gì đã thay đổi bên trong:

- nokogiri và rubyzip giờ có thể ở bất kỳ phiên bản nào trong các khoảng nêu trên, nên docxi cài được cùng các phiên bản Rails hiện tại.
- docxi không còn cần ActiveSupport đã được nạp.
- Tệp nén được ghi mà không có bản ghi Zip64.
- Hình ảnh được nạp mà không dùng `Kernel#open`, nên URL hoạt động trên Ruby 3.
- `tab`, `frame` và số trang trong đầu trang đều hoạt động.

Ứng dụng chạy Ruby cũ hơn 2.5 vẫn tiếp tục nhận bản 0.0.6, vì 0.1.0 khai báo Ruby 2.5 là phiên bản tối thiểu. [Nhật ký thay đổi](Changelog) liệt kê mọi thay đổi.

## <a id="versions-tested-in-ci"></a>Các phiên bản được kiểm thử trong CI

Mỗi phiên bản Rails được chạy với toàn bộ framework được nạp cùng docxi.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

CI còn chạy:

- các bài kiểm thử không có Rails trên Ruby 2.5, 2.6 và 4.0
- các bài kiểm thử với rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) và nokogiri 1.11 (Ruby 2.5)
- một bước kiểm tra LibreOffice có chuyển đổi được các tài liệu mẫu hay không

## <a id="platform-notes"></a>Ghi chú về nền tảng

- nokogiri cung cấp gem biên dịch sẵn cho các nền tảng Linux, macOS và Windows phổ biến. Nếu không có bản phù hợp, gem sẽ được biên dịch từ mã nguồn và cần một trình biên dịch C. Xem [hướng dẫn cài đặt nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- CI chạy trên Ubuntu 24.04, và các bài kiểm thử cũng đã được chạy trên macOS. Windows không được kiểm thử.

## <a id="if-installation-fails"></a>Nếu cài đặt thất bại

Xung đột của Bundler quanh nokogiri hoặc rubyzip gần như luôn có nghĩa là bản docxi 0.0.6 cũ đang được chọn. Xem [Khắc phục sự cố](Troubleshooting#installation).

## <a id="related-pages"></a>Trang liên quan

- [Bắt đầu](Getting-Started)
- [Cấu hình](Configuration)
- [Triển khai](Deployment)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
