<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# Lộ trình và các vấn đề đã biết của docxi

docxi không có lịch phát hành hay lộ trình đã cam kết. Công việc đã lên kế hoạch duy nhất được ghi lại là khắc phục các vấn đề đã biết dưới đây. Việc này dự kiến cho 0.2.0, chưa có ngày cụ thể. Chúng được để ngoài 0.1.0 vì việc khắc phục sẽ làm thay đổi những tài liệu mà người dùng hiện tại đang tạo ra.

## <a id="known-issues-planned-for-020"></a>Các vấn đề đã biết dự kiến khắc phục trong 0.2.0

| Vấn đề | Cách khắc phục tạm thời trong 0.1.0 |
|---|---|
| Hình ảnh định vị, hình ảnh trong đầu trang và hình ảnh trong ô bảng tính kích thước bằng `width * height * 14.92` EMU cho cả hai chiều | Dùng hình ảnh nội dòng được đặt kích thước bằng `style:`, hoặc xem [Thực hành tốt](Best-Practices#images) |
| `borders:` ở cấp bảng được áp dụng cho từng ô thay vì cho đường viền ngoài của bảng | Nếu chỉ cần đường viền ngoài, đặt `borders:` cho các cạnh phía ngoài của các ô ở rìa bảng |
| Mã trường số trang được ghi thiếu dấu gạch chéo ngược trong `\* MERGEFORMAT` | Không có |
| Số trang trong ô bảng có thêm "GlobalOptions" phía trước | Đặt số trang ở đầu trang hoặc chân trang |
| Hình ảnh luôn được lưu dưới dạng part PNG, và `[Content_Types].xml` chỉ khai báo hai đầu trang và hai chân trang đầu tiên | Dùng hình ảnh PNG, và tối đa hai đầu trang cùng hai chân trang |
| Chỉ truyền một số lề sẽ khiến các lề còn lại bị ghi rỗng | Truyền `top`, `right`, `bottom` và `left` cùng lúc |
| `docProps/core.xml` ghi giờ địa phương kèm dấu UTC và không bao giờ ghi tác giả | Không có |
| `Docxi.test` gây lỗi `NoMethodError` | Đừng gọi phương thức này. Đây là phương thức hỗ trợ còn sót lại từ phiên bản cũ |
| `document.br` và `document.tab` không kèm `page: true` tạo ra XML không hợp lệ | Dùng `p.br` và `p.tab` bên trong một đoạn văn |

## <a id="not-planned"></a>Chưa có kế hoạch

Những mục này không nằm trong danh sách ở trên, và chưa có quyết định nào về chúng:

- chỉnh sửa tài liệu có sẵn hoặc điền vào mẫu
- danh sách đánh số, kiểu đầu đề hoặc mục lục tự động
- khổ trang khác hoặc trang hướng ngang

## <a id="related-pages"></a>Trang liên quan

- [Nhật ký thay đổi](Changelog)
- [Đóng góp](Contributing)
- [Câu hỏi thường gặp](FAQ)

---

[← Trang chủ wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
