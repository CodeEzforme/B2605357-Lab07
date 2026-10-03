# Hướng dẫn tạo tài nguyên Google

1. Mở <https://script.google.com> và tạo dự án mới.
2. Thay nội dung `Code.gs` bằng tệp `Code.gs` trong thư mục này.
3. Trong Project Settings, bật hiển thị tệp manifest rồi thay bằng `appsscript.json`.
4. Chạy hàm `buildLab07` một lần và cấp quyền cho Google Forms, Google Calendar và trigger.
5. Mở Execution log để lấy hai link điền form. Có thể chạy `getLab07Links` để in lại các link.
6. Dán hai link vào `README.md`, mở bằng cửa sổ ẩn danh để xác nhận quyền điền, rồi chụp màn hình thật đưa vào báo cáo.

Script tạo hai Google Forms, trigger đóng form sinh viên vào ngày 01/12/2026 theo múi giờ Asia/Bangkok, trigger đóng form tham quan khi đủ 80 phản hồi, và một Google Calendar riêng cho kế hoạch tháng 10/2026.
