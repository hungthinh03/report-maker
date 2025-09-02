== Report Maker Application ==
Made by thinh.hung.du

- Hướng dẫn sử dụng:

1. Setup
- Cài đặt Visual Code
- Mở Visual Code và chạy lệnh sau trong Terminal: pip install openpyxl pillow

2. Chuẩn bị dữ liệu
- File Excel:
	- Hãy để một file excel trong thư mục /default để thêm ảnh vào
- Dữ liệu ảnh:
	- Hãy để tất cả file ảnh HOẶC một file .zip chứa tất cả ảnh vào trong thư mục /images

3. Cài đặt người dùng
- Trước mỗi lần chạy, người dùng có thể thay đổi cài đặt trong ứng dụng tùy theo nhu cầu
- Các thông số có thể thay đổi:
	- START_ROW: Dòng bắt đầu chèn
	- START_COL: Vị trí cột chèn
	- NEW_WIDTH: Độ rộng của ảnh (pixel)
	- MAX_HEIGHT: Chiều cao tối đa của ảnh (pixel)
	- SPACING: Khoảng cách giữa các ảnh (pixel)
	- IMAGES_PER_ROW: Số ảnh chèn mỗi dòng

4. Chạy
- Vào Visual Code và chọn nút Run and Debug phía bên trái cửa sổ
- Chọn Python Debugger -> Python File để bắt đầu chạy 
- Sau khi chạy thành công, file báo cáo mới sẽ nằm trong thư mục hiện hành

5. Lưu ý
- Tất cả ảnh trong file Excel gốc sẽ được xóa trước khi chèn ảnh mới vào nên không cần thiết phải xóa ảnh 
- Các ảnh sẽ được chèn vào file report mới theo trình tự được sắp xếp trong thư mục /images
- Chỉ nên để MỘT file .zip trong /images HOẶC để tất cả file ảnh trong /images, không nên để cả hai loại file trong cùng thư mục này
- Mỗi lần chạy, chương trình sẽ trả về message thành công hoặc lỗi trong Terminal Visual Code, hãy chú ý nếu gặp lỗi khi thực thi 