using System;
using System.Collections.Generic;
using ClassLibrary_Excel;
using Model_HoaDon;

namespace BE_BT5
{
    class Program
    {
        static void Main(string[] args)
        {
            // Tạo danh sách hóa đơn từ file Excel
            var importExcel = new ImportExcel();
            var hoaDons = importExcel.ImportHoaDonFromExcel(@"C:\Users\phi16\OneDrive\Desktop\Book1.xlsx");

            // Lưu lịch sử tương tác
            var quanLyTuongTac = new QuanLyTuongTac();
            quanLyTuongTac.ThemTuongTac("HD001", "Gọi điện", "NhanVienA");
            quanLyTuongTac.ThemTuongTac("HD001", "Gửi mail", "NhanVienB");

            // Lấy lịch sử tương tác của hóa đơn
            var lichSuHD001 = quanLyTuongTac.LayLichSuTuongTac("HD001");

            // Xuất lịch sử tương tác ra file Excel theo tuần hoặc tháng
            var exportExcel = new ExportExcel();
            exportExcel.ExportLichSuTuongTacToExcel(lichSuHD001, "lich_su_tuong_tac.xlsx", "week");

            Console.WriteLine("Chương trình đã chạy thành công. Kiểm tra file 'lich_su_tuong_tac.xlsx' để xem kết quả.");
            Console.ReadKey(); // Đợi người dùng nhấn phím trước khi đóng console
        }
    }
}
