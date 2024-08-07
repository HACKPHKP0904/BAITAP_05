using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace ClassLibrary_Excel
{
    public class ExportExcel
    {
        public void ExportLichSuTuongTacToExcel(List<LichSuTuongTac> lichSuTuongTacs, string filePath, string period)
        {
            var fileInfo = new FileInfo(filePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                // Kiểm tra và xóa bảng tính nếu tồn tại
                var worksheetName = "LichSuTuongTac";
                var existingWorksheet = package.Workbook.Worksheets[worksheetName];
                if (existingWorksheet != null)
                {
                    package.Workbook.Worksheets.Delete(existingWorksheet);
                }

                var worksheet = package.Workbook.Worksheets.Add(worksheetName);

                worksheet.Cells[1, 1].Value = "MaHoaDon";
                worksheet.Cells[1, 2].Value = "ThoiGian";
                worksheet.Cells[1, 3].Value = "HinhThuc";
                worksheet.Cells[1, 4].Value = "NhanVien";

                int row = 2;
                foreach (var tuongTac in lichSuTuongTacs)
                {
                    if ((period == "week" && CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(tuongTac.ThoiGian, CalendarWeekRule.FirstDay, DayOfWeek.Monday) == CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstDay, DayOfWeek.Monday))
                        || (period == "month" && tuongTac.ThoiGian.Month == DateTime.Now.Month))
                    {
                        worksheet.Cells[row, 1].Value = tuongTac.MaHoaDon;
                        worksheet.Cells[row, 2].Value = tuongTac.ThoiGian.ToString("dd/MM/yyyy HH:mm");
                        worksheet.Cells[row, 3].Value = tuongTac.HinhThuc;
                        worksheet.Cells[row, 4].Value = tuongTac.NhanVien;
                        row++;
                    }
                }

                package.Save();
            }
        }
    }
}
