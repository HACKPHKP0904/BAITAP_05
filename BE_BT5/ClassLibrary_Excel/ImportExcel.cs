using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ClassLibrary_Excel
{
    public class ImportExcel
    {
        public List<Model_HoaDon.HoaDon> ImportHoaDonFromExcel(string filePath)
        {
            var hoaDons = new List<Model_HoaDon.HoaDon>();

            var fileInfo = new FileInfo(filePath);

            // Thiết lập LicenseContext
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                // Kiểm tra xem có bất kỳ bảng tính nào trong file không
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("File Excel không chứa bảng tính nào.");
                }

                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var hoaDon = new Model_HoaDon.HoaDon
                    {
                        MaHoaDon = worksheet.Cells[row, 1].Text,
                        MaKhachHang = worksheet.Cells[row, 2].Text,
                        NgayXuatHoaDon = DateTime.Parse(worksheet.Cells[row, 3].Text),
                        TongTien = decimal.Parse(worksheet.Cells[row, 4].Text),
                        TongTienNo = decimal.Parse(worksheet.Cells[row, 5].Text)
                    };

                    hoaDons.Add(hoaDon);
                }
            }

            return hoaDons;
        }
    }
}
