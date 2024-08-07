using System;
using System.Collections.Generic;

namespace ClassLibrary_Excel
{
    public class LichSuTuongTac
    {
        public string MaHoaDon { get; set; }
        public DateTime ThoiGian { get; set; }
        public string HinhThuc { get; set; } // Gọi điện, gửi mail, gặp trực tiếp
        public string NhanVien { get; set; }
    }

    public class QuanLyTuongTac
    {
        private List<LichSuTuongTac> lichSuTuongTacs = new List<LichSuTuongTac>();

        public void ThemTuongTac(string maHoaDon, string hinhThuc, string nhanVien)
        {
            lichSuTuongTacs.Add(new LichSuTuongTac
            {
                MaHoaDon = maHoaDon,
                ThoiGian = DateTime.Now,
                HinhThuc = hinhThuc,
                NhanVien = nhanVien
            });
        }

        public List<LichSuTuongTac> LayLichSuTuongTac(string maHoaDon)
        {
            return lichSuTuongTacs.FindAll(lt => lt.MaHoaDon == maHoaDon);
        }
    }
}
