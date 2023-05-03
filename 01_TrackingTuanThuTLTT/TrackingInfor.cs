using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BIMSoftLib.MVVM;

namespace _01_TrackingTuanThuTLTT
{
    public class TrackingInfor: PropertyChangedBase
    {
        public string combineDateName { get; set; }
        public DateTime NgayBanHang { get; set; }
        public string MaNVBH { get; set; }
        public string checkinDateTime { get; set; }
        public string checkoutDateTime { get; set; }  
        public string MaKhachHang { get; set; }
        public string TinhTrangVT { get; set; }
        public string DungTraiTuyen { get; set; }
        public string TinhTrangDonHang { get; set; }
        public string LyDo {get; set; }
        public string ThoiGianVT { get; set; }
        public string FakeGPScheckin { get; set; }
        public string FakeGPScheckout { get; set; }
    }
}
