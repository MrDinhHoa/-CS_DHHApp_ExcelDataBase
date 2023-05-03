using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace _01_TrackingTuanThuTLTT
{
    public partial class Ribbon_Excel
    {

        public void Ribbon_Excel_Load(object sender, RibbonUIEventArgs e)
        {
        
        }

        private void LoadData_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Daily_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook curWorkbook = Globals.ThisAddIn.GetActiveWorkBook();
            Worksheet dailyWorkSheet = curWorkbook.Worksheets["Daily"];
            Worksheet dataVTWorkSheet = curWorkbook.Worksheets["DATA VT"];
            DateTime dateCheck = (DateTime)(dailyWorkSheet.Cells[3, 3] as Excel.Range).Value;
            Range lastDaily = dailyWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range lastData = dataVTWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int rowData = lastData.Row;
            int rowsDaily = lastDaily.Row ;
            

            List<TrackingInfor> trackingsList = new List<TrackingInfor>();
            for (int i = 2; i <= rowData; i++)
            {
                TrackingInfor trackingInfor = new TrackingInfor();
                DateTime dateTracking = (DateTime)(dataVTWorkSheet.Cells[i, 6] as Excel.Range).Value;
                string MaNVBH = (string)(dataVTWorkSheet.Cells[i, 10] as Excel.Range).Value;
                string strdateCheckin = (string)(dataVTWorkSheet.Cells[i, 24] as Excel.Range).Value.ToString();
                string strdateCheckout = (string)(dataVTWorkSheet.Cells[i, 25] as Excel.Range).Value.ToString();
                string MaKH = (string)(dataVTWorkSheet.Cells[i,12] as Excel.Range).Value.ToString();
                string TtrangVT = (string)(dataVTWorkSheet.Cells[i,28] as Excel.Range).Value.ToString();
                string dungtraiTuyen = (string)(dataVTWorkSheet.Cells[i, 23] as Excel.Range).Value.ToString(); 
                string tinhtrangDH = (string)(dataVTWorkSheet.Cells[i,34] as Excel.Range).Value.ToString();
                string lyDo = (string)(dataVTWorkSheet.Cells[i, 33] as Excel.Range).Value.ToString();
                string thoigianVT = (string)(dataVTWorkSheet.Cells[i,26] as Excel.Range).Value.ToString();
                string fakeGPScheckin = (string)(dataVTWorkSheet.Cells[i, 42] as Excel.Range).Value.ToString();
                string fakeGPScheckout;
                if ((dataVTWorkSheet.Cells[i, 43] as Excel.Range).Value == null)
                {
                    fakeGPScheckout = "";
                }
                else
                {
                    fakeGPScheckout = (string)(dataVTWorkSheet.Cells[i, 43] as Excel.Range).Value.ToString();
                }
                trackingInfor.combineDateName = dateTracking.Day.ToString()+dateTracking.Month.ToString()+MaNVBH;
                trackingInfor.checkinDateTime = strdateCheckin;
                trackingInfor.checkoutDateTime = strdateCheckout;
                trackingInfor.MaKhachHang = MaKH;
                trackingInfor.TinhTrangVT = TtrangVT;
                trackingInfor.DungTraiTuyen = dungtraiTuyen;
                trackingInfor.TinhTrangDonHang = tinhtrangDH;
                trackingInfor.LyDo = lyDo;
                trackingInfor.ThoiGianVT = thoigianVT;
                trackingInfor.FakeGPScheckin = fakeGPScheckin;
                trackingInfor.FakeGPScheckout = fakeGPScheckout;
                trackingsList.Add(trackingInfor);
            }
            MessageBox.Show(trackingsList.Count.ToString());
 
            for (int j = 8; j <= rowsDaily; j++)
            {
                List<DateTime> checkinTimes = new List<DateTime>();
                List<DateTime> checkoutTimes = new List<DateTime>();
                List<double> checkinTimeSang = new List<double>();
                List<int> checkinTimeChieu = new List<int>();
                List<int> checkinTimeToi = new List<int>();
                List<double> thoigianVT3p = new List<double>();
                List<double> thoigianVT30p = new List<double>();
                string MaNVBH = (string)(dailyWorkSheet.Cells[j, 9] as Excel.Range).Value;
                string combine = dateCheck.Day.ToString() + dateCheck.Month.ToString() + MaNVBH;
                var listdate_MaNVBH = trackingsList.Where(x => x.combineDateName == combine);
                dailyWorkSheet.Cells[j, 36].Clear();
                #region Lấy tất cả các Nhân viên có hoạt động trong ngày
                if (listdate_MaNVBH.Count() > 0)
                {
                    #region Lấy thời gian check in, check out
                    var strlistdate_MaNVBH_CheckInOut = listdate_MaNVBH.Where(x => x.checkinDateTime != "");
                    if(strlistdate_MaNVBH_CheckInOut.Count() > 0)
                    {
                        foreach (TrackingInfor etrackingInfor in strlistdate_MaNVBH_CheckInOut)
                        {
                            double doubleTimeCheckin = Double.Parse(etrackingInfor.checkinDateTime);
                            DateTime dateCheckin = DateTime.FromOADate(doubleTimeCheckin);
                            checkinTimes.Add(dateCheckin);

                            if(dateCheckin.Hour < 13)
                            { checkinTimeSang.Add(dateCheckin.Hour); }
                            else if(dateCheckin.Hour > 20)
                            { checkinTimeToi.Add(dateCheckin.Hour); }    
                            else { checkinTimeChieu.Add(dateCheckin.Hour); }

                            double doubleTimeCheckout = Double.Parse(etrackingInfor.checkoutDateTime);
                            DateTime dateCheckout = DateTime.FromOADate(doubleTimeCheckout);
                            checkoutTimes.Add(dateCheckout);
                        }
                        dailyWorkSheet.Cells[j, 11].Clear();
                        dailyWorkSheet.Cells[j, 12].Clear();
                        if (checkinTimes.Min().Hour > 9) { dailyWorkSheet.Cells[j, 11] = "x"; }
                        if (checkinTimes.Max().Hour < 16) { dailyWorkSheet.Cells[j, 12] = "x"; }

                        dailyWorkSheet.Cells[j, 26].Clear();
                        dailyWorkSheet.Cells[j, 27].Clear();
                        dailyWorkSheet.Cells[j, 28].Clear();
                        dailyWorkSheet.Cells[j, 26] = checkinTimeSang.Count();
                        dailyWorkSheet.Cells[j, 27] = checkinTimeChieu.Count();
                        dailyWorkSheet.Cells[j, 28] = checkinTimeToi.Count();
                    }
                    #endregion

                    #region Lấy số lần check in Fake GPS
                    var strlistfakeCheckin = listdate_MaNVBH.Where(x => x.FakeGPScheckin != "");
                    var strlistfakeCheckout = listdate_MaNVBH.Where(x => x.FakeGPScheckout != "");
                    dailyWorkSheet.Cells[j, 36].Clear();
                    dailyWorkSheet.Cells[j, 36] = strlistfakeCheckin.Count() + strlistfakeCheckout.Count();
                    #endregion

                    #region Lấy tất cả Mã khách hàng 
                    var strlist_MaKhachHang = listdate_MaNVBH.Where(x => x.MaKhachHang != "");
                    if (strlist_MaKhachHang.Count() > 0)
                    {
                        #region Kiểm tra số call thực hiện
                        var strlist_callVT = strlist_MaKhachHang.Where(x => x.TinhTrangVT == "Đã VT");
                        if (strlist_callVT.Count() > 0)
                        {

                            dailyWorkSheet.Cells[j, 21].Clear();
                            dailyWorkSheet.Cells[j, 21] = strlist_callVT.Count();

                            #region Kiểm tra thời gian viếng thăm dưới 3p và hơn 30p

                            foreach (TrackingInfor etrackingInfor in strlist_callVT)
                            {
                                if (etrackingInfor.ThoiGianVT != "")
                                {
                                    if (double.Parse(etrackingInfor.ThoiGianVT) < 3)
                                    { thoigianVT3p.Add(double.Parse(etrackingInfor.ThoiGianVT)); }
                                    else if (double.Parse(etrackingInfor.ThoiGianVT) > 30)
                                    { thoigianVT30p.Add(double.Parse(etrackingInfor.ThoiGianVT)); }
                                }    
                                
                            }
                            #endregion

                            dailyWorkSheet.Cells[j, 29].Clear();
                            dailyWorkSheet.Cells[j, 29] = thoigianVT3p.Count();
                            dailyWorkSheet.Cells[j, 30].Clear();
                            dailyWorkSheet.Cells[j, 30] = thoigianVT30p.Count();
                            dailyWorkSheet.Cells[j, 31].Clear();
                            dailyWorkSheet.Cells[j, 31].NumberFormat = "0";
                            dailyWorkSheet.Cells[j, 31] = strlist_callVT.Sum(item => double.Parse(item.ThoiGianVT));
                        }
                        #endregion

                        #region Kiểm tra thời gian di chuyển.
                        if (checkinTimes.Count> 0)
                        {
                            TimeSpan timeSpan = checkoutTimes.Max() - checkinTimes.Min();
                            double totalTimesVT = timeSpan.TotalMinutes;
                            dailyWorkSheet.Cells[j, 33].Clear();
                            dailyWorkSheet.Cells[j, 33].NumberFormat = "0";
                            dailyWorkSheet.Cells[j, 33] = totalTimesVT - strlist_callVT.Sum(item => double.Parse(item.ThoiGianVT));
                        }
                        #endregion

                        #region Kiểm tra có đơn hàng hay không
                        var strList_KHCoDonHang = strlist_MaKhachHang.Where(x => x.TinhTrangDonHang != "");
                        if(strList_KHCoDonHang.Count() > 0) 
                        {
                            dailyWorkSheet.Cells[j, 16].Clear();
                            dailyWorkSheet.Cells[j, 16] = strList_KHCoDonHang.Count();
                        }
                        #endregion

                        #region Kiểm tra lý do không có đơn hàng: Khách hàng đóng cửa
                        var strList_KHDongCua = strlist_MaKhachHang.Where(x => x.LyDo == "Khách hàng đóng cửa");
                        if(strList_KHDongCua.Count()>0)
                        {
                            dailyWorkSheet.Cells[j, 18].Clear();
                            dailyWorkSheet.Cells[j, 18]= strList_KHDongCua.Count();
                        }
                        #endregion

                        #region Kiểm tra số call đúng tuyến
                        var strlist_callVTDungTuyen = strlist_MaKhachHang.Where(x => x.DungTraiTuyen == "Đúng tuyến" & x.TinhTrangVT == "Đã VT");
                        if (strlist_callVTDungTuyen.Count() == 0)
                        {
                            dailyWorkSheet.Cells[j, 24].Clear();
                            dailyWorkSheet.Cells[j, 24] = 0;
                        }
                        else
                        {
                            dailyWorkSheet.Cells[j, 24].Clear();
                            dailyWorkSheet.Cells[j, 24] = strlist_callVTDungTuyen.Count();
                        }
                        #endregion

                        #region Kiểm tra số call trái tuyến
                        var strlist_callVTTraiTuyen = strlist_MaKhachHang.Where(x => x.DungTraiTuyen == "Trái tuyến" & x.TinhTrangVT == "Đã VT");
                        if (strlist_callVTTraiTuyen.Count() == 0)
                        {
                            dailyWorkSheet.Cells[j, 23].Clear();
                            dailyWorkSheet.Cells[j, 23] = 0;
                        }
                        else
                        {
                            dailyWorkSheet.Cells[j, 23].Clear();
                            dailyWorkSheet.Cells[j, 23] = strlist_callVTTraiTuyen.Count();
                        }
                        #endregion

                        #region Kiểm tra số khách hàng đã viếng thăm
                        var strlist_TTrangVT = strlist_MaKhachHang.Where(x => x.TinhTrangVT == "Đã VT");
                        if (strlist_TTrangVT.Count() > 0)
                        {
                            dailyWorkSheet.Cells[j, 14].Clear();
                            dailyWorkSheet.Cells[j, 14] = strlist_TTrangVT.Count();
                        }

                        #region Kiểm tra số khách hàng có kế hoạch viếng thăm trong ngày
                        var list_KH_KHViengTham = strlist_MaKhachHang.GroupBy(x => x.MaKhachHang).Select(y => y.First()).ToList();
                        dailyWorkSheet.Cells[j, 13].Clear();
                        dailyWorkSheet.Cells[j, 13] = list_KH_KHViengTham.Count();
                        #endregion

                       
                        #endregion

                    }
                    #endregion
                }
                #endregion
            }
        }

        private void Monthly_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook curWorkbook = Globals.ThisAddIn.GetActiveWorkBook();
            Worksheet monthlyWorkSheet = curWorkbook.Worksheets["MTD"];
            Worksheet dataVTWorkSheet = curWorkbook.Worksheets["DATA VT"];
            Range lastMonthly = monthlyWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range lastData = dataVTWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            int rowData = lastData.Row ;
            int rowsMonthly = lastMonthly.Row;

            List<TrackingInfor> trackingsList = new List<TrackingInfor>();
            for (int i = 2; i <= rowData; i++)
            {
                TrackingInfor trackingInfor = new TrackingInfor();
                DateTime ngayBanHang = (DateTime)(dataVTWorkSheet.Cells[i, 6] as Excel.Range).Value;
                string maNVBH = (string)(dataVTWorkSheet.Cells[i, 10] as Excel.Range).Value;
                string strdateCheckin = (string)(dataVTWorkSheet.Cells[i, 24] as Excel.Range).Value.ToString();
                string strdateCheckout = (string)(dataVTWorkSheet.Cells[i, 25] as Excel.Range).Value.ToString();
                string MaKH = (string)(dataVTWorkSheet.Cells[i, 12] as Excel.Range).Value.ToString();
                string TtrangVT = (string)(dataVTWorkSheet.Cells[i, 28] as Excel.Range).Value.ToString();
                string dungtraiTuyen = (string)(dataVTWorkSheet.Cells[i, 23] as Excel.Range).Value.ToString();
                string tinhtrangDH = (string)(dataVTWorkSheet.Cells[i, 34] as Excel.Range).Value.ToString();
                string lyDo = (string)(dataVTWorkSheet.Cells[i, 33] as Excel.Range).Value.ToString();
                string thoigianVT = (string)(dataVTWorkSheet.Cells[i, 26] as Excel.Range).Value.ToString();
                string fakeGPScheckin = (string)(dataVTWorkSheet.Cells[i, 42] as Excel.Range).Value.ToString();
                string fakeGPScheckout;
                if ((dataVTWorkSheet.Cells[i, 43] as Excel.Range).Value == null)
                {
                    fakeGPScheckout = "";
                }
                else
                {
                    fakeGPScheckout = (string)(dataVTWorkSheet.Cells[i, 43] as Excel.Range).Value.ToString();
                }
                trackingInfor.combineDateName = ngayBanHang.Day.ToString() + ngayBanHang.Month.ToString() + maNVBH;
                trackingInfor.NgayBanHang = ngayBanHang;
                trackingInfor.MaNVBH = maNVBH;
                trackingInfor.checkinDateTime = strdateCheckin;
                trackingInfor.checkoutDateTime = strdateCheckout;
                trackingInfor.MaKhachHang = MaKH;
                trackingInfor.TinhTrangVT = TtrangVT;
                trackingInfor.DungTraiTuyen = dungtraiTuyen;
                trackingInfor.TinhTrangDonHang = tinhtrangDH;
                trackingInfor.LyDo = lyDo;
                trackingInfor.ThoiGianVT = thoigianVT;
                trackingInfor.FakeGPScheckin = fakeGPScheckin;
                trackingInfor.FakeGPScheckout = fakeGPScheckout;
                trackingsList.Add(trackingInfor);
            }
            MessageBox.Show(trackingsList.Count.ToString());

            for (int j = 8; j <= rowsMonthly; j++)
            {
                List<DateTime> checkinTimes = new List<DateTime>();
                List<DateTime> checkinTimes_ByDay = new List<DateTime>();
                List<double> checkinTimeBefore9h = new List<double>();
                List<double> checkoutTimeAfter16h = new List<double>(); 
                List<DateTime> checkoutTimes = new List<DateTime>();
                List<DateTime> checkoutTimes_ByDay = new List<DateTime>();
                List<double> checkinTimeSang = new List<double>();
                List<int> checkinTimeChieu = new List<int>();
                List<int> checkinTimeToi = new List<int>();
                List<double> thoigianVT3p = new List<double>();
                List<double> thoigianVT30p = new List<double>();
                string maNVBH = (string)(monthlyWorkSheet.Cells[j, 9] as Range).Value;
                List<string> listCombineDateName = new List<string>();
                List<double> listThoiGianBH = new List<double>();
                #region Kiểm tra thời gian check in - check out
                var strMaNVBH = trackingsList.Where(x => x.MaNVBH == maNVBH);
                var listNBH_NVBH = strMaNVBH.GroupBy(x => x.NgayBanHang).Where(g => g.Count()>0).SelectMany(y => y).ToList();
                var strlistdate_MaNVBH_CheckInOut = listNBH_NVBH.Where(x => x.checkinDateTime != "");
                var list_SoNgayBH = strMaNVBH.GroupBy(x => x.NgayBanHang).Select( y => y.First()).ToList();
                foreach ( var y in list_SoNgayBH )
                {
                    listCombineDateName.Add(y.combineDateName);
                }    
                foreach (string z in listCombineDateName )
                {
                    var listdateMaNVBH_InOut = strlistdate_MaNVBH_CheckInOut.Where(x => x.combineDateName.Equals(z)).ToList();
                    if (listdateMaNVBH_InOut.Count() > 0 )
                    {
                        DateTime checkinMin = listdateMaNVBH_InOut
                            .Where(i => i.combineDateName == z)
                            .Min(i => DateTime.FromOADate(Double.Parse(i.checkinDateTime)));
                        if (checkinMin.Hour > 9) { checkinTimeBefore9h.Add(checkinMin.Hour); }
                        DateTime checkoutMax = listdateMaNVBH_InOut
                            .Where(i => i.combineDateName == z)
                            .Max(i => DateTime.FromOADate(Double.Parse(i.checkoutDateTime)));
                        if (checkoutMax.Hour < 16) { checkoutTimeAfter16h.Add(checkoutMax.Hour); }
                        listThoiGianBH.Add((checkoutMax - checkinMin).TotalMinutes);
                    }    
                    
                }
                monthlyWorkSheet.Cells[j, 12].Clear();
                monthlyWorkSheet.Cells[j, 13].Clear();
                monthlyWorkSheet.Cells[j, 12] = checkinTimeBefore9h.Count();
                monthlyWorkSheet.Cells[j, 13] = checkoutTimeAfter16h.Count();
                #endregion

                #region Kiểm tra số ngày bán hàng
                if (list_SoNgayBH.Count > 0)
                {
                    list_SoNgayBH.RemoveAll(x => x.NgayBanHang.DayOfWeek == DayOfWeek.Sunday);
                    monthlyWorkSheet.Cells[j, 11].Clear();
                    monthlyWorkSheet.Cells[j, 11] = list_SoNgayBH.Count();
                }
                #endregion

                #region Lấy các thông tin liên quan đến mã khách hàng
                var strlist_MaKhachHang = strMaNVBH.Where(x => x.MaKhachHang != "");
                if (strlist_MaKhachHang.Count() > 0)
                {
                    #region Kiểm tra số call thực hiện
                    var strlist_callVT = strlist_MaKhachHang.Where(x => x.TinhTrangVT == "Đã VT");
                    if (strlist_callVT.Count() > 0)
                    {

                        monthlyWorkSheet.Cells[j, 22].Clear();
                        monthlyWorkSheet.Cells[j, 22] = strlist_callVT.Count();

                        #region Kiểm tra thời gian viếng thăm dưới 3p và hơn 30p

                        foreach (TrackingInfor etrackingInfor in strlist_callVT)
                        {
                            if (etrackingInfor.ThoiGianVT != "")
                            {
                                if (double.Parse(etrackingInfor.ThoiGianVT) < 3)
                                { thoigianVT3p.Add(double.Parse(etrackingInfor.ThoiGianVT)); }
                                else if (double.Parse(etrackingInfor.ThoiGianVT) > 30)
                                { thoigianVT30p.Add(double.Parse(etrackingInfor.ThoiGianVT)); }
                            }

                        }
                        #endregion
                        monthlyWorkSheet.Cells[j, 27].Clear();
                        monthlyWorkSheet.Cells[j, 27] = thoigianVT3p.Count();
                        monthlyWorkSheet.Cells[j, 28].Clear();
                        monthlyWorkSheet.Cells[j, 28] = thoigianVT30p.Count();
                        monthlyWorkSheet.Cells[j, 29].Clear();
                        monthlyWorkSheet.Cells[j, 29].NumberFormat = "0";
                        monthlyWorkSheet.Cells[j, 29] = strlist_callVT.Sum(item => double.Parse(item.ThoiGianVT));
                        monthlyWorkSheet.Cells[j, 31].NumberFormat = "0";
                        monthlyWorkSheet.Cells[j, 31] = listThoiGianBH.Sum(i => i) - strlist_callVT.Sum(item => double.Parse(item.ThoiGianVT));
                    }
                    #endregion

                    #region kiểm tra Fake GPS
                    var strlistfakeCheckin = strlist_MaKhachHang.Where(x => x.FakeGPScheckin != "");
                    var strlistfakeCheckout = strlist_MaKhachHang.Where(x => x.FakeGPScheckout != "");
                    monthlyWorkSheet.Cells[j, 34].Clear();
                    monthlyWorkSheet.Cells[j, 34] = strlistfakeCheckin.Count() + strlistfakeCheckout.Count();
                    #endregion

                    #region Kiểm tra có đơn hàng hay không
                    var strList_KHCoDonHang = strlist_MaKhachHang.Where(x => x.TinhTrangDonHang != "");
                    if (strList_KHCoDonHang.Count() > 0)
                    {
                        monthlyWorkSheet.Cells[j, 17].Clear();
                        monthlyWorkSheet.Cells[j, 17] = strList_KHCoDonHang.Count();
                    }

                    #endregion

                    #region Kiểm tra khách hàng đóng cửa
                    var strList_KHDongCua = strlist_MaKhachHang.Where(x => x.LyDo == "Khách hàng đóng cửa");
                    if (strList_KHDongCua.Count() > 0)
                    {
                        monthlyWorkSheet.Cells[j, 19].Clear();
                        monthlyWorkSheet.Cells[j, 19] = strList_KHDongCua.Count();
                    }
                    #endregion

                    #region Kiểm tra số call đúng tuyến
                    var strlist_callVTDungTuyen = strlist_MaKhachHang.Where(x => x.DungTraiTuyen == "Đúng tuyến" & x.TinhTrangVT == "Đã VT");
                    if (strlist_callVTDungTuyen.Count() == 0)
                    {
                        monthlyWorkSheet.Cells[j, 25].Clear();
                        monthlyWorkSheet.Cells[j, 25] = 0;
                    }
                    else
                    {
                        monthlyWorkSheet.Cells[j, 25].Clear();
                        monthlyWorkSheet.Cells[j, 25] = strlist_callVTDungTuyen.Count();
                    }
                    #endregion

                    #region Kiểm tra số call trái tuyến
                    var strlist_callVTTraiTuyen = strlist_MaKhachHang.Where(x => x.DungTraiTuyen == "Trái tuyến" & x.TinhTrangVT == "Đã VT");
                    if (strlist_callVTTraiTuyen.Count() == 0)
                    {
                        monthlyWorkSheet.Cells[j, 24].Clear();
                        monthlyWorkSheet.Cells[j, 24] = 0;
                    }
                    else
                    {
                        monthlyWorkSheet.Cells[j, 24].Clear();
                        monthlyWorkSheet.Cells[j, 24] = strlist_callVTTraiTuyen.Count();
                    }
                    #endregion

                    #region Kiểm tra số khách hàng có kế hoạch viếng thăm trong ngày
                    var list_KH_KHViengTham = strlist_MaKhachHang.GroupBy(x => x.MaKhachHang).Select(y => y.First()).ToList();
                    monthlyWorkSheet.Cells[j, 14].Clear();
                    monthlyWorkSheet.Cells[j, 14] = list_KH_KHViengTham.Count();
                    #endregion

                    #region Kiểm tra số khách hàng đã viếng thăm
                    var strlist_TTrangVT = strlist_MaKhachHang.Where(x => x.TinhTrangVT == "Đã VT");
                    if (strlist_TTrangVT.Count() > 0)
                    {
                        monthlyWorkSheet.Cells[j, 15].Clear();
                        monthlyWorkSheet.Cells[j, 15] = strlist_TTrangVT.Count();
                    }
                    #endregion
                }
                #endregion

            }
        }


    }
}
