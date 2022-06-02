using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Net.Http;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Diagnostics;
using OfficeOpenXml.Table;
using System.Security.Cryptography;

namespace SenLineAndNote
{
    class Program
    {
        public static string Decrypt(string encodedText, string key)
        {
            TripleDESCryptoServiceProvider desCryptoProvider = new TripleDESCryptoServiceProvider();
            MD5CryptoServiceProvider hashMD5Provider = new MD5CryptoServiceProvider();

            byte[] byteHash;
            byte[] byteBuff;

            byteHash = hashMD5Provider.ComputeHash(Encoding.UTF8.GetBytes(key));
            desCryptoProvider.Key = byteHash;
            desCryptoProvider.Mode = CipherMode.ECB; //CBC, CFB
            byteBuff = Convert.FromBase64String(encodedText);

            string plaintext = Encoding.UTF8.GetString(desCryptoProvider.CreateDecryptor().TransformFinalBlock(byteBuff, 0, byteBuff.Length));
            return plaintext;
        }
        static void Main(string[] args)
        {
            //check ECB

            string[] ECB = File.ReadAllLines("Config\\FileConfig.txt");

            if (Convert.ToInt64(Decrypt(ECB[0], "LAMSON")).ToECB())
            {
                Console.WriteLine("Error!!!\r\nSQL Server Configuration Manager connection");

                return;
            }

            DataTable dtVanKienIso = new DataTable();
            DateTime Now = DateTime.Now;

            string query;
            // CHUẨN BỊ ĐỂ XUẤT EXCEL
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            query = "SELECT	* FROM	dbo.VanKienIso";
            dtVanKienIso = DataProvider.Instance.ExecuteQuery(query);

            List<VanKienIso> lsVanKienIsos = new List<VanKienIso>();
            foreach (DataRow dr in dtVanKienIso.Rows)
            {
                VanKienIso vanKienIso = new VanKienIso();

                // add DataTable to List ( ten va dinh dang column phai giong nhau)
                foreach (PropertyInfo objProperty in vanKienIso.GetType().GetProperties())
                {
                    if (dtVanKienIso.Columns.Contains(objProperty.Name) && objProperty.PropertyType == dr[objProperty.Name].GetType())
                    {
                        objProperty.SetValue(vanKienIso, dr[objProperty.Name], null); // add tung gia tri theo ten column
                    }
                }
                lsVanKienIsos.Add(vanKienIso); // add tung dong trong list
            }

            // loc cac van kien qua han
            var VanKienQuaHan = (from temp in lsVanKienIsos
                                  where temp.MocNhacNho <= DateTime.Today? (DateTime.Today - temp.MocNhacNho).Days%7 ==0 : ( temp.MocNhacNho- DateTime.Today).Days ==30
                                  select temp).ToList();
          //  var VanKienQuaHan = lsVanKienIsos.Where(r => r.NgayPhatHanh.AddMonths(r.ChuKy) < DateTime.Today).ToList();

            // lay cac phong PTN
            var lsPTN = lsVanKienIsos.Select(r => r.PTN).Distinct().ToList();

            // moi PTN lay ra danh sach nguoi nhac nho
            foreach (var item in lsPTN)
            {
                // CHUẨN BỊ ĐỂ XUẤT EXCEL
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage pck = new ExcelPackage())
                {

                    pck.Workbook.Properties.Author = "阮林山";
                    pck.Workbook.Properties.Company = "FHS";
                    pck.Workbook.Properties.Title = "Exported by 阮林山";
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Infomation");
                    //Định dạng toàn Sheet
                    ws.Cells.Style.Font.Name = "Times New Roman";
                    ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells.Style.Font.Size = 14;
                    ws.View.ZoomScale = 50;
                    ws.Rows[1].Height = 40;
                    ws.Rows[2].Height = 40;
                    // ws.Rows.Height = 40; // file xuat ra rat nang

                    ws.Column(1).Width = 10;
                    ws.Column(2).Width = 20;
                    ws.Column(3).Width = 100;
                    ws.Column(4).Width = 60;
                    ws.Column(5).Width = 20;
                    ws.Column(6).Width = 20;
                    ws.Column(7).Width = 20;
                    ws.Column(8).Width = 20;
                    ws.Column(9).Width = 20;
                    //ws.Column(1).Style.Numberformat.Format = "MM/dd hh:mm";
                    ws.Column(7).Style.Numberformat.Format = "yyy/MM/dd";
                    ws.Column(9).Style.Numberformat.Format = "yyy/MM/dd";

                    ws.Cells["A1"].Value = "文件提醒名單";
                    ws.Cells["A1:I1"].Merge = true;
                    ws.Cells["A1"].Style.Font.Size = 20;
                    ws.Cells["A1"].Style.Font.Bold = true;

                    ws.Cells["A2"].Value = "項次";
                    ws.Cells["B2"].Value = "文件號";
                    ws.Cells["C2"].Value = "越文名稱";
                    ws.Cells["D2"].Value = "中文名稱";
                    ws.Cells["E2"].Value = "實驗室";
                    ws.Cells["F2"].Value = "制訂/修訂人";
                    ws.Cells["G2"].Value = "發佈日期";
                    ws.Cells["H2"].Value = "週期(月)";
                    ws.Cells["I2"].Value = "到期日";

                    //ws.Cells[2, 1, 2, 6].Style.Font.Bold = true;
                    //ws.Cells[2, 1, 2, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    // ws.Cells[2, 1, 2, 6].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    string pathFile = Path.Combine(System.IO.Directory.GetCurrentDirectory(), $"Report-{item}.xlsx");
                    FileInfo excelFile = new FileInfo(pathFile);


                    string NguoiNhacNho = "";
                    if (item == string.Empty) continue;

                    query = $"SELECT NguoiNhacNho FROM dbo.ComboBox WHERE PhongThiNghiem = N'{item}'";
                    NguoiNhacNho = DataProvider.Instance.ExecuteScalar(query).ToString();

                    var VanKienQuaHan_PTN = VanKienQuaHan.Where(r => r.PTN == item).ToList(); // Loc van kien qua han theo PTN
                    string br = "<br/>"; // xuong hang
                    string space = "&ensp;"; // dau cach

                    string msg = $"<h3>1. Để phù hợp với quy định quản lý tài liệu \"Các phòng thí nghiệm cần đảm bảo tài liệu thường xuyên được xem xét và cập nhật { br } {space} {space}khi cần thiết \" của ISO 17025 và ISO 9001, mỗi đơn vị cần thường xuyên rà soát và cập nhật tài liệu.{ br } {space} {space}為符合ISO 17025與ISO 9001文件管理規定「實驗室應確保定期審查文件與必要時更新」，故各單位應定期執行文件資料審查{ br } {space} {space}與進版。{ br }2. Các đơn vị có các văn kiện sắp hoặc đã hết hạn như phụ kiện xin vui lòng sắp xếp nhân viên phụ trách tiến hành xem xét { br } {space} {space}nội dung và cập nhật văn kiện.{ br } {space} {space}各單位即將過期或是已過期之文件資料如下清單所示，請貴單位應立即安排負責人員進行進行審查文件資料內容與進版作業。{ br }3. Đối với quản lý bất thường với các văn kiện đã hết hạn mà đơn vị không không định kỳ sửa đổi mà xảy ra bất thường khi có { br } {space} {space}kiểm tra nội bộ/bên ngoài thì trách nhiệm sẽ thuộc về đơn vị phụ trách.{ br } {space}{space}針對已過期文件管理異常，貴單位仍無定期改善，造成內 / 外部稽核缺失則由貴單位自行負責。</h3>";  // Create Header

                    int STT = 1;

                    string space100 = "*************************************************************************************************************** ";


                    foreach (var data in VanKienQuaHan_PTN)
                    {
                        ws.Cells[$"A{STT + 2}"].Value = STT;
                        ws.Cells[$"B{STT + 2}"].Value = data.MaVanKien;
                        ws.Cells[$"C{STT + 2}"].Value = data.TenTiengViet;
                        ws.Cells[$"D{STT + 2}"].Value = data.TenTiengTrung;
                        ws.Cells[$"E{STT + 2}"].Value = data.PTN;
                        ws.Cells[$"F{STT + 2}"].Value = data.NguoiLap;
                        ws.Cells[$"G{STT + 2}"].Value = data.NgayPhatHanh;
                        ws.Cells[$"H{STT + 2}"].Value = data.ChuKy;
                        ws.Cells[$"I{STT + 2}"].Value = data.NgayPhatHanh.AddMonths(data.ChuKy);

                        // msg += $"<p>{ STT.ToString("00") }. {data.MaVanKien}{ br }&emsp;&ensp;{data.TenTiengViet}{ br }&emsp;&ensp;{data.TenTiengTrung}</p>\r";
                        STT++;

                        ws.Rows[STT + 1].Height = 40;
                    }


                    ws.Column(4).Style.Font.Name = "DFKai-SB";
                    ws.Column(5).Style.Font.Name = "DFKai-SB";
                    ws.Column(6).Style.Font.Name = "DFKai-SB";
                    ws.Cells[1, 1, 2, 9].Style.Font.Name = "DFKai-SB";

                    ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Column(3).Style.WrapText = true;

                    ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    ws.Cells[2, 1, 2, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[2, 1, 2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //create a range for the table
                    ExcelRange rangeTabel = ws.Cells[2, 1, ws.Dimension.End.Row, ws.Dimension.End.Column];

                    //add a table to the range
                    ExcelTable tab = ws.Tables.Add(rangeTabel, "Table1");

                    //format the table
                    tab.TableStyle = TableStyles.Medium1;

                    pck.SaveAs(excelFile);

                    msg += $"<p>{space100}{space100}</p>";
                    if (NguoiNhacNho != string.Empty)
                    {
                        MainAsync(NguoiNhacNho, msg, item, pathFile).Wait();
                        Console.WriteLine(item);
                    }
                }
            }

            Console.ReadKey();
        }
        static async Task MainAsync(string NguoiNhan, string NoiDung, string PTN, string filePath)
        {
            //Notes
            using (var client = new HttpClient())
            {
                string br = "<br/>"; // xuong hang

                client.BaseAddress = new Uri("http://10.199.1.32:1234");
                // var fileLocations = new List<string>() { @"D:\CSharp\20. SenLineAndNote\SenLineAndNote\SenLineAndNote\bin\Debug\Report-非破壞實驗室.xlsx" };
                var fileLocations = new List<string>() { filePath };
                var mail = new Mail()
                {
                    To = NguoiNhan,
                    //    To = "VNW0014732@VNFPG,VNW0010532@VNFPG,VNW0012950@VNFPG",
                    //CC = "VNW000XXX@VNFPG",
                    Subject = $"Đề nghị quý đơn vị cập nhật hoặc phát hành bản mới đối với các văn kiện sắp hoặc đã quá hạn! {br} 針對即將或已過期之文件，請貴單位執行更新或進版 -- {PTN}",
                    Content = NoiDung,
                    Attachments = new List<AttachmentFile>()
                };
                foreach (var fileLocation in fileLocations)
                {
                    var file = File.Open(fileLocation, FileMode.Open);
                    var file_byteCode = new byte[file.Length];
                    file.Read(file_byteCode, 0, (int)file.Length);
                    var file_string = Convert.ToBase64String(file_byteCode);
                    mail.Attachments.Add(
                        new AttachmentFile()
                        {
                            Name = file.Name.Substring(file.Name.LastIndexOf('\\') + 1),
                            FileOfBase64String = file_string
                        });
                }
                var json_string = Newtonsoft.Json.JsonConvert.SerializeObject(mail);
                var requestContent = new StringContent(json_string, Encoding.UTF8, "application/json");
                var response = await client.PostAsync("/api/Mail", requestContent);
                var responseContent = await response.Content.ReadAsStringAsync();
                Console.WriteLine(PTN);

                //  Console.ReadKey();
            }


        }
        public class Mail
        {
            public string To { get; set; }
            public string CC { get; set; }
            public List<AttachmentFile> Attachments { get; set; }
            public string Subject { get; set; }
            public string From { get; set; }
            public string Content { get; set; }
            public string SystemName { get; set; }
            public string SystemOwner { get; set; }
        }
        public class AttachmentFile
        {
            public string Name { get; set; }
            public string FileOfBase64String { get; set; }
        }

        class VanKienIso
        {
            public string MaVanKien { get; set; }
            public string PTN { get; set; }
            public string TenTiengTrung { get; set; }
            public string TenTiengViet { get; set; }
            public string NguoiLap { get; set; }
            public DateTime MocNhacNho { get; set; }
            public int ChuKy { get; set; }

            public DateTime NgayPhatHanh { get; set; }
        }
    }
}
