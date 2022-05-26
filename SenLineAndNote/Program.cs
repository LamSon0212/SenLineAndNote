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

namespace SenLineAndNote
{
    class Program
    {
        static void Main(string[] args)
        {

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
            var VanKienQuaHan1 = (from temp in lsVanKienIsos
                                  where temp.NgayPhatHanh.AddMonths(temp.ChuKy) < DateTime.Today
                                  select temp).ToList();
            var VanKienQuaHan = lsVanKienIsos.Where(r => r.NgayPhatHanh.AddMonths(r.ChuKy) < DateTime.Today).ToList();

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

                    ws.Column(1).Width = 10;
                    ws.Column(2).Width = 20;
                    ws.Column(3).Width = 150;
                    ws.Column(4).Width = 60;
                    ws.Column(5).Width = 20;
                    ws.Column(6).Width = 20;
                    //ws.Column(1).Style.Numberformat.Format = "MM/dd hh:mm";
                    ws.Column(6).Style.Numberformat.Format = "yyy/MM/dd";

                    ws.Cells["A1"].Value = "文件提醒名單";
                    ws.Cells["A1:F1"].Merge = true;
                    ws.Cells["A1"].Style.Font.Size = 20;
                    ws.Cells["A1"].Style.Font.Bold = true;

                    ws.Cells["A2"].Value = "項次";
                    ws.Cells["B2"].Value = "文件號";
                    ws.Cells["C2"].Value = "越問名稱";
                    ws.Cells["D2"].Value = "中文名稱";
                    ws.Cells["E2"].Value = "實驗室";
                    ws.Cells["F2"].Value = "發佈日期";

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

                    string msg = $"<h3>1.為符合ISO 17025與ISO 9001文件管理規定「實驗室應確保定期審查文件與必要時更新」，故各單位應定期執行文件資料審查與進版。{ br }2.各單位即將過期或是已過期之文件資料如下清單所示，請貴單位應立即安排負責人員進行進行審查文件資料內容與進版作業。{ br }3.針對已過期文件管理異常，貴單位仍無定期改善，造成內 / 外部稽核缺失則由貴單位自行負責。</h3>";  // Create Header
                                                                                                                                                                                                                                        //string msg = "<p>";
                    int STT = 1;

                    string space100 = "*************************************************************************************************************** ";


                    foreach (var data in VanKienQuaHan_PTN)
                    {
                        ws.Cells[$"A{STT + 2}"].Value = STT;
                        ws.Cells[$"B{STT + 2}"].Value = data.MaVanKien;
                        ws.Cells[$"C{STT + 2}"].Value = data.TenTiengViet;
                        ws.Cells[$"D{STT + 2}"].Value = data.TenTiengTrung;
                        ws.Cells[$"E{STT + 2}"].Value = data.PTN;
                        ws.Cells[$"F{STT + 2}"].Value = data.NgayPhatHanh;

                        msg += $"<p>{ STT.ToString("00") }. {data.MaVanKien}{ br }&emsp;&ensp;{data.TenTiengViet}{ br }&emsp;&ensp;{data.TenTiengTrung}</p>\r";
                        STT++;
                    }

                    
                    ws.Column(4).Style.Font.Name = "DFKai-SB";
                    ws.Column(5).Style.Font.Name = "DFKai-SB";
                    ws.Cells[1, 1, 2,6].Style.Font.Name = "DFKai-SB";
                    //// vẽ Boder
                    //ws.Cells[1, 1, STT + 1, 6].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    //ws.Cells[1, 1, STT + 1, 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //ws.Cells[1, 1, STT + 1, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    //ws.Cells[1, 1, STT + 1, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    ws.Cells[2, 1, 2, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[2, 1, 2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //create a range for the table
                    ExcelRange rangeTabel = ws.Cells[2, 1, ws.Dimension.End.Row, ws.Dimension.End.Column];
                    //  ExcelRange rangeTabel1 = ws1.Cells[1, 1, ws1.Dimension.End.Row, 1];


                    //add a table to the range
                    ExcelTable tab = ws.Tables.Add(rangeTabel, "Table1");
                    //ExcelTable tab1 = ws1.Tables.Add(rangeTabel1, "Table2");
                    int xxx = 1;
                    //format the table
                    tab.TableStyle = TableStyles.Medium1;
                    //// tab1.TableStyle = TableStyles.Medium1;


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
            //if (NguoiNhacNho != string.Empty)
            //{
            //     MainAsync(NguoiNhacNho, msg).Wait();
            //}

            //var VanKienQuaHan = (from temp in dtVanKienQuaHan.AsEnumerable()
            //              where temp.Field<DateTime>("NgayPhatHanh").AddMonths(temp.Field<int>("ChuKy")) < DateTime.Today
            //              select temp).ToList();

            //query = "SELECT PhongThiNghiem FROM dbo.ComboBox ";
            //DataTable dtPhongThiNgiem = DataProvider.Instance.ExecuteQuery(query);
            //var lsPTN = dtPhongThiNgiem.AsEnumerable().Select(r => r.Field<string>("PhongThiNghiem")).ToList();

            //for (int i = 0; i < lsPTN.Count; i++)
            //{
            //    var query1 = from data in VanKienQuaHan.AsEnumerable() where data.Field<string>("PTN") =


            //}

            //for (int i = 0; i < dtVanKienQuaHan.Rows.Count; i++)
            //{
            //    string nguoiNhan = dtVanKienQuaHan.Rows[i][5].ToString();
            //    string noiDung = dtVanKienQuaHan.Rows[i][0].ToString() + " - " + dtVanKienQuaHan.Rows[i][1].ToString() + " - " + dtVanKienQuaHan.Rows[i][2].ToString() + " : 要更新，請注意！";

            //    if (nguoiNhan != string.Empty)
            //    {
            //       // MainAsync(nguoiNhan, noiDung).Wait();
            //    }
            //}


        }
        static async Task MainAsync(string NguoiNhan, string NoiDung, string PTN, string filePath)
        {
            //Notes
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("http://10.199.1.32:1234");
               // var fileLocations = new List<string>() { @"D:\CSharp\20. SenLineAndNote\SenLineAndNote\SenLineAndNote\bin\Debug\Report-非破壞實驗室.xlsx" };
                var fileLocations = new List<string>() { filePath };
                var mail = new Mail()
                {
                    To = "VNW0014732@VNFPG",
                    //CC = "VNW000XXX@VNFPG",
                    Subject = $"針對即將或已過期之文件，請貴單位執行更新或進版 -- {PTN}",
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
            //using (var client = new HttpClient())
            //{
            //    client.BaseAddress = new Uri("http://10.199.1.32:1234");

            //    var fileLocations = new List<string>() { filePath };

            //    var mail = new Mail()
            //    {
            //       // To = NguoiNhan,
            //        To = "VNW0014732@VNFPG",
            //        //CC = "VNW000XXX@VNFPG",
            //        Subject = $"針對即將或已過期之文件，請貴單位執行更新或進版 -- {PTN}",
            //        Content = NoiDung,
            //        Attachments = new List<AttachmentFile>()
            //    };
            //    var json_string = Newtonsoft.Json.JsonConvert.SerializeObject(mail);
            //    var requestContent = new StringContent(json_string, Encoding.UTF8, "application/json");
            //    var response = await client.PostAsync("/api/Mail", requestContent);
            //    var responseContent = await response.Content.ReadAsStringAsync();
            //    Console.WriteLine(responseContent);

            //    foreach (var fileLocation in fileLocations)
            //    {
            //        var file = File.Open(fileLocation, FileMode.Open);
            //        var file_byteCode = new byte[file.Length];
            //        file.Read(file_byteCode, 0, (int)file.Length);
            //        var file_string = Convert.ToBase64String(file_byteCode);
            //        mail.Attachments.Add(
            //            new AttachmentFile()
            //            {
            //                Name = file.Name.Substring(file.Name.LastIndexOf('\\') + 1),
            //                FileOfBase64String = file_string
            //            });
            //    }
            //}


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
            public int ChuKy { get; set; }
            public DateTime NgayPhatHanh { get; set; }
        }
    }
}
