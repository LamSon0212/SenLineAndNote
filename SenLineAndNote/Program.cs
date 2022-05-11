using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Net.Http;
using System.Reflection;

namespace SenLineAndNote
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dtVanKienQuaHan = new DataTable();
            DateTime Now = DateTime.Now;

            //string From = "1999-01-01";
            //DateTime ToTime = Now.AddYears(-3);
            string query;




            //query = "SELECT dbo.VanKienIso.MaVanKien,dbo.VanKienIso.TenTiengTrung,dbo.VanKienIso.TenTiengViet,dbo.VanKienIso.NguoiDuaLen,NgayPhatHanh, dbo.ComboBox.NguoiNhacNho FROM dbo.VanKienIso INNER JOIN dbo.ComboBox ON dbo.VanKienIso.PTN = PhongThiNghiem WHERE dbo.VanKienIso.NgayPhatHanh BETWEEN N'" + From + "' AND N'" + ToTime + "'";
            //dtVanKienQuaHan = DataProvider.Instance.ExecuteQuery(query);

            //List<VanKienQuaHan> lsVanKienQuaHan = new List<VanKienQuaHan>();

            //VanKienQuaHan Row_VanKienQuaHan = new VanKienQuaHan();




            query = "SELECT	* FROM	dbo.VanKienIso";
            dtVanKienQuaHan = DataProvider.Instance.ExecuteQuery(query);

            List<VanKienIso> lsVanKienIsos = new List<VanKienIso>();
            foreach (DataRow dr in dtVanKienQuaHan.Rows)
            {
                VanKienIso vanKienIso = new VanKienIso();
                foreach (PropertyInfo objProperty in vanKienIso.GetType().GetProperties())
                {
                    if (dtVanKienQuaHan.Columns.Contains(objProperty.Name) && objProperty.PropertyType == dr[objProperty.Name].GetType())
                    {
                        objProperty.SetValue(vanKienIso, dr[objProperty.Name], null);
                    }
                }
                lsVanKienIsos.Add(vanKienIso);
            }

            var VanKienQuaHan1 = (from temp in lsVanKienIsos
                                  where temp.NgayPhatHanh.AddMonths(temp.ChuKy) < DateTime.Today
                                  select temp).ToList();
            var VanKienQuaHan = lsVanKienIsos.Where(r => r.NgayPhatHanh.AddMonths(r.ChuKy) < DateTime.Today).ToList();


            var lsPTN = lsVanKienIsos.Select(r => r.PTN).Distinct().ToList();

            foreach (var item in lsPTN)
            {
                string NguoiNhacNho = "";
                if (item == string.Empty) continue;

                query = $"SELECT NguoiNhacNho FROM dbo.ComboBox WHERE PhongThiNghiem = N'{item}'";
                NguoiNhacNho = DataProvider.Instance.ExecuteScalar(query).ToString();
                var PhongThiNghiem = VanKienQuaHan.Where(r => r.PTN == item).ToList();

                //%0D%0A
                string msg = $"<h1>這些文件已過期,請更新!</h1>";
                //string msg = "<p>";
                int STT = 1;

                string space100 = "*************************************************************************************************************** ";
                string br = "<br/>";

                foreach (var data in PhongThiNghiem)
                {
                    msg += $"<p>{ STT.ToString("00") }. {data.MaVanKien}{ br }&emsp;&ensp;{data.TenTiengTrung}{ br }&emsp;&ensp;{data.TenTiengViet}</p>\r";
                    STT++;
                }
                msg += $"<p>{space100}{space100}</p>";
                if (NguoiNhacNho != string.Empty)
                {
                    MainAsync(NguoiNhacNho, msg, item).Wait();
                    Console.WriteLine(item);
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
        static async Task MainAsync(string NguoiNhan, string NoiDung, string PTN)
        {
            //Notes

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("http://10.199.1.32:1234");
                var mail = new Mail()
                {
                    To = NguoiNhan,
                    //CC = "VNW000XXX@VNFPG",
                    Subject = $"文件管理系統通知 -- {PTN}",
                    Content = NoiDung,
                    Attachments = new List<AttachmentFile>()
                };
                var json_string = Newtonsoft.Json.JsonConvert.SerializeObject(mail);
                var requestContent = new StringContent(json_string, Encoding.UTF8, "application/json");
                var response = await client.PostAsync("/api/Mail", requestContent);
                var responseContent = await response.Content.ReadAsStringAsync();
                Console.WriteLine(responseContent);
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
            public int ChuKy { get; set; }
            public DateTime NgayPhatHanh { get; set; }
        }
    }
}
