using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SenLineAndNote
{
    public static class ECB
    {
        public static bool ToECB(this Int64 DateTimeECB)
        {
            bool Job = false;
            string DateTimeNow = DateTime.Now.Year.ToString()
                 + string.Format("{0:00}", DateTime.Now.Month).ToString()
                 + string.Format("{0:00}", DateTime.Now.Day).ToString();

            if (Convert.ToInt64(DateTimeNow) > DateTimeECB)
            {
                Job = true;
            }

            return Job;
        }
    }
}
