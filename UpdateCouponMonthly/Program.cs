using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace UpdateCouponMonthly
{
    class Program
    {
        static void Main(string[] args)
        {
            if (DateTime.Now.Day == DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)) //當月最後一天
            {
                UpdateCoupon.UpdateExchangeTime();
                UpdateCoupon.UpdateSAP_B_Time();
            }
            if (DateTime.Now.Day == 1) //當月第一天
            {
                CommonUtility.MoveFiles();
                if (ConfigurationManager.AppSettings["SendMail"] == "Y")
                {
                    CommonUtility.SendMail();
                }
            }
        }
    }
}
