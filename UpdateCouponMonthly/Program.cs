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
            if (ConfigurationManager.AppSettings["DeBUG"] == "Y")
            {
                //回寫兌出起訖時間
                UpdateCoupon.UpdateExchangeTime();
                UpdateCoupon.UpdateSAP_B_Time();
            }
            else
            {
                if (DateTime.Now.Day == DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)) //當月最後一天
                {
                    //回寫兌出起訖時間
                    UpdateCoupon.UpdateExchangeTime();
                    UpdateCoupon.UpdateSAP_B_Time();
                }
                if (DateTime.Now.Day == 1) //當月第一天
                {
                    //先將CouponSAP產出的檔案移至當日的資料夾
                    CommonUtility.MoveFiles();
                    if (ConfigurationManager.AppSettings["SendMail"] == "Y")
                    {
                        CommonUtility.SendMail();
                    }
                }
            }
        }
    }
}
