using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace UpdateCouponMonthly
{
    class UpdateCoupon
    {

        /// <summary>
        /// 回寫UsedRule兌出起訖時間
        /// </summary>
        public static void UpdateExchangeTime()
        {
            CommonUtility commonUtility = new CommonUtility();
            try
            {
                DB_Connection _db = new DB_Connection();
                List<SqlParameter> _Parameter = new List<SqlParameter>();
                DateTime _after = DateTime.Now.AddMonths(+1);
                DateTime _sDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);//月初
                DateTime _eDate = new DateTime(_after.Year, _after.Month, 1);//月底(隔月第一天)
                if (ConfigurationManager.AppSettings["DeBUG"] == "Y")
                {
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["startDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateS);
                    _sDate = dateS;
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["endDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateE);
                    _eDate = dateE;
                }
                string _SQL = @"
CREATE TABLE #temp_change_log (
    GId INT,
    MallId VARCHAR(5),
    OldExchangeStart DATETIME,
    OldExchangeEnd DATETIME,
    ExchangeStart DATETIME,
    ExchangeEnd DATETIME,
);
                                --回寫每個月遞延與兌回報表的兌出起訖時間
								MERGE INTO UsedRule AS UR
                                USING (select GId,MallId,MIN(ExchangeStart) as FixExchangeStart,MAX(ExchangeEnd) as FixExchangeEnd
			                                FROM ExchangeMall EM left join Exchange_Gifts EG on EM.ERId=EG.ERId 
			                                where GId in (
                                                --遞延&兌回
												SELECT GId  FROM [Gifts].[dbo].[Coupon] C
													Where c.CreateOn >=@SDate and c.CreateOn <@EDate And Status = 'U' and Type='C' 
													group by GId
												union
                                                --遞延
												select GId from Coupon C Left join Gifts G on C.GId = G.Id
													Where c.CreateOn >=@SDate and c.CreateOn <@EDate And G.Type = 'C' And Status != 'F'  And SAPType <> 'B'
													group by GId
			                                )
		                                group by GId,MallId) AS EM
                                ON  UR.GId = EM.GId AND
                                    UR.MallId = EM.MallId
                                WHEN MATCHED AND 
									(UR.ExchangeStart IS NULL OR
									 UR.ExchangeEnd IS NULL) THEN
	                                 UPDATE SET UR.ExchangeStart = CASE WHEN UR.ExchangeStart IS NULL THEN EM.FixExchangeStart ELSE UR.ExchangeStart END,
												UR.ExchangeEnd = CASE WHEN UR.ExchangeEnd IS NULL THEN EM.FixExchangeEnd ELSE UR.ExchangeEnd END
								OUTPUT 
									deleted.GId, 
									deleted.MallId, 
									deleted.ExchangeStart, 
									deleted.ExchangeEnd, 
									inserted.ExchangeStart, 
									inserted.ExchangeEnd
								INTO #temp_change_log (GId, MallId, OldExchangeStart, OldExchangeEnd, ExchangeStart, ExchangeEnd);
--變更紀錄
select * from #temp_change_log order by GId,MallId

DROP TABLE #temp_change_log";
                _Parameter.Add(new SqlParameter("SDate", _sDate));
                _Parameter.Add(new SqlParameter("EDate", _eDate));

                //_db.SQLExecute(_SQL, _Parameter);
                DataTable changeLog = _db.GetDataTable(_SQL, _Parameter);
                commonUtility.Txt(@"/********************回寫UsedRule兌出起訖時間********************\");
                commonUtility.Txt($"{string.Format("{0,6}", "GId")} {string.Format("{0,6}", "MallId")} " +
                        $"{string.Format("{0,27}", "OldExchangeStart")} {string.Format("{0,27}", "OldExchangeEnd")} " +
                        $"{string.Format("{0,27}", "ExchangeStart")} {string.Format("{0,27}", "ExchangeEnd")}");
                foreach (var item in changeLog.Select())
                {
                    commonUtility.Txt($"{string.Format("{0,6}", item["GId"].ToString())} {string.Format("{0,6}", item["MallId"].ToString())} " +
                        $"{string.Format("{0,25}", item["OldExchangeStart"].ToString())} " +
                        $"{string.Format("{0,25}", item["OldExchangeEnd"].ToString())} " +
                        $"{string.Format("{0,25}", item["ExchangeStart"].ToString())} {string.Format("{0,25}", item["ExchangeEnd"].ToString())}");
                }
            }
            catch (Exception ex)
            {
                string Failedmessage = "回寫兌出起訖時間發生錯誤，錯誤訊息：" + ex;
                commonUtility.Txt(Failedmessage);
            }
            finally
            {
                commonUtility.Txt(@"/****************************************************************\");
            }
        }
        /// <summary>
        /// 回寫無償券(SAPType B)兌出起訖時間
        /// </summary>
        public static void UpdateSAP_B_Time()
        {
            CommonUtility commonUtility = new CommonUtility();
            try
            {
                DB_Connection _db = new DB_Connection();
                List<SqlParameter> _Parameter = new List<SqlParameter>();
                DateTime _after = DateTime.Now.AddMonths(+1);
                DateTime _sDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);//月初
                DateTime _eDate = new DateTime(_after.Year, _after.Month, 1);//月底(隔月第一天)
                if (ConfigurationManager.AppSettings["DeBUG"] == "Y")
                {
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["startDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateS);
                    _sDate = dateS;
                    DateTime.TryParseExact(ConfigurationManager.AppSettings["endDate"], "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateE);
                    _eDate = dateE;
                }

                string _SQL = @"
CREATE TABLE #temp_change_log (
    GId INT,
    MallId VARCHAR(5),
    OldExchangeStart DATETIME,
    OldExchangeEnd DATETIME,
    ExchangeStart DATETIME,
    ExchangeEnd DATETIME,
);		                        
                                --直接用Coupon UsedStart UsedEnd 當作起訖時間
								MERGE INTO UsedRule AS UR
                                USING (select GId,C.MallId,MIN(C.UsedStart) as FixExchangeStart,MAX(C.UsedEnd) as FixExchangeEnd 
										from Coupon C Left join Gifts G on C.GId = G.Id where SAPType='B' and 
											c.CreateOn >=@SDate and c.CreateOn <@EDate 
										group by GId,C.MallId) AS C
                                ON  UR.GId = C.GId AND
                                    UR.MallId = C.MallId
                                WHEN MATCHED AND 
									(UR.ExchangeStart IS NULL OR
									 UR.ExchangeEnd IS NULL) THEN
	                                 UPDATE SET UR.ExchangeStart = CASE WHEN UR.ExchangeStart IS NULL THEN C.FixExchangeStart ELSE UR.ExchangeStart END,
												UR.ExchangeEnd = CASE WHEN UR.ExchangeEnd IS NULL THEN C.FixExchangeEnd ELSE UR.ExchangeEnd END
								OUTPUT 
									deleted.GId, 
									deleted.MallId, 
									deleted.ExchangeStart, 
									deleted.ExchangeEnd, 
									inserted.ExchangeStart, 
									inserted.ExchangeEnd
								INTO #temp_change_log (GId, MallId, OldExchangeStart, OldExchangeEnd, ExchangeStart, ExchangeEnd);

--變更紀錄
select * from #temp_change_log order by GId,MallId

DROP TABLE #temp_change_log";
                _Parameter.Add(new SqlParameter("SDate", _sDate));
                _Parameter.Add(new SqlParameter("EDate", _eDate));

                DataTable changeLog = _db.GetDataTable(_SQL, _Parameter);
                commonUtility.Txt(@"/*********************回寫無償券兌出起訖時間*********************\");
                commonUtility.Txt($"{string.Format("{0,6}", "GId")} {string.Format("{0,6}", "MallId")} " +
                        $"{string.Format("{0,27}", "OldExchangeStart")} {string.Format("{0,27}", "OldExchangeEnd")} " +
                        $"{string.Format("{0,27}", "ExchangeStart")} {string.Format("{0,27}", "ExchangeEnd")}");
                foreach (var item in changeLog.Select())
                {
                    commonUtility.Txt($"{string.Format("{0,6}", item["GId"].ToString())} {string.Format("{0,6}", item["MallId"].ToString())} " +
                        $"{string.Format("{0,25}", item["OldExchangeStart"].ToString())} " +
                        $"{string.Format("{0,25}", item["OldExchangeEnd"].ToString())} " +
                        $"{string.Format("{0,25}", item["ExchangeStart"].ToString())} {string.Format("{0,25}", item["ExchangeEnd"].ToString())}");
                }
            }
            catch (Exception ex)
            {
                string Failedmessage = "回寫無償券兌出起訖時間發生錯誤，錯誤訊息：" + ex;
                commonUtility.Txt(Failedmessage);
            }
            finally
            {
                commonUtility.Txt(@"/****************************************************************\");
            }
        }
    }
}
