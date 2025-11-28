using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace UpdateCouponMonthly
{
    class CommonUtility
    {
        public void Txt(string ex)
        {
            //string sourcePath = @"C:\Website\UpdateCouponMonthly\Log" + @"\" + DateTime.Now.ToString("yyyy-MM-dd.HH") + "狀態.txt";       //正式
            //string sourcePath = @"D:\Website\UpdateCouponMonthly\Log\" + DateTime.Now.ToString("yyyy-MM-dd.HH") + "狀態.txt";     //測試
            string sourcePath = @"E:\POJHIH\Program\UpdateCouponMonthly\UpdateCouponMonthly\bin\Debug\Log\" + DateTime.Now.ToString("yyyy-MM-dd.HH") + "狀態.txt";    //本地
            #region  寫入txt檔
            string txt = "";
            if (File.Exists(sourcePath))
            {
                #region 取TXT檔內文串
                StreamReader str = new StreamReader(sourcePath);
                //str.ReadLine(); (一行一行讀取)
                txt = str.ReadToEnd();//(一次讀取全部)
                str.Close(); //(關閉str)
                #endregion
                #region 寫入txt檔
                using (StreamWriter sw = new StreamWriter(sourcePath))
                {
                    sw.WriteLine(txt + DateTime.Now + "-" + ex);
                    sw.Close();
                }
                #endregion
            }
            else
            {
                #region 創建txt檔
                FileStream fileStream = new FileStream(sourcePath, FileMode.Create);
                fileStream.Close();
                #endregion
                #region 寫入txt檔
                using (StreamWriter sw = new StreamWriter(sourcePath))
                {
                    sw.WriteLine(DateTime.Now + "-" + ex);
                    sw.Close();
                }
                #endregion
            }
            #endregion
        }
        /// <summary>
        ///發送mail
        /// </summary>
        //SendMail("主旨",信件內容,寄件者信箱,收件者信箱)
        public static void SendMail()
        {
            CommonUtility commonUtility = new CommonUtility();
            try
            {
                commonUtility.Txt("Function: SendMail Start");
                //string folderPath = $@"C:\Website\CouponSAP\file\{DateTime.Now.ToString("yyyyMMdd")}"; //正式
                string folderPath = $@"E:\POJHIH\Website\CouponSAP\file\{DateTime.Now.ToString("yyyyMMdd")}"; //本地

                clsINI ini = new clsINI(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.ini"));
                string MailFrom, SmtpClient, Port, RealEMailAddress, RealEMailPassWord;
                MailFrom = ini.IniReadValue("EmailSetting", "MailFrom");
                SmtpClient = ini.IniReadValue("EmailSetting", "SmtpClient");
                Port = ini.IniReadValue("EmailSetting", "Port");
                RealEMailAddress = ini.IniReadValue("EmailSetting", "RealEMailAddress");
                RealEMailPassWord = ini.IniReadValue("EmailSetting", "RealEMailPassWord");
                MailMessage mail = new MailMessage();
                //寄件者
                mail.From = new MailAddress(MailFrom, "抵用券排程");
                // 新增收件人
                using (StreamReader sw = new StreamReader((Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "toEmail.txt")), true))
                {
                    string sl = sw.ReadLine();
                    while (sl != null)
                    {
                        mail.To.Add(sl);
                        sl = sw.ReadLine();
                    }
                }
                // 新增副本收件人
                using (StreamReader sw = new StreamReader((Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ccEmail.txt")), true))
                {
                    string sl = sw.ReadLine();
                    while (sl != null)
                    {
                        mail.CC.Add(sl);
                        sl = sw.ReadLine();
                    }
                }

                string[] files = Directory.GetFiles(folderPath); //讀取資料夾內檔案

                if (files.Length == 0)
                {
                    throw new Exception("未找到檔案");
                }
                // 新增附件
                foreach (string file in files)
                {
                    mail.Attachments.Add(new Attachment(file));
                }
                DateTime _before = DateTime.Now.AddMonths(-1);
                //主旨
                mail.Subject = $"{_before.Month}月抵用券報表";
                //是否採用HTML格式
                mail.IsBodyHtml = true;
                //信件內容
                mail.Body = $@"<div>
                                    <p>附件為</p>
                                    <p>{_before.Year}年{_before.Month}月抵用券兌回資料</p>
                                    <p>{_before.Year}年{_before.Month}月抵用券流通在外的有效券資料(遞延)</p>
                               </div>";
                //編碼
                mail.BodyEncoding = Encoding.UTF8;

                SmtpClient SC = new SmtpClient(SmtpClient)
                {
                    Port = Convert.ToInt32(Port),
                    EnableSsl = true,
                    Credentials = new System.Net.NetworkCredential(RealEMailAddress, RealEMailPassWord)
                };
                SC.Send(mail);
            }
            catch (Exception ex)
            {
                string ErrorMessage = "寄信失敗，錯誤訊息：" + ex;
                commonUtility.Txt(ErrorMessage);
            }
            finally
            {
                commonUtility.Txt("Function: SendMail end");
            }
        }
        /// <summary>
        /// 移動檔案到當日資料夾
        /// </summary>
        public static void MoveFiles()
        {
            CommonUtility commonUtility = new CommonUtility();
            try
            {
                //string folderPath = @"C:\Website\CouponSAP\file"; //正式
                string folderPath = @"E:\POJHIH\Website\CouponSAP\file"; //本地
                string destinationFolder = $@"{folderPath}\{DateTime.Now.ToString("yyyyMMdd")}"; // 目標資料夾路徑
                string[] extensions = { "*.", "*兌出-兌回.xlsx" };// 欲抓取副檔名

                // 檢查目標資料夾是否存在，若不存在則建立
                if (!Directory.Exists(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                }
                else //已存在資料夾則將舊的移到_1...
                {
                    int counter = 1;
                    string newFolderPath = $@"{folderPath}\{DateTime.Now.ToString("yyyyMMdd")}_{counter}";
                    if (Directory.Exists(destinationFolder))
                    {
                        while (Directory.Exists(newFolderPath))
                        {
                            counter++;
                            newFolderPath = $@"{folderPath}\{DateTime.Now.ToString("yyyyMMdd")}_{counter}";
                        }
                        Directory.Move(destinationFolder, newFolderPath); //修改已存在資料夾名稱 _1...
                        Directory.CreateDirectory(destinationFolder);
                    }
                }

                foreach (string extension in extensions)
                {
                    string[] files = Directory.GetFiles(folderPath, extension); //讀取資料夾內檔案
                    //var files = Directory.EnumerateFiles(folderPath, extension); //讀取資料夾內檔案
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);
                        if (DateTime.Now.ToString("yyyyMMdd") == fileInfo.CreationTime.ToString("yyyyMMdd") ||
                            (extension == "*兌出-兌回.xlsx" && DateTime.Now.ToString("yyyyMMdd") == fileInfo.LastWriteTime.ToString("yyyyMMdd"))) //xlsx多判斷檔案修改日期是否為今天
                        {
                            string destinationFilePath = Path.Combine(destinationFolder, Path.GetFileName(file)); // 目標檔案完整路徑
                            string fileName = Path.GetFileNameWithoutExtension(file); // 檔案名稱（不含副檔名）

                            //// 檢查目標檔案是否已存在
                            //int counter = 1;
                            ////string newFilePath = destinationFilePath;
                            //string newFileName = $"{fileName}_{counter}{Path.GetExtension(file)}";// 加 _1、_2
                            //string newFilePath = Path.Combine(destinationFolder, newFileName);
                            //if (File.Exists(destinationFilePath))
                            //{
                            //    while (File.Exists(newFilePath))
                            //    {
                            //        counter++;
                            //        newFileName = $"{fileName}_{counter}{Path.GetExtension(file)}";// 加 _1、_2
                            //        newFilePath = Path.Combine(destinationFolder, newFileName);
                            //    }
                            //    File.Move(destinationFilePath, newFilePath);  // 原有檔案改名_1、_2
                            //}

                            if (extension == "*.")
                            {
                                // 複製檔案（第二個參數表示是否覆寫已存在檔案）
                                File.Copy(file, destinationFilePath + ".txt", overwrite: true);
                            }
                            else
                            {
                                // 移動檔案
                                File.Move(file, destinationFilePath);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string Failedmessage = "移動報表檔案發生錯誤，錯誤訊息：" + ex;
                commonUtility.Txt(Failedmessage);
            }
        }
    }
}