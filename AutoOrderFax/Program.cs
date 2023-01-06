﻿using System;
using System.Data.SqlClient;
using System.Net;
using System.Text;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Data;
using FluentFTP;
using System.Windows.Documents;
using System.Windows.Markup;
using System.IO.Packaging;
using System.Windows.Xps.Packaging;
using System.Windows.Xps;

namespace AutoOrderFax
{

    public class Program
    {
        public static SqlConnection Connection { set; get; }

        private static string _logFileDirectory;
        private static string _connectionString;
        private static string _outputDirectory;
        private static string _outputMode;
        private static string _ftpHostName;
        private static string _ftpUser;             // FTPサーバ UserID
        private static string _ftpPassword;         // FTPサーバ Password
        //private static string _reqUser;             // サービス UserID
        //private static string _reqPassword;         // サービス Password
        //private static string _mailAddress;         // 通知返信先アドレス
        //private static string _retries;             // リトライ回数
        //private static string _retryInterval;       // リトライ間隔（分）
        //private static string _jikan;               // 時間指定（yyyymmddHHMM）
        //private static string _quality;             // 
        //private static string _paperSize;           // 
        //private static string _direction;           // 
        //private static string _fontSize;            // 
        //private static string _fontType;            // 
        private static RequestContents _rc;
        private static string _lineCount;               // 
        private static string _query;               // 
        private static string _logRetentionDays;         // 

        [STAThread]
        public static void Main(string[] args)
        {
            string currentDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            // 各種設定読み込み
            if (LoadSettings() == false) return;

            // ログファイル
            StreamWriter sw = new StreamWriter(_logFileDirectory + DateTime.Now.ToString("yyyyMMdd_HHmmss") + @".log", true, System.Text.Encoding.GetEncoding("Shift_JIS"));
            Console.SetOut(sw); // 出力先を設定
            Console.WriteLine(@"Output Path... " + _outputDirectory);
            Console.WriteLine(@"Output Mode... " + _outputMode);
            Console.WriteLine(@"FTP Server... " + _ftpHostName);

            try
            {
                ConnectDatabase();
                if (CreateFiles(GetOrderNoList()))
                {
                    // PDFファイルが作成されたら転送
                    CleanRemoteFiles();
                    TransferFiles();
                    DeleteLocalFiles();
                }
                DeleteLogFiles();

            }
            catch (Exception e)
            {
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + e.ToString());
            }
            finally
            {
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Bye.");
                sw.Dispose();
            }

        }

        private static bool LoadSettings()
        {
            string currentDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            XElement xml = XElement.Load(currentDirectory + "Config.xml");
            _logFileDirectory = currentDirectory + @"log\";
            _connectionString = xml.Element("ConnectionString").Value.Trim();
            _outputDirectory = currentDirectory + @"pdf\";
            _outputMode = xml.Element("OutputMode").Value.Trim();
            _logRetentionDays = xml.Element("LogRetentionDays").Value.Trim();
            _lineCount = xml.Element("LineCount").Value;    
            _query = xml.Element("Query").Value;    // 発注データ取得クエリ

            var ftp = (from f in xml.Elements("FTPServer") select f).First();
            _ftpHostName = ftp.Element("Host").Value;
            _ftpUser = ftp.Element("User").Value;
            _ftpPassword = ftp.Element("Password").Value;

            var req = (from f in xml.Elements("Request") select f).First();
            _rc = new RequestContents();
            _rc.ReqUser = req.Element("User").Value;
            _rc.ReqPassword = req.Element("Password").Value;
            _rc.MailAddress = req.Element("MailAddress").Value;
            _rc.Retries = req.Element("Retries").Value;
            _rc.RetryInterval = req.Element("RetryInterval").Value;
            _rc.Jikan = req.Element("Jikan").Value;
            _rc.Quality = req.Element("Quality").Value;
            _rc.PaperSize = req.Element("PaperSize").Value;
            _rc.Direction = req.Element("Direction").Value;
            _rc.FontSize = req.Element("FontSize").Value;
            _rc.FontType = req.Element("FontType").Value;

            // ログ出力先設定
            if (Directory.Exists(_logFileDirectory) == false)
            {
                DirectoryInfo di = new DirectoryInfo(_logFileDirectory);
                di.Create();
            }

            return true;
        }

        private static bool ConnectDatabase()
        {
            Connection = new SqlConnection(_connectionString);
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Connecting to " + Connection.DataSource + @"...");
            Connection.Open();
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"");

            return true;
        }

        private static List<string> GetOrderNoList()
        {
            var res = new List<string>();
            using (SqlCommand command = new SqlCommand(_query, Connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        res = (from IDataRecord r in sdr select r["発注伝票番号"].ToString()).ToList();
                    }
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + res.Count + @" Order(s) found.");
                }
            }
            return (res);
        }

        private static bool CreateFiles(List<string> OrderNoList)
        {
            if (OrderNoList.Count == 0)
            {
                return false;
            }
            foreach (string orderNo in OrderNoList)
            {
                CreateOrderSlip(_outputDirectory, orderNo, _outputMode);

                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Order No. " + orderNo + @" created.");
                UpdateOrderState(orderNo);
            }
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Create PDF, Request Files complete.");

            return true;
        }

        private static void CleanRemoteFiles()
        {
            // リモートのフォルダの中 全削除
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Cleaning remote debris...");
            using (var fc = new FtpClient(_ftpHostName, _ftpUser, _ftpPassword))
            {
                fc.Connect();
                foreach (FtpListItem item in fc.GetListing("", FtpListOption.Recursive))
                {
                    if (item.FullName.IndexOf(@"exclusive.lock") < 0)
                    {
                        fc.DeleteFile(item.Name);
                    }
                }
                // リモートのlock削除
                if (fc.FileExists(@"exclusive.lock"))
                {
                    fc.DeleteFile(@"exclusive.lock");
                }
            }
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");

        }

        private static bool TransferFiles()
        {
            string[] FileList;

            // ローカルにlock作成
            File.WriteAllText(_outputDirectory + @"\exclusive.lock", "");
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"exclusive.lock created.");

            using (WebClient wc = new WebClient())
            {
                wc.Credentials = new NetworkCredential(_ftpUser, _ftpPassword);

                // リモートにlock転送
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfering exclusive.lock ...");
                wc.UploadFile(string.Format("{0}/{1}", _ftpHostName, "exclusive.lock"), _outputDirectory + @"\exclusive.lock");   // ファイルアップロード.
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");

                // PDFファイル転送
                FileList = Directory.GetFiles(_outputDirectory, "*.pdf", System.IO.SearchOption.AllDirectories);
                if (FileList.Length > 0)
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfer PDF Start.");
                    foreach (string FilePath in FileList)
                    {
                        string FileName = System.IO.Path.GetFileName(FilePath);                 // ファイル名取得.
                        wc.UploadFile(string.Format("{0}/{1}", _ftpHostName, FileName), FilePath);   // ファイルアップロード.

                        Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + FileName + @" transferd.");
                    }
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfer PDF complete.");
                }

                // Requestファイル転送
                FileList = Directory.GetFiles(_outputDirectory, "*.req", System.IO.SearchOption.AllDirectories);
                if (FileList.Length > 0)
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfer Request Start.");
                    foreach (string FilePath in FileList)
                    {
                        string FileName = System.IO.Path.GetFileName(FilePath);                 // ファイル名取得.
                        wc.UploadFile(string.Format("{0}/{1}", _ftpHostName, FileName), FilePath);   // ファイルアップロード.

                        Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + FileName + @" transferd.");
                    }
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfer Request complete.");
                }

                // リモートのlock削除
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Deleting remote exclusive.lock ...");
                using (var fc = new FtpClient(_ftpHostName, _ftpUser, _ftpPassword))
                {
                    fc.Connect();
                    if (fc.FileExists(@"exclusive.lock"))
                    {
                        fc.DeleteFile(@"exclusive.lock");
                        Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");
                    }
                    else
                    {
                        Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Not Found.");
                    }
                }
            }
            return true;
        }
        private static void CreateOrderSlip(string OutputDirectory, string OrderNo, string OutputMode)
        {
            OrderHeaderModel ohm = new OrderHeaderModel();

            string sql = "PD発注_注文書 " + OrderNo.ToString();
            using (SqlCommand command = new SqlCommand(sql, Connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        while (sdr.Read())
                        {
                            if ((Int16)sdr["発注行番号"] == 1)
                            {

                                ohm.SupplierName = sdr["仕入先名"].ToString();
                                ohm.CustomerName = sdr["学校名"].ToString();
                                ohm.OperatorName = sdr["入力者名"].ToString();

                                switch ((Int16)sdr["納入区分"])
                                {
                                    case 1:
                                        ohm.DeliveryTypeName = sdr["倉庫名"].ToString() + " 入れ";
                                        ohm.CustomerAddress = sdr["支店住所1"].ToString() + sdr["支店住所2"].ToString();
                                        ohm.CustomerTel = sdr["支店TEL"].ToString();
                                        break;
                                    case 2:
                                        Warehouse w = new Warehouse(_connectionString);
                                        ohm.DeliveryTypeName = w.Name + " 入れ";
                                        ohm.CustomerAddress = w.Address;
                                        ohm.CustomerTel = w.Tel;
                                        break;
                                    case 4:
                                        ohm.DeliveryTypeName = sdr["倉庫名"].ToString() + " 入れ";
                                        ohm.CustomerAddress = "";
                                        ohm.CustomerTel = "";
                                        break;
                                    case 5:
                                        ohm.DeliveryTypeName = "直送";
                                        ohm.CustomerAddress = sdr["学校住所1"].ToString() + sdr["学校住所2"].ToString();
                                        ohm.CustomerTel = sdr["学校TEL"].ToString();
                                        break;
                                    default:
                                        break;

                                }

                                ohm.OrderNo = "J" + sdr["発注表示番号"].ToString();
                                ohm.OrderDate = DateTime.ParseExact(sdr["発注伝票日付"].ToString(), "yyyyMMdd", null).ToString("yyyy年 MM月 dd日");
                                ohm.PrivateNotes = sdr["発注社内伝票摘要"].ToString();
                                ohm.PublicNotes = sdr["発注社外伝票摘要"].ToString();
                                ohm.SelfZipCode = sdr["支店郵便番号"].ToString();
                                ohm.SelfAddress = sdr["支店住所1"].ToString() + sdr["支店住所2"].ToString();
                                CompanyInfo c = new CompanyInfo(_connectionString);
                                ohm.SelfCompanyName = c.Name;
                                ohm.SelfDepartmentName = sdr["支店名称"].ToString();
                                ohm.SelfTel = sdr["支店TEL"].ToString();
                                ohm.SelfFax = sdr["支店FAX"].ToString();
                                ohm.ShippingDate = sdr["出荷日付"].ToString();
                                ohm.OrderNoTimeStamp = "No.H" + sdr["発注伝票番号"].ToString() + "-" + DateTime.Now.ToString("yyyyMMddHHmmss");
                            }

                            OrderDetailModel odm = new OrderDetailModel();
                            odm.LineNo = (Int16)sdr["発注行番号"];

                            // 入ってたら敬称付加
                            odm.TeacherName = (sdr["先生名"].ToString() != "") ? sdr["先生名"].ToString() + "先生" : "";
                            odm.IndividualName = (sdr["個人名"].ToString() != "") ? sdr["個人名"].ToString() + "様" : "";
                            // 入ってたら年組付加
                            odm.SchoolYear = (sdr["学年"].ToString() != "0") ? sdr["学年"].ToString() + "年" : "";
                            odm.SchoolClass = (sdr["組"].ToString() != "") ? sdr["組"].ToString() + "組" : "";

                            odm.ItemName = sdr["商品名"].ToString() + sdr["規格名"].ToString();
                            odm.ItemCode = sdr["商品コード"].ToString();
                            odm.Qty = float.Parse(sdr["数量"].ToString());
                            odm.ReserveQty = float.Parse(sdr["予備数量"].ToString());
                            odm.TeacherQty = float.Parse(sdr["教師数量"].ToString());
                            odm.UnitPrice = float.Parse(sdr["税抜仕入単価"].ToString());

                            var sourceText = "";
                            for (int i = 1; i <= 8; i++)
                            {
                                sourceText += sdr["クラス0" + i.ToString()].ToString();
                            }
                            odm.ClassDivide = ClassDivideEncoder.Encode(sdr["クラス分け入力区分"].ToString(), sourceText);
                            odm.LinePrivateNotes = sdr["発注社内明細摘要"].ToString();
                            odm.LinePublicNotes = sdr["発注社外明細摘要"].ToString();

                            ohm.OrderDetails.Add(odm);
                        }
                    }
                }
            }

            switch (OutputMode)
            {
                case "PDF":
                    {
                        var pf = new PdfCreator(OrderNo);
                        pf.RequestContents = _rc;
                        pf.sql = sql;
                        pf.ConnectionString = _connectionString;
                        pf.ohm = ohm;
                        pf.Create(OutputDirectory);

                        break;
                    }
                case "PRINTER":
                    {
                        //pageReport.Document.Printer.PrinterSettings.PrinterName = "DocuPrint P350 d";
                        //pageDocument.PageReport.Document.Print(false, false, false);

                        break;
                    }
            }

            return;
        }
        private static void UpdateOrderState(string OrderNo)
        {
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Updating order state...");
            string sql = "UPDATE D発注FAX "
                        + "SET "
                        + " 送信Mode = 1 "
                        + "FROM "
                        + " D発注 H "
                        + " INNER JOIN D発注FAX F "
                        + "  ON H.発注伝票番号 = F.発注伝票番号 "
                        + "  AND H.発注伝票番号 = '" + OrderNo + "'";
            using (SqlCommand command = new SqlCommand(sql, Connection))
            {
                command.ExecuteNonQuery();
            }
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");

        }

        private static void DeleteLocalFiles()
        {
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Deleting local transfered files...");
            foreach (string pathFrom in System.IO.Directory.EnumerateFiles(_outputDirectory, "*.*"))
            {
                File.Delete(pathFrom);
            }
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");
        }

        private static void DeleteLogFiles()
        {
            int count = 0;
            DirectoryInfo di = new DirectoryInfo(_logFileDirectory);
            foreach (FileInfo fi in di.GetFiles())
            {
                // 日付の比較
                if (fi.LastWriteTime < DateTime.Today.AddDays(int.Parse(_logRetentionDays) * -1))
                {
                    fi.Delete();
                    count++;
                }
            }
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + count + @" Log File(s) Deleted.");
        }

    }
}
