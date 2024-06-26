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
using FluentFTP.Proxy;
using FluentFTP.Proxy.AsyncProxy;
using FluentFTP.Proxy.SyncProxy;
using Renci.SshNet;
using static System.Net.WebRequestMethods;

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
        private static string _ftpPort;
        private static string _ftpUser;
        private static string _ftpPassword;
        private static string _proxyHostName;
        private static string _proxyPort;
        private static string _proxyUser;
        private static string _proxyPassword;

        private static RequestContents _rc;
        private static string _linesPerPage;
        private static string _query;
        private static string _logRetentionDays;
        private static string _fixedNotes;

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
                    ConnectionInfo connectionInfo;

                    if (_proxyHostName == "" && _proxyPort == "")
                    {
                        // No proxy server.
                        connectionInfo = new ConnectionInfo(_ftpHostName,
                                                            int.Parse(_ftpPort),
                                                            _ftpUser,
                                                            new PasswordAuthenticationMethod(_ftpUser, _ftpPassword));
                    }
                    else
                    {
                        // create an FTP client connecting through a HTTP 1.1 Proxy
                        connectionInfo = new ConnectionInfo(_ftpHostName,
                                                            int.Parse(_ftpPort),
                                                            _ftpUser,
                                                            ProxyTypes.Http,
                                                            _proxyHostName,
                                                            int.Parse(_proxyPort),
                                                            _proxyUser,
                                                            _proxyPassword,
                                                            new PasswordAuthenticationMethod(_ftpUser, _ftpPassword));
                    }
                    SftpClient fc = new SftpClient(connectionInfo);
                    fc.Connect();
                    // CleanRemoteFiles(fc);
                    TransferFiles(fc);
                    fc.Disconnect();

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
            _linesPerPage = xml.Element("LineCount").Value;
            _query = xml.Element("Query").Value;    // 発注データ取得クエリ
            _fixedNotes = xml.Element("FixedNotes").Value;    // 発注データ取得クエリ

            var ftp = (from f in xml.Elements("FTPServer") select f).First();
            _ftpHostName = ftp.Element("Host").Value;
            _ftpPort = ftp.Element("Port").Value;
            _ftpUser = ftp.Element("User").Value;
            _ftpPassword = ftp.Element("Password").Value;

            var proxy = (from f in xml.Elements("ProxyServer") select f).First();
            _proxyHostName = proxy.Element("Host").Value;
            _proxyPort = proxy.Element("Port").Value;
            _proxyUser = proxy.Element("User").Value;
            _proxyPassword = proxy.Element("Password").Value;

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
                CreateOrderSlipPdf(_outputDirectory, orderNo, _outputMode);

                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Order No. " + orderNo + @" created.");
                UpdateOrderState(orderNo);
            }
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Create PDF, Request Files complete.");

            return true;
        }

        private static void CleanRemoteFiles(SftpClient fc)
        {
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Cleaning remote debris...");

            // リモートのlock以外を先に全削除
            foreach (var ftpfile in fc.ListDirectory("."))
            {
                if (ftpfile.FullName.IndexOf(@"exclusive.lock") < 0)
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Deleting " + ftpfile);
                    fc.DeleteFile(ftpfile.Name);
                }
            }

            // リモートのlock削除
            if (fc.Exists(@"exclusive.lock"))
            {
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Deleting exclusive.lock");
                fc.DeleteFile(@"exclusive.lock");
            }
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");

        }

        private static bool TransferFiles(SftpClient client)
        {
            string[] FileList;
            Stream stream;

            // ローカルにlock作成
            System.IO.File.WriteAllText(_outputDirectory + @"\exclusive.lock", "");
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"exclusive.lock created.");

            // リモートにlock転送
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfering exclusive.lock ...");
            stream = System.IO.File.OpenRead(_outputDirectory + @"\exclusive.lock");
            client.UploadFile(stream, @"exclusive.lock");
            stream.Dispose();
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");

            // PDFファイル転送
            FileList = Directory.GetFiles(_outputDirectory, "*.pdf", System.IO.SearchOption.AllDirectories);
            if (FileList.Length > 0)
            {
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfer PDF Start.");
                foreach (string FilePath in FileList)
                {
                    stream = System.IO.File.OpenRead(FilePath);
                    string FileName = Path.GetFileName(FilePath);                 // ファイル名取得.
                    client.UploadFile(stream, FileName);   // ファイルアップロード.
                    stream.Dispose();
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
                    stream = System.IO.File.OpenRead(FilePath);
                    string FileName = Path.GetFileName(FilePath);                 // ファイル名取得.
                    client.UploadFile(stream, FileName);   // ファイルアップロード.
                    stream.Dispose();
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + FileName + @" transferd.");
                }
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Transfer Request complete.");
            }

            // リモートのlock削除
            Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Deleting remote exclusive.lock ...");
            if (client.Exists(@"exclusive.lock"))
            {
                client.DeleteFile(@"exclusive.lock");
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"OK");
            }
            else
            {
                Console.WriteLine(DateTime.Now.ToString("HH:mm:ss ") + @"Not Found.");
            }


            return true;
        }
        private static int CalculateTotalPages(int totalLines)
        {
            // 総行数を1ページあたりの行数で割る
            int pages = totalLines / int.Parse(_linesPerPage);

            // 割り切れない場合は、最後のページを加える
            if (totalLines % int.Parse(_linesPerPage) != 0)
            {
                pages++; // 最後の部分があるため、ページ数を1増やす
            }

            return pages;
        }

        private static void CreateOrderSlipPdf(string OutputDirectory, string OrderNo, string OutputMode)
        {
            string sql = "";

            // 新しい HeaderModel インスタンスを作成
            var ohm = new OrderHeaderModel
            {
                OrderNo = OrderNo

            };
            var ohms = new List<OrderHeaderModel>();

            // 先に納入区分を見て、PD発注_注文書 か PD発注_注文書_即売 の切り分けをする
            sql = "PD発注_注文書 " + OrderNo.ToString();
            using (SqlCommand command = new SqlCommand(sql, Connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        if (sdr.Read())
                        {
                            if ((Int16)sdr["納入区分"] == 6)
                            {
                                sql = "PD発注_注文書_即売 " + OrderNo.ToString();
                            } else {
                                sql = "PD発注_注文書 " + OrderNo.ToString();
                            }
                        }
                    }
                }
            }

            
            using (SqlCommand command = new SqlCommand(sql, Connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        var orderInfo = new OrderInfo(_connectionString, OrderNo);
                        var pageCount = 0;
                        var allPageCount = CalculateTotalPages(orderInfo.LineCount);

                        while (sdr.Read()) 
                        {
                            if ((Int16)sdr["発注行番号"] == 1)
                            {
                                Headerinitialize(ohm, sdr);
                            }

                            OrderDetailModel odm = new OrderDetailModel();
                            odm.LineNo = (Int16)sdr["発注行番号"];

                            // 入ってたら敬称付加
                            odm.TeacherName = (sdr["先生名"].ToString() != "") ? sdr["先生名"].ToString() + "先生" : "";
                            odm.IndividualName = (sdr["個人名"].ToString() != "") ? sdr["個人名"].ToString() + "様" : "";
                            // 入ってたら年組付加
                            odm.SchoolYear = (sdr["学年"].ToString() != "0") ? sdr["学年"].ToString() + "年" : "";
                            odm.SchoolClass = (sdr["組"].ToString() != "") ? sdr["組"].ToString() + "組" : "";
                            odm.BillingTargetType = int.Parse(sdr["請求区分"].ToString());
                            odm.ItemName = sdr["商品名"].ToString() + sdr["規格名"].ToString();
                            odm.ItemCode = sdr["商品コード"].ToString();
                            odm.Qty = float.Parse(sdr["数量"].ToString());
                            odm.ReserveQty = float.Parse(sdr["予備数量"].ToString());
                            odm.TeacherQty = float.Parse(sdr["教師数量"].ToString());
                            odm.UnitPrice = float.Parse(sdr["税抜仕入単価"].ToString());
                            odm.Price = float.Parse(sdr["税抜仕入金額"].ToString());
                            odm.UnitSalePrice = float.Parse(sdr["売上単価"].ToString());
                            odm.SalePrice = float.Parse(sdr["税込売上金額"].ToString());

                            for (int i = 0; i < 8; i++)
                            {
                                odm.ClassDivide[i] = (Int16)sdr["クラス0" + ((i + 1).ToString())];
                            }
                            odm.LinePrivateNotes = sdr["発注社内明細摘要"].ToString();                           
                            odm.LinePublicNotes = sdr["発注社外明細摘要"].ToString();

                            ohm.OrderDetails.Add(odm);



                            // ページ行数ごとに1ページ分のデータを確定させる
                            if (((Int16)sdr["発注行番号"] % int.Parse(_linesPerPage)) == 0)
                            {
                                pageCount++;
                                ohm.PageDescription = pageCount.ToString() + "/" + allPageCount.ToString();
                                ohms.Add(ohm);

                                // 新しいHeaderModelインスタンスを作成
                                ohm = new OrderHeaderModel
                                {
                                    OrderNo = OrderNo
                                };

                                Headerinitialize(ohm, sdr);

                            }
                        }
                        if (ohm.OrderDetails.Count() != 0)
                        {
                            pageCount++;
                            ohm.PageDescription = pageCount.ToString() + "/" + allPageCount.ToString();
                            ohms.Add(ohm);
                        }

                        // PDF・Requestファイル 出力
                        var pf = new PdfCreator(OrderNo);
                        pf.RequestContents = _rc;
                        pf.sql = sql;
                        pf.ConnectionString = _connectionString;
                        pf.ohms = ohms;
                        pf.CreatePdfAndRequestFiles(OutputDirectory);
                    }
                }
            }
            return;
        }

        private static void Headerinitialize(OrderHeaderModel ohm, SqlDataReader sdr)
        {

            ohm.SupplierName = sdr["仕入先名"].ToString();
            ohm.CustomerCode = (int)sdr["学校コード"];
            ohm.CustomerName = sdr["学校名"].ToString();
            ohm.OperatorName = sdr["操作者名"].ToString();

            // 納入区分による届先の切り分け
            switch ((Int16)sdr["納入区分"])
            {
                case 1:
                    ohm.DeliveryTypeName = sdr["倉庫名"].ToString() + " 入れ";
                    ohm.CustomerZip = FormatZipCode(sdr["支店郵便番号"].ToString());
                    ohm.CustomerAddress = sdr["支店住所1"].ToString() + sdr["支店住所2"].ToString();
                    ohm.CustomerTel = sdr["支店TEL"].ToString();
                    break;
                case 2:
                case 6:
                    Warehouse w = new Warehouse(_connectionString);
                    ohm.DeliveryTypeName = w.Name + " 入れ";
                    ohm.CustomerZip = FormatZipCode(w.Zip);
                    ohm.CustomerAddress = w.Address;
                    ohm.CustomerTel = w.Tel;
                    break;
                case 4:
                    ohm.DeliveryTypeName = sdr["倉庫名"].ToString() + " 入れ";
                    ohm.CustomerZip = "";
                    ohm.CustomerAddress = "";
                    ohm.CustomerTel = "";
                    break;
                case 5:
                    ohm.DeliveryTypeName = "直送";
                    ohm.CustomerZip = FormatZipCode(sdr["学校郵便番号"].ToString()); 
                    ohm.CustomerAddress = sdr["学校住所1"].ToString() + sdr["学校住所2"].ToString();
                    ohm.CustomerTel = sdr["学校TEL"].ToString();
                    break;
                case 7:
                case 8:
                    ohm.DeliveryTypeName = sdr["倉庫名"].ToString() + " 入れ";
                    ohm.CustomerZip = FormatZipCode(sdr["支店郵便番号"].ToString());
                    ohm.CustomerAddress = sdr["支店住所1"].ToString() + sdr["支店住所2"].ToString();
                    ohm.CustomerTel = sdr["支店TEL"].ToString();
                    break;
                default:
                    break;

            }

            ohm.OrderNo = "H" + sdr["発注伝票番号"].ToString();
            ohm.OrderDate = DateTime.ParseExact(sdr["発注伝票日付"].ToString(), "yyyyMMdd", null).ToString("yyyy年 M月 d日");
            ohm.PrivateNotes = sdr["発注社内伝票摘要"].ToString();
            ohm.PublicNotes = sdr["発注社外伝票摘要"].ToString();

            // 自社名以外の発注元情報（発注書の右上）
            // 受注と紐づく（学校コードが存在する）場合は学校の情報を
            // 紐づかない時は操作者マスタから印字
            CompanyInfo c = new CompanyInfo(_connectionString);
            ohm.SelfCompanyName = c.Name;
            if ((int)sdr["学校コード"] == 0)
            {
                ohm.SelfZipCode = FormatZipCode(sdr["操作者支店郵便番号"].ToString());
                ohm.SelfAddress = sdr["操作者支店住所1"].ToString() + sdr["操作者支店住所2"].ToString();
                ohm.SelfDepartmentName = sdr["操作者支店名称"].ToString();
                ohm.SelfTel = sdr["操作者支店TEL"].ToString();
                ohm.SelfFax = sdr["操作者支店FAX"].ToString();
            }
            else
            {
                ohm.SelfZipCode = FormatZipCode(sdr["学校担当者支店郵便番号"].ToString());
                ohm.SelfAddress = sdr["学校担当者支店住所1"].ToString() + sdr["学校担当者支店住所2"].ToString();
                ohm.SelfDepartmentName = sdr["学校担当者支店名称"].ToString();
                ohm.SelfTel = sdr["学校担当者支店TEL"].ToString();
                ohm.SelfFax = sdr["学校担当者支店FAX"].ToString();
            }

            ohm.ShippingDate = ((int)sdr["出荷日付"] == 0) ? "" : String.Format("納期指定：{0}", DateTime.ParseExact(sdr["出荷日付"].ToString(), "yyyyMMdd", null).ToString("yyyy/MM/dd"));

            // 右下の受注No
            // 納入区分が6の時は「即売No」
            // 発注→受注が紐づいてるときは「受注No」
            // 紐づかないときは「発注No」
            if ((Int16)sdr["納入区分"] == 6)
            {
                ohm.OrderNoTimeStamp = "即売No.SB" + sdr["受注表示番号"].ToString() + "-" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }
            else if (sdr["受注表示番号"].ToString() == "")
            {
                ohm.OrderNoTimeStamp = "No.H" + sdr["発注表示番号"].ToString() + "-" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }
            else
            {
                ohm.OrderNoTimeStamp = "受注No.J" + sdr["受注表示番号"].ToString() + "-" + DateTime.Now.ToString("yyyyMMddHHmmss");
            }
            ohm.FixedNotes = _fixedNotes.Trim();

        }

        private static string FormatZipCode(string input)
        {
            // 7桁の数字の場合
            if (input.Length == 7 && long.TryParse(input, out _))
            {
                return input.Substring(0, 3) + "-" + input.Substring(3, 4);
            }

            // その他の場合
            return input;
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
                System.IO.File.Delete(pathFrom);
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
