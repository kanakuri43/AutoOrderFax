using System;
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
        private static string _reqUser;             // サービス UserID
        private static string _reqPassword;         // サービス Password
        private static string _mailAddress;         // 通知返信先アドレス
        private static string _retries;             // リトライ回数
        private static string _retryInterval;       // リトライ間隔（分）
        private static string _jikan;               // 時間指定（yyyymmddHHMM）
        private static string _quality;             // 
        private static string _paperSize;           // 
        private static string _direction;           // 
        private static string _fontSize;            // 
        private static string _fontType;            // 
        //private static string _reportFile;          // 
        private static string _logLifespan;         // 

        [STAThread]
        public static void Main(string[] args)
        {
            // 各種設定読み込み
            if (LoadSettings() == false) return;

            // ログファイル
            StreamWriter sw = new StreamWriter(_logFileDirectory + @"\" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + @".log", true, System.Text.Encoding.GetEncoding("Shift_JIS"));
            Console.SetOut(sw); // 出力先を設定
            Console.WriteLine(@"Output Path... " + _outputDirectory);
            Console.WriteLine(@"Output Mode... " + _outputMode);
            Console.WriteLine(@"FTP Server... " + _ftpHostName);

            try
            {
                ConnectDatabase();
                if (CreateFiles(GetOrderNoList()))
                {
                    //Process ps = new Process();
                    //ps.StartInfo.FileName = @"C:\okita\pdf\test.pdf";
                    //ps.Start();

                    // PDFファイルが作成されたら転送
                    CleanRemoteFiles();
                    TransferFiles();
                    DeleteLocalFiles();
                }
                DeleteLogFiles();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                Console.WriteLine(@"Bye.");
                sw.Dispose();
            }

        }

        private static bool LoadSettings()
        {
            string currentDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            XElement xml = XElement.Load(currentDirectory + "Config.xml");
            //_reportFile = xml.Element("ReportFile").Value.Trim();
            _logFileDirectory = xml.Element("LogFileDirectory").Value.Trim();
            _connectionString = xml.Element("ConnectionString").Value.Trim();
            _outputDirectory = xml.Element("OutputDirectory").Value.Trim();
            _outputMode = xml.Element("OutputMode").Value.Trim();
            _logLifespan = xml.Element("LogLifespan").Value.Trim();

            var ftp = (from f in xml.Elements("FTPServer") select f).First();
            _ftpHostName = ftp.Element("Host").Value;
            _ftpUser = ftp.Element("User").Value;
            _ftpPassword = ftp.Element("Password").Value;

            var req = (from f in xml.Elements("Request") select f).First();
            _reqUser = req.Element("User").Value;
            _reqPassword = req.Element("Password").Value;
            _mailAddress = req.Element("MailAddress").Value;
            _retries = req.Element("Retries").Value;
            _retryInterval = req.Element("RetryInterval").Value;
            _jikan = req.Element("Jikan").Value;
            _quality = req.Element("Quality").Value;
            _paperSize = req.Element("PaperSize").Value;
            _direction = req.Element("Direction").Value;
            _fontSize = req.Element("FontSize").Value;
            _fontType = req.Element("FontType").Value;

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
            Console.WriteLine(@"Connecting to " + Connection.DataSource + @"...");
            Connection.Open();
            Console.WriteLine(@"OK");
            Console.WriteLine(@"");

            return true;
        }

        private static List<string> GetOrderNoList()
        {
            var res = new List<string>();
            string sql = "SELECT "
                        + "  発注伝票番号 "
                        + "FROM "
                        + "  D発注FAX "
                        + "WHERE "
                        + "  送信Mode = 0 "
                        + "GROUP BY "
                        + "  発注伝票番号 ";
            using (SqlCommand command = new SqlCommand(sql, Connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        res = (from IDataRecord r in sdr select r["発注伝票番号"].ToString()).ToList();
                    }
                    Console.WriteLine(res.Count + @" Order(s) found.");
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

                Console.WriteLine(@"Order No. " + orderNo + @" created.");
                UpdateOrderState(orderNo);
            }
            Console.WriteLine(@"Create PDF, Request Files complete.");

            return true;
        }

        private static void CleanRemoteFiles()
        {
            // リモートのフォルダの中 全削除
            Console.WriteLine(@"Cleaning remote debris...");
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
            Console.WriteLine(@"OK");

        }

        private static bool TransferFiles()
        {
            string[] FileList;

            // ローカルにlock作成
            File.WriteAllText(_outputDirectory + @"\exclusive.lock", "");
            Console.WriteLine(@"exclusive.lock created.");

            using (WebClient wc = new WebClient())
            {
                wc.Credentials = new NetworkCredential(_ftpUser, _ftpPassword);

                // リモートにlock転送
                Console.WriteLine(@"Transfering exclusive.lock ...");
                wc.UploadFile(string.Format("{0}/{1}", _ftpHostName, "exclusive.lock"), _outputDirectory + @"\exclusive.lock");   // ファイルアップロード.
                Console.WriteLine(@"OK");

                // FTPファイル転送
                FileList = Directory.GetFiles(_outputDirectory, "*.pdf", System.IO.SearchOption.AllDirectories);
                if (FileList.Length > 0)
                {
                    Console.WriteLine(@"Transfer PDF Start.");
                    foreach (string FilePath in FileList)
                    {
                        string FileName = System.IO.Path.GetFileName(FilePath);                 // ファイル名取得.
                        wc.UploadFile(string.Format("{0}/{1}", _ftpHostName, FileName), FilePath);   // ファイルアップロード.

                        Console.WriteLine(FileName + @" transferd.");
                    }
                    Console.WriteLine(@"Transfer PDF complete.");
                }

                // Requestファイル転送
                FileList = Directory.GetFiles(_outputDirectory, "*.req", System.IO.SearchOption.AllDirectories);
                if (FileList.Length > 0)
                {
                    Console.WriteLine(@"Transfer Request Start.");
                    foreach (string FilePath in FileList)
                    {
                        string FileName = System.IO.Path.GetFileName(FilePath);                 // ファイル名取得.
                        wc.UploadFile(string.Format("{0}/{1}", _ftpHostName, FileName), FilePath);   // ファイルアップロード.

                        Console.WriteLine(FileName + @" transferd.");
                    }
                    Console.WriteLine(@"Transfer Request complete.");
                }

                // リモートのlock削除
                Console.WriteLine(@"Deleting remote exclusive.lock ...");
                using (var fc = new FtpClient(_ftpHostName, _ftpUser, _ftpPassword))
                {
                    fc.Connect();
                    if (fc.FileExists(@"exclusive.lock"))
                    {
                        fc.DeleteFile(@"exclusive.lock");
                        Console.WriteLine(@"OK");
                    }
                    else
                    {
                        Console.WriteLine(@"Not Found.");
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
                        // PDFファイル作成
                        OrderSlip orderSlip = new OrderSlip();
                        orderSlip.DataContext = ohm;
                        System.Windows.Documents.FixedPage fixedPage = new System.Windows.Documents.FixedPage();
                        fixedPage.Children.Add(orderSlip);

                        // A4よこ
                        fixedPage.Width = 11.69 * 96;
                        fixedPage.Height = 8.27 * 96;
                        PageContent pc = new PageContent();
                        ((IAddChild)pc).AddChild(fixedPage);
                        FixedDocument fixedDocument = new FixedDocument();
                        fixedDocument.Pages.Add(pc);

                        using (Package p = Package.Open(OutputDirectory + @"/" + OrderNo.ToString() + @".xps", FileMode.Create))

                        {
                            using (XpsDocument d = new XpsDocument(p))
                            {
                                XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(d);
                                writer.Write(fixedDocument.DocumentPaginator);
                            }
                        }

                        PdfSharp.Xps.XpsConverter.Convert(OutputDirectory + @"/" + OrderNo.ToString() + @".xps", OutputDirectory + @"/" + OrderNo.ToString() + @".pdf", 0);
                        File.Delete(OutputDirectory + @"/" + OrderNo.ToString() + @".xps");



                        // リクエストファイル作成
                        using (SqlCommand command = new SqlCommand(sql, Connection))
                        {
                            using (SqlDataReader sdr = command.ExecuteReader())
                            {
                                if (sdr.HasRows)
                                {
                                    while (sdr.Read())
                                    {
                                        StreamWriter sw = new StreamWriter(OutputDirectory + @"\" + OrderNo + ".req", false, Encoding.GetEncoding("Shift_JIS"));
                                        Supplier s = new Supplier(_connectionString, (int)sdr["仕入先コード"]);
                                        sw.WriteLine("FAXNO=" + s.FAX);
                                        sw.WriteLine("USERID=" + _reqUser);
                                        sw.WriteLine("PASSWORD=" + _reqPassword);
                                        sw.WriteLine("NAME=" + sdr["仕入先名"].ToString() + "_" + sdr["発注表示番号"] + "." + sdr["支店コード"]);
                                        sw.WriteLine("SCODE1=");
                                        sw.WriteLine("SCODE2=");
                                        sw.WriteLine("SCODE3=");
                                        sw.WriteLine("RETMAIL=" + _mailAddress);
                                        sw.WriteLine("RETRY=" + _retries);
                                        sw.WriteLine("RTRYINTERVAL=" + _retryInterval);
                                        sw.WriteLine("JIKAN=" + _jikan);
                                        sw.WriteLine("QUALITY=" + _quality);
                                        sw.WriteLine("PAPERSIZE=" + _paperSize);
                                        sw.WriteLine("DIRECTION=" + _direction);
                                        sw.WriteLine("FONTSIZE=" + _fontSize);
                                        sw.WriteLine("FONTTYPE=" + _fontType);
                                        sw.Close();
                                    }
                                }
                            }
                        }

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
            Console.WriteLine(@"Updating order state...");
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
            Console.WriteLine(@"OK");

        }

        private static void DeleteLocalFiles()
        {
            Console.WriteLine(@"Deleting local transfered files...");
            foreach (string pathFrom in System.IO.Directory.EnumerateFiles(_outputDirectory, "*.*"))
            {
                File.Delete(pathFrom);
            }
            Console.WriteLine(@"OK");
        }

        private static void DeleteLogFiles()
        {
            int count = 0;
            DirectoryInfo di = new DirectoryInfo(_logFileDirectory);
            foreach (FileInfo fi in di.GetFiles())
            {
                // 日付の比較
                if (fi.LastWriteTime < DateTime.Today.AddDays(int.Parse(_logLifespan) * -1))
                {
                    fi.Delete();
                    count++;
                }
            }
            Console.WriteLine(count + @" Log File(s) Deleted.");
        }

    }
}
