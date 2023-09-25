using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO.Packaging;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Xps;
using System.Windows;
using System.Windows.Xps.Packaging;

namespace AutoOrderFax
{
    public class PdfCreator
    {
        public RequestContents RequestContents { get; set; }
        public string sql { get; set; }
        public string ConnectionString { get; set; }
        public List<OrderHeaderModel> ohms { get; set; }

        private string _orderNo;

        public PdfCreator(string OrderNo)
        {
            _orderNo = OrderNo;
        }

        public void Create(string OutputDirectory)
        {

            FixedDocument fixedDocument = new FixedDocument();

            foreach(var ohm in ohms) 
            {

                OrderSlip orderSlip = new OrderSlip();
                orderSlip.DataContext = ohm;

                FixedPage fixedPage = new FixedPage();
                fixedPage.Children.Add(orderSlip);
                // A4よこ
                fixedPage.Width = 11.69 * 96;
                fixedPage.Height = 8.27 * 96;
                PageContent pc = new PageContent();
                ((IAddChild)pc).AddChild(fixedPage);

                fixedDocument.Pages.Add(pc);
            }

            using (Package p = Package.Open(OutputDirectory + @"/" + _orderNo.ToString() + @".xps", FileMode.Create))

            {
                using (XpsDocument d = new XpsDocument(p))
                {
                    XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(d);
                    writer.Write(fixedDocument.DocumentPaginator);
                }
            }

            PdfSharp.Xps.XpsConverter.Convert(OutputDirectory + @"/" + _orderNo.ToString() + @".xps", OutputDirectory + @"/" + _orderNo.ToString() + @".pdf", 0);
            File.Delete(OutputDirectory + @"/" + _orderNo.ToString() + @".xps");



            // リクエストファイル作成
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        while (sdr.Read())
                        {
                            StreamWriter sw = new StreamWriter(OutputDirectory + @"\" + _orderNo + ".req", false, Encoding.GetEncoding("Shift_JIS"));
                            Supplier s = new Supplier(ConnectionString, (int)sdr["仕入先コード"]);
                            sw.WriteLine("FAXNO=" + s.FAX);
                            sw.WriteLine("USERID=" + RequestContents.ReqUser);
                            sw.WriteLine("PASSWORD=" + RequestContents.ReqPassword);
                            sw.WriteLine("NAME=" + sdr["仕入先名"].ToString() + "_" + sdr["発注表示番号"] + "." + sdr["支店コード"]);
                            sw.WriteLine("SCODE1=");
                            sw.WriteLine("SCODE2=");
                            sw.WriteLine("SCODE3=");
                            sw.WriteLine("RETMAIL=" + RequestContents.MailAddress);
                            sw.WriteLine("RETRY=" + RequestContents.Retries);
                            sw.WriteLine("RTRYINTERVAL=" + RequestContents.RetryInterval);
                            sw.WriteLine("JIKAN=" + RequestContents.Jikan);
                            sw.WriteLine("QUALITY=" + RequestContents.Quality);
                            sw.WriteLine("PAPERSIZE=" + RequestContents.PaperSize);
                            sw.WriteLine("DIRECTION=" + RequestContents.Direction);
                            sw.WriteLine("FONTSIZE=" + RequestContents.FontSize);
                            sw.WriteLine("FONTTYPE=" + RequestContents.FontType);
                            sw.Close();
                        }
                    }
                }
            }
        }
    }
}
