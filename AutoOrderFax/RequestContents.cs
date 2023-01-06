using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOrderFax
{
    public class RequestContents
    {
        public string ReqUser { get; set; }
        public string ReqPassword { get; set; }
        public string MailAddress { get; set; }
        public string Retries { get; set; }
        public string RetryInterval { get; set; }
        public string Jikan { get; set; }
        public string Quality { get; set; }
        public string PaperSize { get; set; }
        public string Direction { get; set; }
        public string FontSize { get; set; }
        public string FontType { get; set; }

        public RequestContents()
        {

        }
    }
}
