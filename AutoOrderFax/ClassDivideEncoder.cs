using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOrderFax
{
    public static class ClassDivideEncoder
    {
        public static string Encode(string DivideType, string SourceText)
        {
            if (DivideType == "0")
            {
                return "";
            }
            else
            {
                string EncodedText = SourceText.TrimEnd('0');
                if (EncodedText.Length > 0)
                {
                    return EncodedText;
                }
                else
                {
                    return "0";
                }

            }

        }


    }
}
