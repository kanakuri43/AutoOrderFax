using PdfSharp.Pdf.Content.Objects;
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
        public static string Encode(string DivideType, int[] ClassDivideArray)
        {
            if (DivideType == "0")
            {
                return "";
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                List<int> validElements = new List<int>();

                // 末尾から見て、0でない要素を見つける
                int lastIndex = ClassDivideArray.Length - 1;
                while (lastIndex >= 0 && ClassDivideArray[lastIndex] == 0)
                {
                    lastIndex--;
                }

                // 0でない要素をリストに追加
                for (int i = 0; i <= lastIndex; i++)
                {
                    validElements.Add(ClassDivideArray[i]);
                }

                // リストの要素をカンマで連結して文字列を作成
                sb.Append(string.Join(" ", validElements));
                return sb.ToString();

            }

        }


    }
}
