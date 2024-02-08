using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOrderFax
{
    public class OrderInfo
    {
        public int LineCount { get; }
        public string ConnectionString { get; set; }


        public OrderInfo(string ConnectionString, string OrderNo)
        {
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();
            string sql = "SELECT "
                        + "  MAX(発注行番号)  行数"
                        + " FROM "
                        + "   D発注 "
                        + " WHERE "
                        + "   削除区分 = 0 "
                        + "   AND 発注伝票番号 = " + OrderNo
                        ;
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        if (sdr.Read())
                        {
                            LineCount = (Int16)sdr["行数"];
                        }
                    }
                }
            }
        }
    }
}
