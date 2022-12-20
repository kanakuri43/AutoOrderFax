using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOrderFax
{
    public class Warehouse
    {
        public string Name { get; }
        public string Address { get; }
        public string Tel { get; }

        public Warehouse(string ConnectionString)
        {
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();
            string sql = "SELECT "
                        + "  倉庫名 "
                        + "  , 住所1 + 住所2 住所"
                        + "  , TEL "
                        + " FROM "
                        + "   M支店 "
                        + "     LEFT OUTER JOIN M倉庫 "
                        + "     ON (M支店.倉庫コード = M倉庫.倉庫コード) "
                        + " WHERE "
                        + "   M支店.支店コード = 1 ";
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        if (sdr.Read())
                        {
                            Name = sdr["倉庫名"].ToString();
                            Address = sdr["住所"].ToString();
                            Tel = sdr["TEL"].ToString();
                        }
                    }
                }
            }

        }
    }
}
