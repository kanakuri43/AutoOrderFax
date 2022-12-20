using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOrderFax
{
    public class CompanyInfo
    {
        public string Name { get; }

        public CompanyInfo(string ConnectionString)
        {
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();
            string sql = "SELECT "
                        + "  社名 "
                        + " FROM "
                        + "   Mユーザー情報 ";
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        if (sdr.Read())
                        {
                            Name = sdr["社名"].ToString();
                        }
                    }
                }
            }

        }
    }
}
