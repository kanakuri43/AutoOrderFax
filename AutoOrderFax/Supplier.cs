using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOrderFax
{
    public class Supplier
    {
        public string Name { get; }
        public string FAX { get; }

        public Supplier(string ConnectionString, int SupplierCode)
        {
            SqlConnection connection = new SqlConnection(ConnectionString);
            connection.Open();
            string sql = "SELECT "
                        + "  仕入先名 "
                        + "  , FAX "
                        + " FROM "
                        + "   M仕入先 "
                        + " WHERE "
                        + "   削除区分 = 0 "
                        + "   AND 仕入先コード =" + SupplierCode.ToString()
                        ;
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    if (sdr.HasRows)
                    {
                        if (sdr.Read())
                        {
                            Name = sdr["仕入先名"].ToString();
                            FAX = sdr["FAX"].ToString();
                        }
                    }
                }
            }

        }
    }
}
