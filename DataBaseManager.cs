using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelComparer
{
    public class DatabaseManager
    {

        private string GetConnectionString()
        {
            return Properties.Settings.Default.ConnectionString;
        }

        public SqlConnection CreateConnection()
        {
            return new SqlConnection(GetConnectionString());
        }
    }
}
