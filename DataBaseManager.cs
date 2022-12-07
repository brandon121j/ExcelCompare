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
