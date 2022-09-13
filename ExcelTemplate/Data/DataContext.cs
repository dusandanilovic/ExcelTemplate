using MySql.Data.MySqlClient;
using System.Data;

namespace ExcelTemplate.Data
{
    public class DataContext
    {
        public MySqlConnection connection { get; set; }
        public DataContext(string connectionString)
        {
            connection = new MySqlConnection(connectionString);
        }

        public DataSet ExecuteProcedure(string name, string reportDate, string reportType)
        {
            var result = new DataSet();
            try
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = name;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 20000;

                command.Parameters.Add(new MySqlParameter { ParameterName = "in_report_date", Direction = ParameterDirection.Input, Value = reportDate, MySqlDbType = MySqlDbType.Date });
                command.Parameters.Add(new MySqlParameter { ParameterName = "in_report_type_code", Direction = ParameterDirection.Input, Value = reportType, MySqlDbType = MySqlDbType.VarChar });

                //command.Parameters = parameters;
                var reader = command.ExecuteReader();

                DataSet ds = new DataSet();

                while (!reader.IsClosed)
                    ds.Tables.Add().Load(reader);

                return ds;

            }
            catch (Exception e)
            {

            }
            finally
            {
                connection.Close();
            }

            return result;
        }
    }
}
