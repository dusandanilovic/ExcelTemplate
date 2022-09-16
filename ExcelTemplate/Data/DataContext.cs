using MySql.Data.MySqlClient;
using System.Data;
using Syncfusion.XlsIO;

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

        public void GenerateExcel()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Open an existing spreadsheet, which will be used as a template for generating the new spreadsheet.
                //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
                IWorkbook workbook;

                //Open existing Excel template
                var cfFileStream = new FileStream(@"C:\xls\XlsTemplates\CR_SATemplate.xlsx", FileMode.Open, FileAccess.Read);
                workbook = excelEngine.Excel.Workbooks.Open(cfFileStream);

                //The first worksheet in the workbook is accessed.
                IWorksheet worksheet = workbook.Worksheets[0];

                ////Create Template Marker processor.
                ////Apply the marker to export data from datatable to worksheet.
                //ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();
                //marker.AddVariable("SalesList", northwindDt);
                //marker.ApplyMarkers();

                //Saving and closing the workbook
                FileStream fileStream = new FileStream(@"C:\xls\Output.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                workbook.SaveAs(fileStream);
                //Close the workbook
                workbook.Close();
            }
        }
    }
}
