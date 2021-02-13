using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Npgsql;
using ClosedXML.Excel;

namespace demo_dotnet_console_excel_writing
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            //reading configuration
            IConfiguration configuration = new ConfigurationBuilder().AddJsonFile("appsettings.json", true, true).Build();
            String connectionString = configuration.GetConnectionString("postgresql");

            Task<List<Province>> provinces = GetDataFromDb(connectionString);

            using (IXLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Province");
                int row = 1;
                provinces.Result.ForEach(province =>
                {
                    Console.WriteLine(province.id + " - " + province.name);
                    int column = 1;
                    ws.Cell(row, column++).Value = province.id;
                    ws.Cell(row++, column).Value = province.name;
                });
                wb.SaveAs("helo.xlsx");
            }
        }

        public static async Task<List<Province>> GetDataFromDb(String connectionString)
        {
            List<Province> provinces = new List<Province>();
            try
            {
                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    Console.WriteLine("Postgresql version = " + connection.ServerVersion);
                    String sql = "SELECT * FROM province";
                    using (NpgsqlCommand command = new NpgsqlCommand(sql, connection))
                    {
                        NpgsqlDataReader dataReader = await command.ExecuteReaderAsync();
                        while (dataReader.Read())
                        {
                            provinces.Add(new Province(Convert.ToInt32(dataReader[0]), Convert.ToString(dataReader[1])));
                        }
                    }
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return provinces;
        }
    }
}
