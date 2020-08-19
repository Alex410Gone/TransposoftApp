using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TransposoftTestApp.ModelExcel
{

    class ExcelReader
    {
        //считываем первую страницу Excel - Груз
        public static List<DataCargo> CreateCargo(string filepath)
        {
            List<DataCargo> cargos = new List<DataCargo>();

            string xlConnStr = $"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={filepath};Extended Properties='Excel 8.0;HDR=Yes;';";
            var xlConn = new OleDbConnection(xlConnStr);
            OleDbCommand cmd = new OleDbCommand("SELECT [Груз], [Дата прихода на склад], [Дата ухода со склада] FROM [Груз$]", xlConn);

            DataTable dt = new DataTable();
            OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
            adapter.Fill(dt);

            foreach (DataRow row in dt.Rows)
            {
                var cells = row.ItemArray;

                if (cells[0].ToString() != "")
                {
                    cargos.Add(new DataCargo()
                    {
                        NameCargo = cells[0].ToString(),
                        ArrivalDate = DateTime.Parse(cells[1].ToString()),
                        DepartureDate = cells[2].ToString()
                    });
                }
                else
                {
                    break;
                }
            }

            return cargos;
        }

        //считываем вторую страницу Excel - Тариф
        public static List<DataRate> CreateRate(string filepath)
        {
            List<DataRate> rates = new List<DataRate>();

            string xlConnStr = $"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={filepath};Extended Properties='Excel 8.0;HDR=Yes;';";
            var xlConn = new OleDbConnection(xlConnStr);
            OleDbCommand cmd = new OleDbCommand("SELECT [№], [Начало периода], [Окончание периода], [Ставка] FROM [Тариф$]", xlConn);

            DataTable dt = new DataTable();
            OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
            adapter.Fill(dt);

            foreach (DataRow row in dt.Rows)
            {
                var cells = row.ItemArray;

                if (cells[0].ToString() != "")
                {
                    rates.Add(new DataRate()
                    {
                        Num = double.Parse(cells[0].ToString()),
                        StartPeriod = double.Parse(cells[1].ToString()),
                        EndPeriod = double.TryParse(cells[2].ToString(), out double num) ? num : 0,
                        Price = double.Parse(cells[3].ToString())
                    });
                }
                else
                {
                    break;
                }
            }

            return rates;
        }
    }
}
