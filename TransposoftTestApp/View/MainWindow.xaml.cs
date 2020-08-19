using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TransposoftTestApp.ModelExcel;
using Excel = Microsoft.Office.Interop.Excel;

namespace TransposoftTestApp
{

    public partial class MainWindow : Window
    {
        private List<DataCargo> c1 = new List<DataCargo>(); //Для хранения 1 листа Груз после загрузки excel 
        private List<DataRate> r1 = new List<DataRate>(); //Для хранения 2 листа Тариф после загрузки excel 

        public MainWindow()
        {
            InitializeComponent();
        }

        //закрытие программы
        private void CloseProgram_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        //создание таблицы при загрузке Excel-файла
        private void GetStartTable(List<DataCargo> excelCargo)
        {
            //Очистка DataGrid
            dtTable.Columns.Clear();
            dtTable.Items.Clear();

            //Создание колонок
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Товар";
            c1.Binding = new Binding("NameCargo");
            dtTable.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Дата прихода на склад";
            c2.Binding = new Binding("ArrivalDate");
            dtTable.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Дата ухода со склада";
            c3.Binding = new Binding("DepartureDate");
            dtTable.Columns.Add(c3);

            //Добавление строки
            for (int i = 0; i < excelCargo.Count; i++)
            {
                dtTable.Items.Add(new {NameCargo = excelCargo[i].NameCargo, ArrivalDate = excelCargo[i].ArrivalDate.ToString("g"), DepartureDate = excelCargo[i].DepartureDate });
            }
        }

        //создание расчетной таблица
        private void GetCalculate(List<DataCalculate> calculates)
        {
            //Создание колонок
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Товар";
            c1.Binding = new Binding("NameCargo");
            dtTable.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Дата прихода на склад";
            c2.Binding = new Binding("ArrivalDate");
            dtTable.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Дата ухода со склада";
            c3.Binding = new Binding("DepartureDate");
            dtTable.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Начало расчета";
            c4.Binding = new Binding("StartCalculate");
            dtTable.Columns.Add(c4);
            DataGridTextColumn c5 = new DataGridTextColumn();
            c5.Header = "Окончание расчета";
            c5.Binding = new Binding("EndCalculate");
            dtTable.Columns.Add(c5);
            DataGridTextColumn c6 = new DataGridTextColumn();
            c6.Header = "Кол-во дней хранения";
            c6.Binding = new Binding("TimeStorage");
            dtTable.Columns.Add(c6);
            DataGridTextColumn c7 = new DataGridTextColumn();
            c7.Header = "Ставка";
            c7.Binding = new Binding("Price");
            dtTable.Columns.Add(c7);
            DataGridTextColumn c8 = new DataGridTextColumn();
            c8.Header = "Примечание";
            c8.Binding = new Binding("Info");
            dtTable.Columns.Add(c8);

            //Добавление строки
            for (int i = 0; i < calculates.Count; i++)
            {
                dtTable.Items.Add(new { 
                    NameCargo = calculates[i].NameCargo,
                    ArrivalDate = calculates[i].ArrivalDate.ToString("g"),
                    DepartureDate = calculates[i].DepartureDate,
                    StartCalculate = calculates[i].StartCalculate,
                    EndCalculate = calculates[i].EndCalculate,
                    TimeStorage = calculates[i].TimeStorage,
                    Price = calculates[i].Price,
                    Info = calculates[i].Info
                });
            }
        }

        //открытие excel-файла
        private void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Filter = "Файлы Excel (*.xlsx;*.xsl;*.xls;)|*.xlsx;*.xsl;*.xls;";

                if (dlg.ShowDialog() == true)
                {
                    List<DataCargo> cargos = ExcelReader.CreateCargo(dlg.FileName);
                    List<DataRate> rates = ExcelReader.CreateRate(dlg.FileName);
                    c1 = cargos;
                    r1 = rates;
                    GetStartTable(cargos);
                } 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        //расчет времени
        private void ButtStartCalc_Click(object sender, RoutedEventArgs e)
        {
            DateTime minCargo;

            if (c1.Count == 0)
            {
                MessageBox.Show("Excel-таблица не загружена!", "Ошика!", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (c1.Count != 0 && StartCalc.SelectedDate == null && EndCalc.SelectedDate == null)
            {
                GetStartTable(c1);
            } 
            else
            {

                if (StartCalc.SelectedDate == null || EndCalc.SelectedDate == null)
                {
                    MessageBox.Show("Не выбраны даты Начало расчета и/или Окончания расчета!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else if (StartCalc.SelectedDate > EndCalc.SelectedDate)
                {
                    MessageBox.Show(" \"Начало расчета\" не может быть БОЛЬШЕ \"Окончания расчета\"", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                } 
                else
                {
                    //Добавляем в minCargo - минимальную дату Прихода на склад
                    minCargo = c1.Min(a => a.ArrivalDate);

                    if (EndCalc.SelectedDate != null && minCargo.Date > EndCalc.SelectedDate)
                    {
                        MessageBox.Show("Товар в заданный период не сущестувает", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        //Очистка DataGrid
                        dtTable.Columns.Clear();
                        dtTable.Items.Clear();

                        //Определяем Начало расчета
                        DateTime dStart = StartCalc.SelectedDate.Value;
                        TimeSpan ts = new TimeSpan(00, 00, 00);
                        dStart = dStart.Date + ts;

                        //Определяем Конец расчета
                        DateTime dEnd = EndCalc.SelectedDate.Value;
                        ts = new TimeSpan(23, 59, 59);
                        dEnd = dEnd.Date + ts;

                        //Производим расчет
                        List<DataCalculate> calculates = new List<DataCalculate>();
                        for (int i = 0; i < c1.Count; i++)
                        {
                            if (Calculate.FinalTable(dStart, dEnd, c1[i], r1) != null)
                            {
                                calculates.AddRange(Calculate.FinalTable(dStart, dEnd, c1[i], r1));
                            }
                        }

                        //Добавляем в таблицу
                        GetCalculate(calculates);
                    }
                }
            }
        }
    }
}
