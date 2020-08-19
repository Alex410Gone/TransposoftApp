using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransposoftTestApp.ModelExcel
{
    class Calculate
    {
        public static List<DataCalculate> FinalTable(DateTime dStart, DateTime dEnd, DataCargo cargos, List<DataRate> rates)
        {

            DateTime dtArrival;
            DateTime dtDeparture;
            TimeSpan timeTotal;
            int daysInStock;
            int storageStartDay;
            int days = 0;
            DateTime endDeparture;

            List<DataCalculate> calculates = new List<DataCalculate>();

            //определяем значение для Начало рачета
            dtArrival = cargos.ArrivalDate;

            //определяем значение для Конец расчета
            if (cargos.DepartureDate == "" || DateTime.Parse(cargos.DepartureDate) > dEnd)
            {
                dtDeparture = dEnd;
            }
            else
            {
                dtDeparture = DateTime.Parse(cargos.DepartureDate);
            }

            //Кол-во дней от КонцаРасчета до ДатыПрихода
            timeTotal = dtDeparture.Date - dtArrival.Date;
            daysInStock = (int)timeTotal.TotalDays + 1;

            //Кол-во дней от НачалаРасчета до Даты прихода
            timeTotal = dStart.Date - dtArrival.Date;
            storageStartDay = (int)timeTotal.TotalDays + 1;

            //Производим расчет дней с конца
            for (int i = rates.Count - 1; i >= 0; i--)
            {
                //НачальныйДень = НачалоПериода или Кол-во дней от НачалаРасчета
                int startDay = (int)rates[i].StartPeriod > storageStartDay ? (int)rates[i].StartPeriod : storageStartDay;

                //Начальный день не может быть 0
                if (startDay < 1)
                {
                    startDay = 1;
                }

                //Если ВсегоДней >= НачальныйДень
                if (daysInStock >= startDay)
                {
                    int diffDays = daysInStock - startDay + 1;
                    daysInStock = daysInStock - diffDays;

                    //Переопределяем НачалоРасчета для Таблицы
                    days += diffDays;
                    dtArrival = dtDeparture - new TimeSpan(days - 1, 0, 0, 0);
                    TimeSpan ts_1 = new TimeSpan(00, 00, 00);
                    dtArrival = dtArrival.Date + ts_1;

                    if (dtArrival.Date == cargos.ArrivalDate.Date )
                    {
                        dtArrival = cargos.ArrivalDate;
                    }

                    //переопределяем КонецРасчета для Таблицы                    
                    endDeparture = dtArrival + new TimeSpan(diffDays - 1, 0, 0, 0);
                    TimeSpan ts_2 = new TimeSpan(23, 59, 59);
                    endDeparture = endDeparture.Date + ts_2;

                    if (cargos.DepartureDate != "" && endDeparture.Date.CompareTo(DateTime.Parse(cargos.DepartureDate).Date) >= 0)
                    {
                        endDeparture = DateTime.Parse(cargos.DepartureDate);
                    }

                    //Добавляем в таблицу
                    if (diffDays > 0)
                    {
                        //Добавляем вначало, если есть след Дата, то она встает на 0 место
                        calculates.Insert(0, new DataCalculate
                        {
                            NameCargo = cargos.NameCargo,
                            ArrivalDate = cargos.ArrivalDate,
                            DepartureDate = cargos.DepartureDate,
                            StartCalculate = dtArrival.ToString("g"),
                            EndCalculate = endDeparture.ToString("g"),
                            TimeStorage = diffDays,
                            Price = rates[i].Price,
                            Info = "Период №" + rates[i].Num
                        });
                    }
                }
            }

            return calculates;
        }
    }
}