using System;
using System.Data;
using System.Net;
using System.Threading;

namespace ExcelMacroAdd.Servises
{
    /// <summary>
    /// Класс запроса курса валют
    /// </summary>
    public class GetValuteTSB
    {
        public delegate void ValuteUSD(double usdValute, double evroValute, double cnhValute);

        public ValuteUSD ValuteUSDHandler { get; set; }

        public void Start()
        {
            while (true)
            {
                try
                {
                    double usdPrice = default,
                           evroPrice = default,
                           cnyPrice = default;

                    string url = "http://www.cbr.ru/scripts/XML_daily.asp";
                    DataSet ds = new DataSet();
                    ds.ReadXml(url);
                    DataTable currency = ds.Tables["Valute"];
                    foreach (DataRow row in currency.Rows)
                    {

                        //Поиск доллара
                        if (row["CharCode"].ToString() == "USD")
                        {
                            int nominal = Convert.ToInt32(row["Nominal"]);
                            usdPrice = Math.Round(Convert.ToDouble(row["Value"]) / nominal, 2);
                        }

                        // Поиск ЕВРО
                        if (row["CharCode"].ToString() == "EUR")
                        {
                            int nominal = Convert.ToInt32(row["Nominal"]);
                            evroPrice = Math.Round(Convert.ToDouble(row["Value"]) / nominal, 2);
                        }
                        // Поиск Юаня
                        if (row["CharCode"].ToString() == "CNY")
                        {
                            int nominal = Convert.ToInt32(row["Nominal"]);
                            cnyPrice = Math.Round(Convert.ToDouble(row["Value"]) / nominal, 2);
                        }
                    }
                    ValuteUSDHandler(usdPrice, evroPrice, cnyPrice);
                }
                catch (WebException)
                {
                    ValuteUSDHandler(0.0, 0.0, 0.0);
                }
                Thread.Sleep(30000);
            }       
        }    
    }
}
