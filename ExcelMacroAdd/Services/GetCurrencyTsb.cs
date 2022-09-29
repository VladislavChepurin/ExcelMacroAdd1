using System;
using System.Data;
using System.Net;
using System.Threading;

namespace ExcelMacroAdd.Services
{
    /// <summary>
    /// Класс запроса курса валют
    /// </summary>
    public class GetCurrencyTsb
    {
        public delegate void Currency(double usdCurrency, double euroCurrency, double cnhCurrency);

        public Currency CurrencyHandler { get; set; }

        public void Start()
        {
            while (true)
            {
                try
                {
                    double usdPrice = default,
                           euroPrice = default,
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
                            euroPrice = Math.Round(Convert.ToDouble(row["Value"]) / nominal, 2);
                        }
                        // Поиск Юаня
                        if (row["CharCode"].ToString() == "CNY")
                        {
                            int nominal = Convert.ToInt32(row["Nominal"]);
                            cnyPrice = Math.Round(Convert.ToDouble(row["Value"]) / nominal, 2);
                        }
                    }
                    CurrencyHandler(usdPrice, euroPrice, cnyPrice);
                }
                catch (WebException)
                {
                    CurrencyHandler(0.0, 0.0, 0.0);
                }
                Thread.Sleep(60000);
            }       
        }    
    }
}
