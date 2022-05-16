using System;
using System.Data;
using System.Net;

namespace ExcelMacroAdd
{
    /// <summary>
    /// Класс запроса курса валют
    /// </summary>
    internal class GetValuteTSB
    {
        public GetValuteTSB()
        {
            try 
            {
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
                        USDRate = Math.Round(Convert.ToDouble(row["Value"])/nominal, 2);
                    }

                    // Поиск ЕВРО
                    if (row["CharCode"].ToString() == "EUR")
                    {
                        int nominal = Convert.ToInt32(row["Nominal"]);
                        EvroRate = Math.Round(Convert.ToDouble(row["Value"])/nominal, 2);
                    }
                    // Поиск Юаня
                    if (row["CharCode"].ToString() == "CNY")
                    {
                        int nominal = Convert.ToInt32(row["Nominal"]);
                        CnyRate = Math.Round(Convert.ToDouble(row["Value"])/nominal, 2);
                    }
                }
            }
            catch (WebException)
            {
                USDRate = 0;
                EvroRate = 0;
                CnyRate = 0;
            }
           
        }

        public double USDRate { get; private set; }
        public double EvroRate { get; private set; }
        public double CnyRate { get; private set; }

    }
}
