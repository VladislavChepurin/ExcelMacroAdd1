﻿using System;
using System.Data;
using System.Net;

namespace ExcelMacroAdd
{
    internal class GetValuteTSB
    {
        public GetValuteTSB()
        {
            try
            {
                string url = "http://www.cbr.ru/scripts/XML_daily.asp";
                //XmlDocument xml_doc = new XmlDocument();
                //xml_doc.Load(url);
                DataSet ds = new DataSet();
                ds.ReadXml(url);
                DataTable currency = ds.Tables["Valute"];
                foreach (DataRow row in currency.Rows)
                {
                    //Поиск доллара
                    if (row["CharCode"].ToString() == "USD")
                    {
                        USDRate = Math.Round(Convert.ToDouble(row["Value"]), 2);
                            
                    }
                    // Поиск ЕВРО
                    if (row["CharCode"].ToString() == "EUR")//Ищу нужный код валюты
                    {
                        EvroRate = Math.Round(Convert.ToDouble(row["Value"]), 2);
                    }
                    // Поиск Юаня
                    if (row["CharCode"].ToString() == "CNY")//Ищу нужный код валюты
                    {
                        CnyRate = Math.Round(Convert.ToDouble(row["Value"]), 2);
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

        public double USDRate { get; }
        public double EvroRate { get; }
        public double CnyRate { get; }

    }
}
