﻿using ExcelMacroAdd.Serializable.Entity;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Searilizator
{
    /// <summary>
    /// Вспомогательная программа для создания JSON файла
    /// Она не участвует в работе основной программы.
    /// </summary>
    internal class Program
    {  
        static void Main()
        {
            string[] circuitBreakerCurrent = new string[19] { "1", "2", "3", "4", "5", "6", "8", "10", "13", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            string[] circuitBreakerCurve = new string[6] { "B", "C", "D", "K", "L", "Z" };
            string[] maxCircuitBreakerCurrent = new string[4] { "4,5", "6", "10", "15" };
            string[] amountOfPolesCircuitBreaker = new string[6] { "1", "2", "3", "4", "1N", "3N" };
            string[] circuitBreakerVendor = new string[11] { "IEK BA47", "IEK BA47М", "IEK Armat", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DKC", "DEKraft", "Schneider", "TDM" };
            //Массивы параметров выключателей нагрузки
            string[] loadSwitchCurrent = new string[10] { "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            string[] amountOfPolesLoadSwitch = new string[4] { "1", "2", "3", "4" };
            string[] loadSwitchVendor = new string[8] { "IEK", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DEKraft", "Schneider", "TDM" };
            ResourcesForm2 resourcesForm2 = new ResourcesForm2(circuitBreakerCurrent, circuitBreakerCurve,
                                                                                 maxCircuitBreakerCurrent, amountOfPolesCircuitBreaker,
                                                                                 circuitBreakerVendor, loadSwitchCurrent,
                                                                                 amountOfPolesLoadSwitch, loadSwitchVendor);


            string nameFileJornal = "Книга1";
            int heihgtMaxBox = 1500;

            string templeteWall = "Паспорт_навесные.docx";
            string templeteFloor = "Паспорт_напольные.docx";
      
            Resources resourcesForm1 = new Resources(nameFileJornal, heihgtMaxBox, templeteWall, templeteFloor);

            AppSettings appSettings = new AppSettings( resourcesForm1, resourcesForm2);

            JsonSerializer serializer = new JsonSerializer();
            serializer.Converters.Add(new JavaScriptDateTimeConverter());
            serializer.NullValueHandling = NullValueHandling.Ignore;

            using (StreamWriter sw = new StreamWriter("appSettings.json"))
            {
                using (JsonWriter writer = new JsonTextWriter(sw) { Formatting = Formatting.Indented })
                {
                    serializer.Serialize(writer, appSettings);
                }
            }              

            Console.ReadKey();
        }       
    }
}
