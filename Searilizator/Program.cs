using ExcelMacroAdd.Serializable.Entity;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.IO;

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
            var appSettings = new AppSettings(Resources(), ResourcesForm2(), ResourcesForm4(), GlobalDateBaseLocation());

            var serializer = new JsonSerializer();
            serializer.Converters.Add(new JavaScriptDateTimeConverter());
            serializer.NullValueHandling = NullValueHandling.Ignore;

            using (var sw = new StreamWriter("appSettings.json"))
            {
                using (JsonWriter writer = new JsonTextWriter(sw) { Formatting = Formatting.Indented })
                {
                    serializer.Serialize(writer, appSettings);
                }
            }

            Console.WriteLine(@"Файл json успешно создан!");
            Console.ReadKey();
        }

        private static string GlobalDateBaseLocation()
        {
            return "D:/Data/";
        }


        private static Resources Resources()
        {
            string nameFileJournal = "Книга1";
            int heightMaxBox = 1500;

            string templateWall = "Паспорт_навесные.docx";
            string templateFloor = "Паспорт_напольные.docx";

            return new Resources(nameFileJournal, heightMaxBox, templateWall, templateFloor);
        }

        private static ResourcesForm2 ResourcesForm2()
        {
            object[] circuitBreakerCurrent = { "1", "2", "3", "4", "5", "6", "8", "10", "13", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            object[] circuitBreakerCurve = { "B", "C", "D", "K", "L", "Z" };
            object[] maxCircuitBreakerCurrent = { "4,5", "6", "10", "15" };
            object[] amountOfPolesCircuitBreaker = { "1", "2", "3", "4", "1N", "3N" };
            object[] circuitBreakerVendor = { "IEK BA47", "IEK BA47М", "IEK Armat", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DKC", "DEKraft", "Schneider", "TDM" };
            //Массивы параметров выключателей нагрузки
            object[] loadSwitchCurrent = { "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            object[] amountOfPolesLoadSwitch = { "1", "2", "3", "4" };
            object[] loadSwitchVendor = { "IEK", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DEKraft", "Schneider", "TDM" };
            return new ResourcesForm2(circuitBreakerCurrent, circuitBreakerCurve,
                maxCircuitBreakerCurrent, amountOfPolesCircuitBreaker,
                circuitBreakerVendor, loadSwitchCurrent,
                amountOfPolesLoadSwitch, loadSwitchVendor);
        }

        private static ResourcesForm4 ResourcesForm4()
        {
            string[] transformerCurrent = {"5/5", "10/5", "15/5", "20/5", "25/5", "30/5", "40/5", "50/5", "60/5", "75/5", "80/5", "100/5",
                "120/5", "125/5", "150/5", "200/5", "250/5", "300/5", "400/5", "500/5", "600/5", "750/5", "800/5", "1000/5", "1200/5", "1250/5",
                "1500/5", "1600/5", "2000/5", "2250/5", "2500/5", "3000/5", "4000/5", "5000/5" };
            return new ResourcesForm4(transformerCurrent);
        }
    }
}
