using ExcelMacroAdd.Serializable;
using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace Searilizator
{
    internal class Program
    {
        static void Main(string[] args)
        {           
            MainAsync().GetAwaiter().GetResult();
        }

        private static async Task MainAsync()
        {
            string[] circuitBreakerCurrent = new string[19] { "1", "2", "3", "4", "5", "6", "8", "10", "13", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            string[] circuitBreakerCurve = new string[6] { "B", "C", "D", "K", "L", "Z" };
            string[] maxCircuitBreakerCurrent = new string[4] { "4,5", "6", "10", "15" };
            string[] amountOfPolesCircuitBreaker = new string[6] { "1", "2", "3", "4", "1N", "3N" };
            string[] circuitBreakerVendor = new string[11] { "IEK ВА47", "IEK BA47М", "IEK Armat", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DKC", "DEKraft", "Schneider", "TDM" };
            //Массивы параметров выключателей нагрузки
            string[] loadSwitchCurrent = new string[10] { "16", "20", "25", "32", "40", "50", "63", "80", "100", "125" };
            string[] amountOfPolesLoadSwitch = new string[4] { "1", "2", "3", "4" };
            string[] loadSwitchVendor = new string[8] { "IEK", "EKF PROxima", "EKF AVERS", "KEAZ", "ABB", "DEKraft", "Schneider", "TDM" };
            StringResourcesForm2 stringResourcesForm2 = new StringResourcesForm2(circuitBreakerCurrent, circuitBreakerCurve,
                                                                                 maxCircuitBreakerCurrent, amountOfPolesCircuitBreaker,
                                                                                 circuitBreakerVendor, loadSwitchCurrent,
                                                                                 amountOfPolesLoadSwitch, loadSwitchVendor);

            string providerData = "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=";
            string nameFileDB = "BdMacro.accdb";
            string realseDirectoryDB = @"\\192.168.100.100\ftp\Info_A\FTP\Производство Абиэлт\Инженеры\";
            string debugDirectoryDB = @"Прайсы\Макро\";
            StringResourcesMainRibbon stringResourcesMainRibbon = new StringResourcesMainRibbon(providerData, nameFileDB, realseDirectoryDB, debugDirectoryDB);

            AppSettings appSettings = new AppSettings(stringResourcesForm2, stringResourcesMainRibbon);

            using (FileStream fs = new FileStream("appSettings.json", FileMode.OpenOrCreate))
            {
                


                await JsonSerializer.SerializeAsync<AppSettings>(fs, appSettings);
                Console.WriteLine("Data has been saved to file");
            }
            Console.ReadKey();
        }
    }
}
