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
            var appSettings = new AppSettings(Resources(), CorrectFontResources(), FormSettings(), GlobalDateBaseLocation(), TypeNkySettings());

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
            return "//192.168.100.100/ftp/Info_A/FTP/Производство Абиэлт/Инженеры/База данных/";
        }

        private static FillingOutThePassportSettings Resources()
        {
            string nameFileJournal = "Книга1";

            string templateWall = "Паспорт_навесные.docx";
            string templateFloor = "Паспорт_напольные.docx";
            string templateWallIt = "Паспорт_навесные IT.docx";
            string templateFloorIt = "Паспорт_напольные IT.docx";

            return new FillingOutThePassportSettings(nameFileJournal, templateWall, templateFloor, templateWallIt, templateFloorIt);
        }

        private static FormSettings FormSettings()
        {
            bool IsTopMost = true;
            return new FormSettings(IsTopMost);
        }

        private static CorrectFontSettings CorrectFontResources()
        {
            string nameFont = "Calibri";
            int sizeFont = 11;
            return new CorrectFontSettings(nameFont, sizeFont);
        }

        private static TypeNkySettings[] TypeNkySettings()
        {
            return null;
        }
    }
}
