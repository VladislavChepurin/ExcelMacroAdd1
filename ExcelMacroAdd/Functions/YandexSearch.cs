using System.Diagnostics;

namespace ExcelMacroAdd.Functions
{
    internal sealed class YandexSearch : AbstractFunctions
    {
        public override void Start()
        {
            string value = Cell.Value;
            if (!string.IsNullOrEmpty(value))
            {
                string url = "http://www.yandex.ru/yandsearch?text=" + value;
                Process.Start(url);
            }
        }
    }
}
