using System.Diagnostics;

namespace ExcelMacroAdd.Functions
{
    internal sealed class GoogleSearch : AbstractFunctions
    {
        public override void Start()
        {
            string value = Cell.Value;
            if (!string.IsNullOrEmpty(value))
            {
                string url = "https://www.google.ru/search?q=" + value;
                Process.Start(url);
            }
        }
    }
}
